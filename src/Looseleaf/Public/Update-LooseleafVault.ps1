function Update-LooseleafVault {
    <#
    .SYNOPSIS
        Stage 2: builds the Obsidian vault from the stage 1 work folder (manifest.json + .docx).

    .DESCRIPTION
        For every configured notebook in <WorkPath>\manifest.json:
          1. converts each page's .docx to Markdown (pandoc + filters/looseleaf.lua),
          2. plans folders and file names from the OneNote hierarchy and order,
          3. names images (media registry), resolves links, adds front matter and attachments,
          4. writes notes and media, protecting notes edited after the last build,
          5. removes output of pages that no longer exist, and saves .looseleaf\build-state.json.
        Does not need OneNote. Supports -WhatIf. See docs/DESIGN.md section 5.

    .PARAMETER ConfigPath
        Looseleaf configuration file (.psd1).

    .PARAMETER StopOnError
        Stop at the first page that fails instead of continuing.

    .OUTPUTS
        One summary object per notebook. Failed pages are listed in FailedPages.

    .EXAMPLE
        Update-LooseleafVault -ConfigPath .\looseleaf.config.psd1 -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [string]$ConfigPath,

        [switch]$StopOnError
    )

    $config = Import-LLConfig -Path $ConfigPath
    $manifest = Import-LLManifest -Path (Join-Path $config.WorkPath 'manifest.json')
    $filterPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'filters\looseleaf.lua'
    if (-not (Test-Path -LiteralPath $filterPath)) { throw "Lua filter not found: $filterPath" }

    $logPath = $null
    if (-not $WhatIfPreference) {
        $logDir = Join-Path $config.WorkPath 'logs'
        [void][IO.Directory]::CreateDirectory($logDir)
        $logPath = Join-Path $logDir ('{0}-build.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    $log = {
        param([string]$Level, [string]$Message)
        if ($logPath) {
            [IO.File]::AppendAllText($logPath, ('{0} {1,-5} {2}{3}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message, [Environment]::NewLine), (New-Object System.Text.UTF8Encoding $false))
        }
        if ($Level -eq 'WARN') { Write-Warning $Message } else { Write-Verbose $Message }
    }

    $wanted = @($config.Notebooks)
    $notebooks = @(Get-LLValue $manifest 'notebooks' @())
    foreach ($name in $wanted) {
        if (-not @($notebooks | Where-Object { (Get-LLValue $_ 'name' '') -eq $name })) {
            & $log 'WARN' "Notebook '$name' is not in the manifest. Run Export-LooseleafNotebook for it first."
        }
    }

    foreach ($notebook in @($notebooks | Where-Object { $wanted -contains (Get-LLValue $_ 'name' '') })) {
        $notebookName = [string](Get-LLValue $notebook 'name' '')
        $notebookDir = Join-Path $config.VaultPath (Get-LLSafeName -Name $notebookName -NameSpaces $config.NameSpaces -MaxLength $config.MaxNameLength)
        $statePath = Join-Path $notebookDir '.looseleaf\build-state.json'
        $oldState = Read-LLBuildState -Path $statePath
        $newState = @{ Files = @{}; Media = @{} }
        & $log 'INFO' "Building notebook '$notebookName' -> $notebookDir"

        $summary = [ordered]@{
            Notebook        = $notebookName
            Path            = $notebookDir
            Pages           = 0
            Created         = 0
            Updated         = 0
            Unchanged       = 0
            SkippedModified = 0
            SkippedForeign  = 0
            BackedUp        = 0
            EmptyPages      = 0
            Failed          = 0
            Removed         = 0
            Images          = 0
            Warnings        = 0
            FailedPages     = New-Object System.Collections.Generic.List[object]
            LogPath         = $logPath
        }

        # --- 1. Convert every page (needs no names yet, only image prefixes) ---------------
        $pages = New-Object System.Collections.Generic.List[object]
        Get-LLNotebookPage -Items @(Get-LLValue $notebook 'items' @()) -SectionName '' -Into $pages
        $summary.Pages = $pages.Count

        # Inside WorkPath (not %TEMP%), so one antivirus exclusion covers all of Looseleaf's files.
        $tempRoot = Join-Path $config.WorkPath ('tmp\' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void][IO.Directory]::CreateDirectory($tempRoot)
        $results = @{}
        $emptyKeys = @{}
        $failedKeys = @{}
        try {
            $i = 0
            foreach ($p in $pages) {
                $i++
                $page = $p.Page
                $key = [string](Get-LLValue $page 'key' '')
                $title = [string](Get-LLValue $page 'name' '')
                Write-Progress -Activity "Looseleaf: converting '$notebookName'" -Status $title -PercentComplete ([int](100 * $i / [Math]::Max(1, $pages.Count)))
                $status = [string](Get-LLValue $page 'status' 'exported')
                $docx = [string](Get-LLValue $page 'docx' '')
                $docxPath = if ($docx) { Join-Path $config.WorkPath $docx } else { $null }
                try {
                    if (-not $key) { throw 'Page has no key in the manifest.' }
                    if ($status -eq 'failed' -or $status -eq 'locked') { throw "Page was not exported in stage 1 (status '$status')." }
                    if (-not $docxPath -or -not (Test-Path -LiteralPath $docxPath -PathType Leaf)) { throw "Word file missing: $docxPath" }

                    $prefix = if ($config.ImageNameStyle -eq 'SectionPage') {
                        '{0}-{1}' -f (Get-LLImageLabel $p.SectionName -MaxLength $config.MaxImageLabelLength), (Get-LLImageLabel $title -MaxLength $config.MaxImageLabelLength)
                    } else {
                        Get-LLImageLabel $title -MaxLength $config.MaxImageLabelLength
                    }
                    $maxPrefix = $config.MaxImageNameLength - 10 # room for '-NN-N.jpeg'
                    if ($prefix.Length -gt $maxPrefix) { $prefix = $prefix.Substring(0, $maxPrefix).TrimEnd('-', '.') }

                    $pageTemp = Join-Path $tempRoot ('{0:D5}' -f $i)
                    $converted = Convert-LLPage -DocxPath $docxPath -TempDir $pageTemp -ImagePrefix $prefix -Config $config -FilterPath $filterPath
                    if ($converted.Warnings) {
                        & $log 'WARN' "pandoc warnings for '$title': $($converted.Warnings)"
                        $summary.Warnings++
                    }
                    $results[$key] = $converted
                    if ($converted.IsEmpty) { $emptyKeys[$key] = $true }
                } catch {
                    $failedKeys[$key] = $true
                    $summary.Failed++
                    $summary.FailedPages.Add([pscustomobject]@{ Title = $title; Key = $key; Error = $_.Exception.Message })
                    & $log 'WARN' "Page '$title' failed: $($_.Exception.Message)"
                    if ($StopOnError) { throw }
                }
            }
            Write-Progress -Activity "Looseleaf: converting '$notebookName'" -Completed
            $summary.EmptyPages = $emptyKeys.Count

            # --- 2. Plan folders and names -----------------------------------------------
            $plan = Get-LLPagePlan -Notebook $notebook -Config $config -NotebookDir $notebookDir -EmptyKeys $emptyKeys
            foreach ($w in $plan.Warnings) { & $log 'WARN' $w; $summary.Warnings++ }

            # Media names already owned (stable across rebuilds).
            $owned = @{}
            foreach ($k in $oldState.Media.Keys) { $owned[$oldState.Media[$k].ToLowerInvariant()] = $k }

            # --- 3./4. Assemble and write notes ------------------------------------------
            foreach ($entry in $plan.Entries) {
                $key = $entry.Key
                if ($failedKeys.ContainsKey($key) -or -not $results.ContainsKey($key)) { continue }
                if (-not $entry.FileName) { continue } # empty parent page: folder only

                $converted = $results[$key]
                $noteRel = if ($entry.RelDir) { Join-Path $entry.RelDir $entry.FileName } else { $entry.FileName }
                $depth = if ($entry.RelDir) { @($entry.RelDir -split '\\').Count } else { 0 }
                $mediaRelDir = if ($config.MediaScope -eq 'Folder') {
                    if ($entry.RelDir) { Join-Path $entry.RelDir 'media' } else { 'media' }
                } else { 'media' }
                $linkPrefix = if ($config.MediaScope -eq 'Folder') { 'media/' } else { ('../' * $depth) + 'media/' }

                try {
                    # Images: registry name, write, then replace placeholders.
                    $imageNames = @{}
                    foreach ($image in $converted.Images) {
                        $index = if ($image.BaseName -match '-(\d+)$') { $Matches[1] } else { $image.BaseName }
                        $fileName = Resolve-LLMediaName -RegistryKey "$key|img|$index" -Desired $image.Name -OldMedia $oldState.Media -NewMedia $newState.Media -Owned $owned
                        $bytes = [IO.File]::ReadAllBytes($image.FullName)
                        $r = Write-LLManagedFile -Root $notebookDir -RelPath (Join-Path $mediaRelDir $fileName) -Bytes $bytes -PageKey $key -OldFiles $oldState.Files -NewFiles $newState.Files -Mode 'Overwrite' -Cmdlet $PSCmdlet
                        if ($r -ne 'WhatIf') { $summary.Images++ }
                        $imageNames[$image.Name] = $fileName
                    }
                    $body = Resolve-LLMediaPlaceholder -Markdown $converted.Markdown -Names $imageNames -LinkPrefix $linkPrefix -LinkStyle $config.LinkStyle

                    # Attachments copied by stage 1.
                    $attachmentLinks = New-Object System.Collections.Generic.List[string]
                    foreach ($att in @(Get-LLValue $entry.Page 'attachments' @())) {
                        $attName = [string](Get-LLValue $att 'name' '')
                        $attFile = [string](Get-LLValue $att 'file' '')
                        $src = if ($attFile) { Join-Path $config.WorkPath $attFile } else { '' }
                        if (-not $src -or -not (Test-Path -LiteralPath $src -PathType Leaf)) {
                            & $log 'WARN' "Attachment '$attName' of '$($entry.Title)' is missing in the work folder."
                            $summary.Warnings++
                            continue
                        }
                        $desired = Get-LLSafeName -Name $attName -NameSpaces Hyphen -MaxLength $config.MaxNameLength
                        $fileName = Resolve-LLMediaName -RegistryKey "$key|att|$attName" -Desired $desired -OldMedia $oldState.Media -NewMedia $newState.Media -Owned $owned
                        [void](Write-LLManagedFile -Root $notebookDir -RelPath (Join-Path $mediaRelDir $fileName) -Bytes ([IO.File]::ReadAllBytes($src)) -PageKey $key -OldFiles $oldState.Files -NewFiles $newState.Files -Mode 'Overwrite' -Cmdlet $PSCmdlet)
                        $attachmentLinks.Add((Format-LLMediaLink -FileName $fileName -Text $attName -LinkPrefix $linkPrefix -LinkStyle $config.LinkStyle))
                    }

                    # PDF of a page with ink (ExportInkPagesAsPdf), linked like an attachment.
                    $pdf = [string](Get-LLValue $entry.Page 'pdf' '')
                    if ($pdf) {
                        $pdfSrc = Join-Path $config.WorkPath $pdf
                        if (Test-Path -LiteralPath $pdfSrc -PathType Leaf) {
                            $desired = (Get-LLImageLabel $entry.Title -MaxLength $config.MaxImageLabelLength) + '-ink.pdf'
                            $fileName = Resolve-LLMediaName -RegistryKey "$key|pdf" -Desired $desired -OldMedia $oldState.Media -NewMedia $newState.Media -Owned $owned
                            [void](Write-LLManagedFile -Root $notebookDir -RelPath (Join-Path $mediaRelDir $fileName) -Bytes ([IO.File]::ReadAllBytes($pdfSrc)) -PageKey $key -OldFiles $oldState.Files -NewFiles $newState.Files -Mode 'Overwrite' -Cmdlet $PSCmdlet)
                            $attachmentLinks.Add((Format-LLMediaLink -FileName $fileName -Text 'Original page with handwriting (PDF)' -LinkPrefix $linkPrefix -LinkStyle $config.LinkStyle))
                        }
                    }

                    $frontMatter = [ordered]@{
                        title        = $entry.Title
                        created      = [string](Get-LLValue $entry.Page 'created' '')
                        modified     = [string](Get-LLValue $entry.Page 'modified' '')
                        'onenote-id' = [string](Get-LLValue $entry.Page 'id' '')
                        source       = $entry.SourcePath
                    }
                    $note = Format-LLNote -Title $entry.Title -Body $body -FrontMatter $frontMatter -AttachmentLinks $attachmentLinks.ToArray() -Config $config
                    $bytes = (New-Object System.Text.UTF8Encoding $false).GetBytes($note)
                    $result = Write-LLManagedFile -Root $notebookDir -RelPath $noteRel -Bytes $bytes -PageKey $key -OldFiles $oldState.Files -NewFiles $newState.Files -Mode $config.OnModifiedFile -Cmdlet $PSCmdlet
                    switch ($result) {
                        'Created' { $summary.Created++ }
                        'Updated' { $summary.Updated++ }
                        'BackedUp' { $summary.BackedUp++; $summary.Updated++ }
                        'Unchanged' { $summary.Unchanged++ }
                        'SkippedModified' { $summary.SkippedModified++; & $log 'WARN' "Kept your edited note: $noteRel" }
                        'SkippedForeign' { $summary.SkippedForeign++; & $log 'WARN' "Not overwritten (file was not created by Looseleaf): $noteRel" }
                    }
                } catch {
                    $failedKeys[$key] = $true
                    $summary.Failed++
                    $summary.FailedPages.Add([pscustomobject]@{ Title = $entry.Title; Key = $key; Error = $_.Exception.Message })
                    & $log 'WARN' "Page '$($entry.Title)' failed while writing: $($_.Exception.Message)"
                    if ($StopOnError) { throw }
                }
            }

            # --- 5. Keep output of failed pages and locked sections; remove stale files ----
            foreach ($k in @($oldState.Files.Keys)) {
                $entryState = $oldState.Files[$k]
                $keepForPage = $failedKeys.ContainsKey($entryState.Page)
                $keepForSection = @($plan.KeepPrefixes | Where-Object { $k.StartsWith(($_.Replace('\', '/') + '/'), [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
                if (($keepForPage -or $keepForSection) -and -not $newState.Files.ContainsKey($k)) {
                    $newState.Files[$k] = $entryState
                }
            }
            foreach ($k in @($oldState.Media.Keys)) {
                $pageKey = ($k -split '\|')[0]
                if ($failedKeys.ContainsKey($pageKey) -and -not $newState.Media.ContainsKey($k)) { $newState.Media[$k] = $oldState.Media[$k] }
            }
            $summary.Removed = Remove-LLStaleFile -Root $notebookDir -OldFiles $oldState.Files -NewFiles $newState.Files -Cmdlet $PSCmdlet -Log $log

            if (-not $WhatIfPreference) { Save-LLBuildState -Path $statePath -State $newState }
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false
        }

        & $log 'INFO' ("Done: {0} pages, {1} created, {2} updated, {3} unchanged, {4} failed, {5} removed." -f $summary.Pages, $summary.Created, $summary.Updated, $summary.Unchanged, $summary.Failed, $summary.Removed)
        [pscustomobject]$summary
    }
}
