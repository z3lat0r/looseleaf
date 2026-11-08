function Export-LooseleafNotebook {
    <#
    .SYNOPSIS
        Stage 1: exports the configured OneNote notebooks to the work folder (docs/DESIGN.md section 3).

    .DESCRIPTION
        Connects to OneNote desktop, reads the hierarchy of the configured notebooks and, per page:
          - skips it if it is unchanged since the last export and its .docx exists (resumable),
          - inspects it (collapsed paragraphs, ink, attachments),
          - publishes it to docx\<key>.docx (and pdf\<key>.pdf for ink pages if configured),
          - copies its attachments to attachments\<key>\.
        Writes <WorkPath>\manifest.json (saved regularly, so an interrupted export can resume).
        Reads from OneNote only; it does not change your notebooks.

    .PARAMETER ConfigPath
        Looseleaf configuration file (.psd1).

    .PARAMETER Force
        Re-export every page, even unchanged ones.

    .PARAMETER StopOnError
        Stop at the first page that fails.

    .OUTPUTS
        One summary object per notebook.

    .EXAMPLE
        Export-LooseleafNotebook -ConfigPath .\looseleaf.config.psd1
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [string]$ConfigPath,

        [switch]$Force,

        [switch]$StopOnError
    )

    $config = Import-LLConfig -Path $ConfigPath
    if ($config.ExpandCollapsed) {
        Write-Warning 'ExpandCollapsed is not implemented yet and is ignored. Pages with collapsed paragraphs are counted and reported instead (see docs/TROUBLESHOOTING.md).'
    }

    $work = $config.WorkPath
    $manifestPath = Join-Path $work 'manifest.json'
    $docxDir = Join-Path $work 'docx'
    $pdfDir = Join-Path $work 'pdf'
    $attachDir = Join-Path $work 'attachments'
    if (-not $WhatIfPreference) {
        foreach ($d in $work, $docxDir, (Join-Path $work 'logs')) { [void][IO.Directory]::CreateDirectory($d) }
    }
    $logPath = if (-not $WhatIfPreference) { Join-Path $work ('logs\{0}-export.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss')) } else { $null }
    $log = {
        param([string]$Level, [string]$Message)
        if ($logPath) {
            [IO.File]::AppendAllText($logPath, ('{0} {1,-5} {2}{3}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message, [Environment]::NewLine), (New-Object System.Text.UTF8Encoding $false))
        }
        if ($Level -eq 'WARN') { Write-Warning $Message } else { Write-Verbose $Message }
    }

    # Previous manifest: lets unchanged pages be skipped.
    $previous = @{}
    if (Test-Path -LiteralPath $manifestPath) {
        try {
            $old = Import-LLManifest -Path $manifestPath
            $stack = New-Object System.Collections.Generic.Stack[object]
            foreach ($nb in @(Get-LLValue $old 'notebooks' @())) { foreach ($i in @(Get-LLValue $nb 'items' @())) { $stack.Push($i) } }
            while ($stack.Count -gt 0) {
                $item = $stack.Pop()
                foreach ($p in @(Get-LLValue $item 'pages' @())) { $previous[[string](Get-LLValue $p 'key' '')] = $p }
                foreach ($i in @(Get-LLValue $item 'items' @())) { $stack.Push($i) }
            }
        } catch {
            & $log 'WARN' "Previous manifest ignored: $($_.Exception.Message)"
        }
    }

    $onenote = Connect-LLOneNote
    try {
        $parsed = ConvertFrom-LLHierarchy -Xml (Get-LLOneNoteHierarchy -OneNote $onenote) -Notebooks $config.Notebooks
        foreach ($m in $parsed.Missing) {
            & $log 'WARN' "Notebook '$m' is not open in OneNote and was skipped. Open it (File > Open) and run the export again."
        }

        $manifest = [ordered]@{
            schemaVersion = 1
            generator     = 'Looseleaf'
            exportedAt    = (Get-Date).ToString('o')
            notebooks     = $parsed.Notebooks
        }
        $saveManifest = {
            if ($WhatIfPreference) { return }
            $manifest.exportedAt = (Get-Date).ToString('o')
            $tmp = "$manifestPath.tmp"
            [IO.File]::WriteAllText($tmp, ($manifest | ConvertTo-Json -Depth 32), (New-Object System.Text.UTF8Encoding $false))
            if (Test-Path -LiteralPath $manifestPath) { Remove-Item -LiteralPath $manifestPath -Force }
            Move-Item -LiteralPath $tmp -Destination $manifestPath
        }

        foreach ($nb in $parsed.Notebooks) {
            $summary = [ordered]@{
                Notebook            = $nb.name
                Pages               = 0
                Exported            = 0
                Unchanged           = 0
                Failed              = 0
                LockedSections      = 0
                PagesWithCollapsed  = 0
                PagesWithInk        = 0
                Attachments         = 0
                FailedPages         = New-Object System.Collections.Generic.List[object]
                ManifestPath        = $manifestPath
                LogPath             = $logPath
            }
            & $log 'INFO' "Exporting notebook '$($nb.name)' to $work"

            # Flatten sections in order.
            $sections = New-Object System.Collections.Generic.List[object]
            $queue = New-Object System.Collections.Generic.List[object]
            foreach ($i in $nb.items) { $queue.Add($i) }
            while ($queue.Count -gt 0) {
                $item = $queue[0]; $queue.RemoveAt(0)
                if ($item.type -eq 'section') { $sections.Add($item) }
                else { $queue.InsertRange(0, [object[]]@(foreach ($c in $item.items) { $c })) }
            }
            # Note: no @() around List[object] - it throws on some Windows PowerShell 5.1 builds (DESIGN 10).
            $total = 0; foreach ($s in $sections) { $total += $s.pages.Count }
            $summary.Pages = $total

            $index = 0
            foreach ($section in $sections) {
                if ($section.status -eq 'locked') {
                    $summary.LockedSections++
                    & $log 'WARN' "Section '$($section.name)' is locked; unlock it in OneNote to export it."
                    continue
                }
                foreach ($page in $section.pages) {
                    $index++
                    Write-Progress -Activity "Looseleaf: exporting '$($nb.name)'" -Status "$($section.name) / $($page.name)" -PercentComplete ([int](100 * $index / [Math]::Max(1, $total)))
                    $key = $page.key
                    $docxRel = "docx/$key.docx"
                    $docxPath = Join-Path $docxDir "$key.docx"
                    $page.status = 'exported'
                    $page.docx = $docxRel
                    $page.pdf = $null
                    $page.collapsedParagraphs = 0
                    $page.inkElements = 0
                    $page.attachments = @()
                    $page.error = $null

                    $prev = $previous[$key]
                    if (-not $Force -and $prev -and [string](Get-LLValue $prev 'modified' '') -eq $page.modified -and
                        [string](Get-LLValue $prev 'status' '') -in @('exported', 'unchanged') -and (Test-Path -LiteralPath $docxPath)) {
                        $page.status = 'unchanged'
                        $page.pdf = Get-LLValue $prev 'pdf'
                        $page.collapsedParagraphs = [int](Get-LLValue $prev 'collapsedParagraphs' 0)
                        $page.inkElements = [int](Get-LLValue $prev 'inkElements' 0)
                        $page.attachments = @(Get-LLValue $prev 'attachments' @())
                        $summary.Unchanged++
                    } elseif ($PSCmdlet.ShouldProcess("$($section.name) / $($page.name)", 'Export page')) {
                        try {
                            $info = Get-LLPageContentInfo -Xml (Get-LLOneNotePageXml -OneNote $onenote -PageId $page.id)
                            $page.collapsedParagraphs = $info.CollapsedParagraphs
                            $page.inkElements = $info.InkElements

                            Publish-LLOneNotePage -OneNote $onenote -PageId $page.id -Destination $docxPath -Format Word

                            if ($config.ExportInkPagesAsPdf -and $info.InkElements -gt 0) {
                                [void][IO.Directory]::CreateDirectory($pdfDir)
                                Publish-LLOneNotePage -OneNote $onenote -PageId $page.id -Destination (Join-Path $pdfDir "$key.pdf") -Format Pdf
                                $page.pdf = "pdf/$key.pdf"
                            }

                            $atts = New-Object System.Collections.Generic.List[object]
                            $usedNames = @{}
                            foreach ($a in $info.Attachments) {
                                $safe = Get-LLSafeName -Name $a.Name -NameSpaces Keep -MaxLength 120
                                $candidate = $safe; $n = 2
                                while ($usedNames.ContainsKey($candidate.ToLowerInvariant())) {
                                    $candidate = '{0}-{1}{2}' -f [IO.Path]::GetFileNameWithoutExtension($safe), $n, [IO.Path]::GetExtension($safe); $n++
                                }
                                $usedNames[$candidate.ToLowerInvariant()] = $true
                                if (-not $a.PathCache -or -not (Test-Path -LiteralPath $a.PathCache -PathType Leaf)) {
                                    & $log 'WARN' "Attachment '$($a.Name)' on '$($page.name)' is not in OneNote's cache; open the page in OneNote once and export again."
                                    continue
                                }
                                $pageAttachDir = Join-Path $attachDir $key
                                [void][IO.Directory]::CreateDirectory($pageAttachDir)
                                Invoke-LLWithRetry -Description "copy attachment $($a.Name)" -Action { [IO.File]::Copy($a.PathCache, (Join-Path $pageAttachDir $candidate), $true) }
                                $atts.Add([pscustomobject][ordered]@{ name = $a.Name; file = "attachments/$key/$candidate" })
                            }
                            $page.attachments = $atts
                            $summary.Attachments += $atts.Count
                            $summary.Exported++
                        } catch {
                            $page.status = 'failed'
                            $page.error = $_.Exception.Message
                            $summary.Failed++
                            $summary.FailedPages.Add([pscustomobject]@{ Section = $section.name; Title = $page.name; Error = $_.Exception.Message })
                            & $log 'WARN' "Page '$($section.name) / $($page.name)' failed: $($_.Exception.Message)"
                            if ($StopOnError) { & $saveManifest; throw }
                        }
                    }
                    if ($page.collapsedParagraphs -gt 0) { $summary.PagesWithCollapsed++ }
                    if ($page.inkElements -gt 0) { $summary.PagesWithInk++ }
                    if ($index % 25 -eq 0) { & $saveManifest }
                }
            }
            Write-Progress -Activity "Looseleaf: exporting '$($nb.name)'" -Completed

            if ($summary.PagesWithCollapsed -gt 0) {
                & $log 'WARN' "$($summary.PagesWithCollapsed) page(s) contain collapsed paragraphs; their hidden content is missing from the export. Expand them in OneNote (or with the Onetastic macro) and export again."
            }
            if ($summary.PagesWithInk -gt 0) {
                & $log 'WARN' "$($summary.PagesWithInk) page(s) contain ink (handwriting/drawings), which Markdown cannot keep. Consider ExportInkPagesAsPdf = `$true."
            }
            & $log 'INFO' ("Done: {0} pages, {1} exported, {2} unchanged, {3} failed." -f $summary.Pages, $summary.Exported, $summary.Unchanged, $summary.Failed)
            [pscustomobject]$summary
        }
        & $saveManifest
    } finally {
        Disconnect-LLOneNote $onenote
        $partial = Join-Path $docxDir '.partial'
        if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false }
    }
}
