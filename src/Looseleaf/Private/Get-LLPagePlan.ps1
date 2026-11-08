function Get-LLPagePlan {
    <#
    .SYNOPSIS
        Works out where every page of a notebook goes in the vault (docs/DESIGN.md section 6.1/6.2).

    .DESCRIPTION
        Walks section groups, sections and pages in OneNote order and returns one plan entry per
        page: its folder (relative to the notebook folder) and file name, or no file for empty
        parent pages with EmptyPages = FolderOnly. Subpages (pageLevel 2/3) go into a folder named
        after their parent page. Names are cleaned, order-prefixed, de-duplicated and shortened to
        fit MaxPathLength.

    .OUTPUTS
        PSCustomObject with Entries (pages), KeepPrefixes (folders of locked sections, whose
        previous output must not be removed) and Warnings.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Notebook,

        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Config,

        [Parameter(Mandatory)]
        [string]$NotebookDir,

        # Keys of pages whose converted content is empty.
        [hashtable]$EmptyKeys = @{}
    )

    $context = @{
        Config       = $Config
        NotebookDir  = $NotebookDir
        EmptyKeys    = $EmptyKeys
        Entries      = New-Object System.Collections.Generic.List[object]
        KeepPrefixes = New-Object System.Collections.Generic.List[string]
        Warnings     = New-Object System.Collections.Generic.List[string]
    }

    $notebookName = [string](Get-LLValue $Notebook 'name' 'Notebook')
    $items = ConvertTo-LLContainerEntry -Items @(Get-LLValue $Notebook 'items' @())
    Add-LLPlanEntry -Entries $items -RelDir '' -SourcePath @($notebookName) -SectionName '' -Context $context

    [pscustomobject]@{
        Entries      = $context.Entries
        KeepPrefixes = $context.KeepPrefixes
        Warnings     = $context.Warnings
    }
}

function ConvertTo-LLContainerEntry {
    # Section groups and sections, in OneNote order.
    param ([object[]]$Items)
    $list = New-Object System.Collections.Generic.List[object]
    $sorted = @($Items | Where-Object { $null -ne $_ } | Sort-Object { [int](Get-LLValue $_ 'order' 0) })
    foreach ($item in $sorted) {
        $list.Add([pscustomobject]@{
                Kind     = 'container'
                Name     = [string](Get-LLValue $item 'name' '')
                Item     = $item
                Children = $null
            })
    }
    , $list
}

function ConvertTo-LLPageTree {
    # Flat page list with pageLevel -> tree (subpages become children of the page above them).
    param ([object[]]$Pages)
    $roots = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.List[object]
    $sorted = @($Pages | Where-Object { $null -ne $_ } | Sort-Object { [int](Get-LLValue $_ 'order' 0) })
    foreach ($page in $sorted) {
        $level = [int](Get-LLValue $page 'level' 1)
        if ($level -lt 1) { $level = 1 }
        $node = [pscustomobject]@{
            Kind     = 'page'
            Name     = [string](Get-LLValue $page 'name' '')
            Item     = $page
            Level    = $level
            Children = New-Object System.Collections.Generic.List[object]
        }
        while ($stack.Count -gt 0 -and $stack[$stack.Count - 1].Level -ge $level) {
            $stack.RemoveAt($stack.Count - 1)
        }
        if ($stack.Count -eq 0) { $roots.Add($node) } else { $stack[$stack.Count - 1].Children.Add($node) }
        $stack.Add($node)
    }
    , $roots
}

function Get-LLVisibleEntry {
    # Applies EmptyPages: drop empty leaf pages (FolderOnly), promote children of empty pages (Skip).
    param ($Entries, [hashtable]$EmptyKeys, [string]$Mode)
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $Entries) {
        if ($entry.Kind -eq 'page') {
            $isEmpty = $EmptyKeys.ContainsKey([string](Get-LLValue $entry.Item 'key' ''))
            if ($isEmpty -and $Mode -eq 'Skip') {
                foreach ($child in (Get-LLVisibleEntry -Entries $entry.Children -EmptyKeys $EmptyKeys -Mode $Mode)) {
                    $out.Add($child)
                }
                continue
            }
            if ($isEmpty -and $Mode -eq 'FolderOnly' -and $entry.Children.Count -eq 0) { continue }
        }
        $out.Add($entry)
    }
    , $out
}

function Add-LLPlanEntry {
    param (
        $Entries,
        [string]$RelDir,
        [string[]]$SourcePath,
        [string]$SectionName,
        [hashtable]$Context
    )
    $cfg = $Context.Config
    $visible = Get-LLVisibleEntry -Entries $Entries -EmptyKeys $Context.EmptyKeys -Mode $cfg.EmptyPages
    $width = if ($visible.Count -gt 99) { 3 } else { 2 }
    $used = @{}
    $index = 0

    foreach ($entry in $visible) {
        $index++
        $base = Get-LLSafeName -Name $entry.Name -NameSpaces $cfg.NameSpaces -MaxLength $cfg.MaxNameLength
        $name = if ($cfg.OrderPrefix) { ("{0:D$width}-{1}" -f $index, $base) } else { $base }

        $candidate = $name
        $n = 2
        while ($used.ContainsKey($candidate.ToLowerInvariant())) {
            $candidate = "$name-$n"
            $n++
        }

        # Keep the longest path that uses this name (<name>.md) within MaxPathLength.
        $dirPath = if ($RelDir) { Join-Path $Context.NotebookDir $RelDir } else { $Context.NotebookDir }
        $fullLength = (Join-Path $dirPath "$candidate.md").Length
        if ($fullLength -gt $cfg.MaxPathLength) {
            $newLength = [Math]::Max(16, $candidate.Length - ($fullLength - $cfg.MaxPathLength))
            $shortened = $candidate.Substring(0, [Math]::Min($newLength, $candidate.Length)).TrimEnd(' ', '-', '.')
            $Context.Warnings.Add("Name shortened to fit MaxPathLength: '$candidate' -> '$shortened' in '$RelDir'.")
            $candidate = $shortened
            if ((Join-Path $dirPath "$candidate.md").Length -gt $cfg.MaxPathLength) {
                $Context.Warnings.Add("Path still longer than MaxPathLength ($($cfg.MaxPathLength)): $(Join-Path $dirPath "$candidate.md")")
            }
        }
        $used[$candidate.ToLowerInvariant()] = $true
        $childRel = if ($RelDir) { "$RelDir\$candidate" } else { $candidate }

        if ($entry.Kind -eq 'container') {
            $type = [string](Get-LLValue $entry.Item 'type' '')
            $childSource = $SourcePath + $entry.Name
            if ($type -eq 'section') {
                if ([string](Get-LLValue $entry.Item 'status' 'ok') -eq 'locked') {
                    $Context.Warnings.Add("Section '$($entry.Name)' is locked and was skipped; its previous output is kept.")
                    $Context.KeepPrefixes.Add($childRel)
                    continue
                }
                $tree = ConvertTo-LLPageTree -Pages @(Get-LLValue $entry.Item 'pages' @())
                Add-LLPlanEntry -Entries $tree -RelDir $childRel -SourcePath $childSource -SectionName $entry.Name -Context $Context
            } else {
                $children = ConvertTo-LLContainerEntry -Items @(Get-LLValue $entry.Item 'items' @())
                Add-LLPlanEntry -Entries $children -RelDir $childRel -SourcePath $childSource -SectionName $SectionName -Context $Context
            }
            continue
        }

        # Page
        $key = [string](Get-LLValue $entry.Item 'key' '')
        $isEmpty = $Context.EmptyKeys.ContainsKey($key)
        $folderOnly = $isEmpty -and $cfg.EmptyPages -eq 'FolderOnly'
        $Context.Entries.Add([pscustomobject]@{
                Key         = $key
                Page        = $entry.Item
                Title       = $entry.Name
                SectionName = $SectionName
                SourcePath  = ($SourcePath -join '/')
                RelDir      = $RelDir
                FileName    = if ($folderOnly) { $null } else { "$candidate.md" }
                IsEmpty     = $isEmpty
            })
        if ($entry.Children.Count -gt 0) {
            Add-LLPlanEntry -Entries $entry.Children -RelDir $childRel -SourcePath ($SourcePath + $entry.Name) -SectionName $SectionName -Context $Context
        }
    }
}
