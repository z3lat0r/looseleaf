$script:LLOneNoteNs = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

function Get-LLXmlNodeList {
    <#
    .SYNOPSIS
        Copies XML nodes into a plain list, so the DOM can be changed while iterating.
    #>
    param ($Nodes)
    $list = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
    if ($null -ne $Nodes) { foreach ($n in $Nodes) { $list.Add($n) } }
    , $list
}

function ConvertFrom-LLHierarchy {
    <#
    .SYNOPSIS
        Turns OneNote hierarchy XML (GetHierarchy, hsPages) into the manifest structure
        (docs/DESIGN.md sections 3.2 and 4). Pure function: no OneNote needed.

    .DESCRIPTION
        Keeps only the requested notebooks (matched on display name or folder name),
        records section groups, sections and pages in OneNote order, skips the recycle bin,
        and marks locked sections. Pages get their stable key (Get-LLPageKey).

    .OUTPUTS
        PSCustomObject: Notebooks (list of ordered dictionaries), Missing (requested names not found).
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Xml,

        [Parameter(Mandatory)]
        [string[]]$Notebooks
    )

    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Xml)
    $ns = New-Object System.Xml.XmlNamespaceManager $doc.NameTable
    $ns.AddNamespace('one', $doc.DocumentElement.NamespaceURI)

    $result = New-Object System.Collections.Generic.List[object]
    $found = @{}
    foreach ($nb in (Get-LLXmlNodeList $doc.SelectNodes('/one:Notebooks/one:Notebook', $ns))) {
        $display = if ($nb.GetAttribute('nickname')) { $nb.GetAttribute('nickname') } else { $nb.GetAttribute('name') }
        $match = @($Notebooks | Where-Object { $_ -eq $display -or $_ -eq $nb.GetAttribute('name') })
        if ($match.Count -eq 0) { continue }
        foreach ($m in $match) { $found[$m] = $true }
        $result.Add([pscustomobject][ordered]@{
                id    = $nb.GetAttribute('ID')
                name  = $display
                path  = $nb.GetAttribute('path')
                items = (ConvertFrom-LLHierarchyContainer -Node $nb -Ns $ns)
            })
    }

    [pscustomobject]@{
        Notebooks = $result
        Missing   = @($Notebooks | Where-Object { -not $found.ContainsKey($_) })
    }
}

function ConvertFrom-LLHierarchyContainer {
    param ([System.Xml.XmlElement]$Node, [System.Xml.XmlNamespaceManager]$Ns)
    $items = New-Object System.Collections.Generic.List[object]
    $order = 0
    foreach ($child in (Get-LLXmlNodeList $Node.ChildNodes)) {
        if ($child -isnot [System.Xml.XmlElement]) { continue }
        if ($child.LocalName -eq 'SectionGroup') {
            if ($child.GetAttribute('isRecycleBin') -eq 'true') { continue }
            $order++
            $items.Add([pscustomobject][ordered]@{
                    type  = 'sectionGroup'
                    id    = $child.GetAttribute('ID')
                    name  = $child.GetAttribute('name')
                    order = $order
                    items = (ConvertFrom-LLHierarchyContainer -Node $child -Ns $Ns)
                })
        } elseif ($child.LocalName -eq 'Section') {
            if ($child.GetAttribute('isInRecycleBin') -eq 'true' -or $child.GetAttribute('isDeletedPages') -eq 'true') { continue }
            $order++
            $locked = $child.GetAttribute('locked') -eq 'true'
            $pages = New-Object System.Collections.Generic.List[object]
            $pageOrder = 0
            foreach ($page in (Get-LLXmlNodeList $child.SelectNodes('one:Page', $Ns))) {
                if ($page.GetAttribute('isInRecycleBin') -eq 'true') { continue }
                $pageOrder++
                $level = 1
                [void][int]::TryParse($page.GetAttribute('pageLevel'), [ref]$level)
                $pages.Add([pscustomobject][ordered]@{
                        id       = $page.GetAttribute('ID')
                        key      = Get-LLPageKey -PageId $page.GetAttribute('ID')
                        name     = $page.GetAttribute('name')
                        order    = $pageOrder
                        level    = [Math]::Max(1, $level)
                        created  = $page.GetAttribute('dateTime')
                        modified = $page.GetAttribute('lastModifiedTime')
                        # filled in by Export-LooseleafNotebook
                        status              = 'pending'
                        docx                = $null
                        pdf                 = $null
                        collapsedParagraphs = 0
                        inkElements         = 0
                        attachments         = @()
                        error               = $null
                    })
            }
            $items.Add([pscustomobject][ordered]@{
                    type   = 'section'
                    id     = $child.GetAttribute('ID')
                    name   = $child.GetAttribute('name')
                    order  = $order
                    status = if ($locked) { 'locked' } else { 'ok' }
                    pages  = $pages
                })
        }
    }
    , $items
}

function Measure-LLNotebook {
    # Counts sections and pages, and lists locked sections, of a parsed notebook.
    param ($Items)
    $stats = [pscustomobject]@{ Sections = 0; Pages = 0; Locked = (New-Object System.Collections.Generic.List[string]) }
    $stack = New-Object System.Collections.Generic.Stack[object]
    # foreach instead of @(): @() on List[object] throws on some Windows PowerShell 5.1 builds.
    foreach ($i in $Items) { $stack.Push($i) }
    while ($stack.Count -gt 0) {
        $item = $stack.Pop()
        if ($item.type -eq 'section') {
            $stats.Sections++
            foreach ($p in $item.pages) { $stats.Pages++ }
            if ($item.status -eq 'locked') { $stats.Locked.Add($item.name) }
        } else {
            foreach ($i in $item.items) { $stack.Push($i) }
        }
    }
    $stats
}

function Get-LLPageKey {
    <#
    .SYNOPSIS
        Stable, file-name-safe page key: first 16 hex characters of SHA-1(page ID).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param ([Parameter(Mandatory)][string]$PageId)
    $sha = [System.Security.Cryptography.SHA1]::Create()
    try {
        $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($PageId))
        ([BitConverter]::ToString($hash)).Replace('-', '').Substring(0, 16).ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Get-LLPageContentInfo {
    <#
    .SYNOPSIS
        Inspects page XML (GetPageContent): collapsed paragraphs, ink, inserted files.
        Pure function: no OneNote needed.
    #>
    [CmdletBinding()]
    param ([Parameter(Mandatory)][string]$Xml)

    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Xml)
    $ns = New-Object System.Xml.XmlNamespaceManager $doc.NameTable
    $ns.AddNamespace('one', $doc.DocumentElement.NamespaceURI)

    $attachments = New-Object System.Collections.Generic.List[object]
    foreach ($f in (Get-LLXmlNodeList $doc.SelectNodes('//one:InsertedFile', $ns))) {
        $attachments.Add([pscustomobject]@{
                Name      = $f.GetAttribute('preferredName')
                PathCache = $f.GetAttribute('pathCache')
            })
    }
    [pscustomobject]@{
        CollapsedParagraphs = $doc.SelectNodes("//one:OE[@collapsed='true'] | //one:OEChildren[@collapsed='true']", $ns).Count
        InkElements         = $doc.SelectNodes('//one:InkDrawing | //one:InkWord | //one:InkParagraph', $ns).Count
        Attachments         = $attachments
    }
}
