function Get-LLNotebookPage {
    <#
    .SYNOPSIS
        Collects all pages of a notebook (manifest) in OneNote order, with their section name.
        Locked sections are skipped.
    #>
    [CmdletBinding()]
    param (
        [object[]]$Items,

        [string]$SectionName,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Into
    )

    foreach ($item in @($Items | Where-Object { $null -ne $_ } | Sort-Object { [int](Get-LLValue $_ 'order' 0) })) {
        $type = [string](Get-LLValue $item 'type' '')
        if ($type -eq 'section') {
            if ([string](Get-LLValue $item 'status' 'ok') -eq 'locked') { continue }
            $name = [string](Get-LLValue $item 'name' '')
            foreach ($page in @(@(Get-LLValue $item 'pages' @()) | Sort-Object { [int](Get-LLValue $_ 'order' 0) })) {
                $Into.Add([pscustomobject]@{ Page = $page; SectionName = $name })
            }
        } else {
            Get-LLNotebookPage -Items @(Get-LLValue $item 'items' @()) -SectionName $SectionName -Into $Into
        }
    }
}
