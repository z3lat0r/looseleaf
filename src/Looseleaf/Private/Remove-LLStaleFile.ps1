function Remove-LLStaleFile {
    <#
    .SYNOPSIS
        Deletes files written by the previous build that the current build no longer produces
        (page renamed, moved or deleted in OneNote). Edited files are kept (docs/DESIGN.md 5.2).
        Folders left empty by the removal are deleted too.
    .OUTPUTS
        Number of files removed.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param (
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [hashtable]$OldFiles,

        [Parameter(Mandatory)]
        [hashtable]$NewFiles,

        [Parameter(Mandatory)]
        [System.Management.Automation.PSCmdlet]$Cmdlet,

        [scriptblock]$Log
    )

    $removed = 0
    $folders = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($rel in @($OldFiles.Keys | Sort-Object)) {
        if ($NewFiles.ContainsKey($rel)) { continue }
        $full = Join-Path $Root ($rel.Replace('/', '\'))
        if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { continue }
        if ((Get-LLContentHash -LiteralPath $full) -ne $OldFiles[$rel].Hash) {
            if ($Log) { & $Log 'WARN' "Kept edited file that is no longer produced: $rel" }
            continue
        }
        if ($Cmdlet.ShouldProcess($full, 'Remove stale file')) {
            Invoke-LLWithRetry -Description "remove $full" -Action { [IO.File]::Delete($full) }
            [void]$folders.Add((Split-Path -Parent $full))
            $removed++
        }
    }

    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    foreach ($folder in $folders) {
        $dir = $folder
        while ($dir -and $dir.Length -gt $rootFull.Length -and (Test-Path -LiteralPath $dir) -and
            @(Get-ChildItem -LiteralPath $dir -Force).Count -eq 0) {
            Remove-Item -LiteralPath $dir -Force
            $dir = Split-Path -Parent $dir
        }
    }
    $removed
}
