function Write-LLManagedFile {
    <#
    .SYNOPSIS
        Writes a vault file while protecting edits made after the last build (docs/DESIGN.md 5.2).

    .DESCRIPTION
        Compares the file on disk with the hash recorded at the last build:
          - same content as the new output     -> Unchanged
          - unchanged since last build          -> overwritten (Updated)
          - edited since last build, or a file Looseleaf never wrote (foreign)
                                                -> OnModifiedFile: Skip | Backup | Overwrite
        Records the result in NewFiles. Honours -WhatIf through the calling cmdlet.

    .OUTPUTS
        Created | Updated | Unchanged | SkippedModified | SkippedForeign | BackedUp | WhatIf
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [string]$RelPath,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [byte[]]$Bytes,

        [Parameter(Mandatory)]
        [string]$PageKey,

        [Parameter(Mandatory)]
        [hashtable]$OldFiles,

        [Parameter(Mandatory)]
        [hashtable]$NewFiles,

        [Parameter(Mandatory)]
        [ValidateSet('Skip', 'Backup', 'Overwrite')]
        [string]$Mode,

        [Parameter(Mandatory)]
        [System.Management.Automation.PSCmdlet]$Cmdlet
    )

    $stateKey = $RelPath.Replace('\', '/')
    $fullPath = Join-Path $Root $RelPath
    $newHash = Get-LLContentHash -Bytes $Bytes
    $result = 'Created'

    if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
        $currentHash = Get-LLContentHash -LiteralPath $fullPath
        if ($currentHash -eq $newHash) {
            $NewFiles[$stateKey] = @{ Hash = $newHash; Page = $PageKey }
            return 'Unchanged'
        }
        $known = $OldFiles.ContainsKey($stateKey)
        $edited = (-not $known) -or ($OldFiles[$stateKey].Hash -ne $currentHash)
        $result = 'Updated'
        if ($edited) {
            switch ($Mode) {
                'Skip' {
                    if ($known) {
                        $NewFiles[$stateKey] = $OldFiles[$stateKey] # keep tracking: still counts as edited
                        return 'SkippedModified'
                    }
                    return 'SkippedForeign'
                }
                'Backup' {
                    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
                    $backup = [IO.Path]::Combine((Split-Path -Parent $fullPath),
                        ('{0}.edited-{1}{2}' -f [IO.Path]::GetFileNameWithoutExtension($fullPath), $stamp, [IO.Path]::GetExtension($fullPath)))
                    if (-not $Cmdlet.ShouldProcess($fullPath, "Back up edited file to $(Split-Path -Leaf $backup)")) { return 'WhatIf' }
                    Invoke-LLWithRetry -Description "back up $fullPath" -Action { [IO.File]::Move($fullPath, $backup) }
                    $result = 'BackedUp'
                }
                'Overwrite' { $result = 'Updated' }
            }
        }
    }

    $action = if ($result -eq 'Created') { 'Create file' } else { 'Overwrite file' }
    if ($result -ne 'BackedUp' -and -not $Cmdlet.ShouldProcess($fullPath, $action)) { return 'WhatIf' }

    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $fullPath))
    Invoke-LLWithRetry -Description "write $fullPath" -Action { [IO.File]::WriteAllBytes($fullPath, $Bytes) }
    $NewFiles[$stateKey] = @{ Hash = $newHash; Page = $PageKey }
    $result
}
