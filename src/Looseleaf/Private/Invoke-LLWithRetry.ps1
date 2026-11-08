function Invoke-LLWithRetry {
    <#
    .SYNOPSIS
        Runs a file operation, retrying on sharing violations and access errors (docs/DESIGN.md 8).

    .DESCRIPTION
        Antivirus scanners can briefly lock (or quarantine) files with security-related content.
        The action is retried with back-off; the final error names antivirus as a likely cause.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [scriptblock]$Action,

        [string]$Description = 'file operation',

        [int]$Attempts = 3,

        [int]$DelayMilliseconds = 500
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            return (& $Action)
        } catch [System.IO.IOException], [System.UnauthorizedAccessException] {
            if ($attempt -eq $Attempts) {
                throw ("{0} failed after {1} attempts: {2} The file may be locked or quarantined by antivirus software; consider excluding WorkPath and VaultPath from scanning." -f $Description, $Attempts, $_.Exception.Message)
            }
            Start-Sleep -Milliseconds ($DelayMilliseconds * $attempt)
        } catch {
            # MethodInvocationException wraps .NET exceptions thrown from method calls
            $inner = $_.Exception.InnerException
            if (($inner -is [System.IO.IOException] -or $inner -is [System.UnauthorizedAccessException]) -and $attempt -lt $Attempts) {
                Start-Sleep -Milliseconds ($DelayMilliseconds * $attempt)
                continue
            }
            if ($inner -is [System.IO.IOException] -or $inner -is [System.UnauthorizedAccessException]) {
                throw ("{0} failed after {1} attempts: {2} The file may be locked or quarantined by antivirus software; consider excluding WorkPath and VaultPath from scanning." -f $Description, $Attempts, $inner.Message)
            }
            throw
        }
    }
}
