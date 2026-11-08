$script:LLPandocPath = $null # resolved on first use

function Invoke-LLPandoc {
    <#
    .SYNOPSIS
        Runs pandoc with the given arguments in a working directory; returns its warnings (stderr).

    .DESCRIPTION
        Uses System.Diagnostics.Process so the working directory is explicit and stderr output
        does not turn into PowerShell errors on Windows PowerShell 5.1. Throws when pandoc fails.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string[]]$Arguments,

        [Parameter(Mandatory)]
        [string]$WorkingDirectory
    )

    if (-not $script:LLPandocPath) {
        $command = Get-Command -Name pandoc -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $command) {
            throw 'pandoc was not found on PATH. Install it from https://pandoc.org/installing.html and open a new PowerShell window.'
        }
        $script:LLPandocPath = $command.Source
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $script:LLPandocPath
    $psi.Arguments = (@($Arguments | ForEach-Object { ConvertTo-LLProcessArgument $_ })) -join ' '
    $psi.WorkingDirectory = $WorkingDirectory
    $psi.UseShellExecute = $false
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true
    $psi.CreateNoWindow = $true
    $psi.StandardErrorEncoding = New-Object System.Text.UTF8Encoding $false
    $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding $false

    $process = [System.Diagnostics.Process]::Start($psi)
    try {
        $stderrTask = $process.StandardError.ReadToEndAsync()
        [void]$process.StandardOutput.ReadToEnd()
        $process.WaitForExit()
        $stderr = $stderrTask.Result
        if ($process.ExitCode -ne 0) {
            throw "pandoc failed (exit code $($process.ExitCode)): $($stderr.Trim())"
        }
        $stderr.Trim()
    } finally {
        $process.Dispose()
    }
}

function ConvertTo-LLProcessArgument {
    <#
    .SYNOPSIS
        Quotes one argument for a Windows command line (CommandLineToArgvW rules).
    #>
    param ([AllowEmptyString()][string]$Value)
    if ($Value -and $Value -notmatch '[\s"]') { return $Value }
    $escaped = $Value -replace '(\\*)"', '$1$1\"'
    $escaped = $escaped -replace '(\\+)$', '$1$1'
    '"' + $escaped + '"'
}
