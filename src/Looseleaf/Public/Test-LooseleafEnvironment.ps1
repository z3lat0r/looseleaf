function Test-LooseleafEnvironment {
    <#
    .SYNOPSIS
        Preflight checks (docs/DESIGN.md section 7). Read-only: changes nothing.

    .DESCRIPTION
        Checks Windows, PowerShell, pandoc, the configuration, the OneNote COM registration,
        whether OneNote runs at the same privilege level, whether the configured notebooks are
        open, and lists locked sections. OneNote is only contacted if it is already running,
        unless -Connect is given (which may start OneNote in the background).

    .OUTPUTS
        One object per check: Check, Status (Pass | Warn | Fail), Message.
        Pages with collapsed paragraphs or ink are counted during the export itself.

    .EXAMPLE
        Test-LooseleafEnvironment -ConfigPath .\looseleaf.config.psd1 | Format-Table -AutoSize
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$ConfigPath,

        [switch]$Connect
    )

    $results = New-Object System.Collections.Generic.List[object]
    $add = { param($Check, $Status, $Message) $results.Add([pscustomobject]@{ Check = $Check; Status = $Status; Message = $Message }) }

    # Platform
    if ($env:OS -eq 'Windows_NT') { & $add 'Windows' 'Pass' ([Environment]::OSVersion.VersionString) }
    else { & $add 'Windows' 'Fail' 'Stage 1 needs Windows (OneNote COM API).' }
    & $add 'PowerShell' 'Pass' ("$($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)" + $(if ($PSVersionTable.PSVersion.Major -gt 5) { ' (best effort; 5.1 is the baseline)' } else { '' }))

    # pandoc
    $pandoc = Get-Command -Name pandoc -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $pandoc) {
        & $add 'pandoc' 'Fail' 'Not found on PATH. Install from https://pandoc.org/installing.html and open a new PowerShell window.'
    } else {
        $versionLine = (& $pandoc.Source --version | Select-Object -First 1)
        $major = if ($versionLine -match '(\d+)\.') { [int]$Matches[1] } else { 0 }
        if ($major -ge 3) { & $add 'pandoc' 'Pass' $versionLine } else { & $add 'pandoc' 'Fail' "$versionLine (3.x required)" }
    }

    # Configuration
    $config = $null
    try {
        $config = Import-LLConfig -Path $ConfigPath
        & $add 'Configuration' 'Pass' $ConfigPath
    } catch {
        & $add 'Configuration' 'Fail' $_.Exception.Message
    }
    if ($config) {
        foreach ($key in 'WorkPath', 'VaultPath') {
            $path = $config[$key]
            $probeDir = $path
            while ($probeDir -and -not (Test-Path -LiteralPath $probeDir)) { $probeDir = Split-Path -Parent $probeDir }
            if (-not $probeDir) { & $add $key 'Fail' "No existing parent folder for $path"; continue }
            try {
                $probe = Join-Path $probeDir ('.looseleaf-probe-' + [guid]::NewGuid().ToString('N'))
                [IO.File]::WriteAllText($probe, '')
                Remove-Item -LiteralPath $probe -Force
                $note = if ($probeDir -ne $path) { " (will be created under $probeDir)" } else { '' }
                & $add $key 'Pass' "$path is writable$note"
            } catch {
                & $add $key 'Fail' "Cannot write to $probeDir`: $($_.Exception.Message)"
            }
        }
    }

    # OneNote
    $registered = Test-Path -LiteralPath 'Registry::HKEY_CLASSES_ROOT\OneNote.Application\CLSID'
    if ($registered) { & $add 'OneNote COM' 'Pass' 'OneNote.Application is registered (OneNote desktop).' }
    else { & $add 'OneNote COM' 'Fail' 'OneNote desktop is not installed. The Store app "OneNote for Windows 10" has no automation API.' }

    $process = Get-Process -Name ONENOTE -ErrorAction SilentlyContinue | Select-Object -First 1
    $selfElevated = (New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $process) {
        & $add 'OneNote running' 'Warn' 'OneNote is not running. Start it normally (not as Administrator) and open your notebooks; the export can also start it.'
    } else {
        & $add 'OneNote running' 'Pass' "ONENOTE.EXE (pid $($process.Id))"
        $oneNoteElevated = Test-LLProcessElevated -Process $process
        if ($null -eq $oneNoteElevated) {
            & $add 'Privilege level' 'Warn' 'Could not determine whether OneNote runs elevated.'
        } elseif ($oneNoteElevated -ne $selfElevated) {
            $who = if ($selfElevated) { 'PowerShell is elevated but OneNote is not' } else { 'OneNote is elevated but PowerShell is not' }
            & $add 'Privilege level' 'Fail' "$who. Close both and start them normally (not as Administrator)."
        } else {
            & $add 'Privilege level' 'Pass' $(if ($selfElevated) { 'Both elevated (works, but running both normally is recommended).' } else { 'Both run normally.' })
        }
    }

    if ($registered -and $config -and ($process -or $Connect)) {
        $onenote = $null
        try {
            $onenote = Connect-LLOneNote
            $parsed = ConvertFrom-LLHierarchy -Xml (Get-LLOneNoteHierarchy -OneNote $onenote) -Notebooks $config.Notebooks
            foreach ($nb in $parsed.Notebooks) {
                $stats = Measure-LLNotebook -Items $nb.items
                & $add "Notebook '$($nb.name)'" 'Pass' "$($stats.Sections) sections, $($stats.Pages) pages"
                if ($stats.Locked.Count -gt 0) {
                    & $add 'Locked sections' 'Warn' ("{0}: unlock them in OneNote or they are skipped." -f ($stats.Locked -join ', '))
                }
            }
            foreach ($m in $parsed.Missing) {
                & $add "Notebook '$m'" 'Fail' 'Not open in OneNote. Open it (File > Open) and try again.'
            }
        } catch {
            & $add 'OneNote connection' 'Fail' $_.Exception.Message
        } finally {
            Disconnect-LLOneNote $onenote
        }
    } elseif ($registered -and $config) {
        & $add 'Notebooks' 'Warn' 'Not checked: OneNote is not running (use -Connect to start it).'
    }

    $results
}
