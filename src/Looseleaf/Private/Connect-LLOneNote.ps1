function Connect-LLOneNote {
    <#
    .SYNOPSIS
        Connects to OneNote desktop through late-bound COM (docs/DESIGN.md 3.1).
        No Office Interop assembly is loaded, so this works beyond PowerShell 7.0.
    #>
    [CmdletBinding()]
    param ()

    try {
        New-Object -ComObject OneNote.Application -ErrorAction Stop
    } catch {
        $msg = $_.Exception.Message
        $hint = if ($msg -match '80040154|Class not registered') {
            'OneNote desktop (2016 / Microsoft 365) is not installed or not registered. The Store app "OneNote for Windows 10" has no automation API.'
        } elseif ($msg -match '80080005|80070005|denied') {
            'PowerShell and OneNote probably run at different privilege levels. Close OneNote and start both normally (not as Administrator).'
        } else {
            'Make sure OneNote desktop is installed and can start normally.'
        }
        throw "Cannot connect to OneNote: $msg $hint"
    }
}

function Disconnect-LLOneNote {
    param ($OneNote)
    if ($null -ne $OneNote) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($OneNote) } catch { }
    }
}

function Get-LLOneNoteHierarchy {
    # GetHierarchy('', hsPages = 4): all open notebooks down to pages.
    param ([Parameter(Mandatory)]$OneNote)
    $xml = ''
    $OneNote.GetHierarchy('', 4, [ref]$xml)
    $xml
}

function Get-LLOneNotePageXml {
    # GetPageContent(pageId, out xml, piBasic = 0)
    param ([Parameter(Mandatory)]$OneNote, [Parameter(Mandatory)][string]$PageId)
    $xml = ''
    $OneNote.GetPageContent($PageId, [ref]$xml, 0)
    $xml
}

function Publish-LLOneNotePage {
    <#
    .SYNOPSIS
        Publishes a page to a file (pfWord = 5, pfPdf = 3). Writes to a temporary name first and
        renames on success, so an interrupted run never leaves a half-written file behind.
    #>
    param (
        [Parameter(Mandatory)]$OneNote,
        [Parameter(Mandatory)][string]$PageId,
        [Parameter(Mandatory)][string]$Destination,
        [ValidateSet('Word', 'Pdf')][string]$Format = 'Word'
    )
    $formatId = if ($Format -eq 'Pdf') { 3 } else { 5 }
    $dir = Split-Path -Parent $Destination
    $partialDir = Join-Path $dir '.partial'
    [void][IO.Directory]::CreateDirectory($partialDir)
    $partial = Join-Path $partialDir (Split-Path -Leaf $Destination)
    if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Force }
    $OneNote.Publish($PageId, $partial, $formatId, '')
    if (-not (Test-Path -LiteralPath $partial -PathType Leaf)) {
        throw "OneNote reported success but did not create $partial"
    }
    if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Force }
    Invoke-LLWithRetry -Description "move $partial" -Action { [IO.File]::Move($partial, $Destination) }
}

function Test-LLProcessElevated {
    <#
    .SYNOPSIS
        $true/$false whether a process runs elevated; $null if it cannot be determined.
        A non-elevated caller cannot open an elevated process, which also means "elevated".
    #>
    param ([Parameter(Mandatory)][System.Diagnostics.Process]$Process)

    if (-not ('Looseleaf.TokenInfo' -as [type])) {
        Add-Type -Namespace Looseleaf -Name TokenInfo -MemberDefinition @'
[DllImport("advapi32.dll", SetLastError = true)]
public static extern bool OpenProcessToken(IntPtr processHandle, uint desiredAccess, out IntPtr tokenHandle);
[DllImport("advapi32.dll", SetLastError = true)]
public static extern bool GetTokenInformation(IntPtr tokenHandle, int tokenInformationClass, out int tokenInformation, int tokenInformationLength, out int returnLength);
[DllImport("kernel32.dll", SetLastError = true)]
public static extern IntPtr OpenProcess(uint access, bool inherit, int processId);
[DllImport("kernel32.dll")]
public static extern bool CloseHandle(IntPtr handle);
public static int IsElevated(int processId) {
    IntPtr process = OpenProcess(0x1000, false, processId); // PROCESS_QUERY_LIMITED_INFORMATION
    if (process == IntPtr.Zero) return -1;
    try {
        IntPtr token;
        if (!OpenProcessToken(process, 0x0008, out token)) return -1; // TOKEN_QUERY
        try {
            int elevated; int length;
            if (!GetTokenInformation(token, 20, out elevated, 4, out length)) return -1; // TokenElevation
            return elevated != 0 ? 1 : 0;
        } finally { CloseHandle(token); }
    } finally { CloseHandle(process); }
}
'@
    }
    $r = [Looseleaf.TokenInfo]::IsElevated($Process.Id)
    if ($r -eq 1) { return $true }
    if ($r -eq 0) { return $false }
    $selfElevated = (New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $selfElevated) { return $true } # access denied from a normal process: target is elevated
    $null
}
