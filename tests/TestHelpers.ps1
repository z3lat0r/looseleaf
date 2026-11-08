# Shared test setup. Dot-source from a *.Tests.ps1 file:  . "$PSScriptRoot\TestHelpers.ps1"
# Written for Pester 3.4 (built into Windows 10/11); also works with Pester 4.

# Same strict mode as the module (Looseleaf.psm1), so tests catch what users would hit.
Set-StrictMode -Version 3.0

$script:LLTestRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:LLSrc = Join-Path $script:LLTestRoot '..\src\Looseleaf'
foreach ($file in @(Get-ChildItem -Path (Join-Path $script:LLSrc 'Private\*.ps1')) + @(Get-ChildItem -Path (Join-Path $script:LLSrc 'Public\*.ps1'))) {
    . $file.FullName
}

$script:LLHasPandoc = [bool](Get-Command -Name pandoc -CommandType Application -ErrorAction SilentlyContinue)

# 1x1 PNG
$script:LLTinyPng = [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==')

function New-LLTestDocx {
    <#
        Creates a .docx that looks like a OneNote page export: title, date and time
        paragraphs, then the given Markdown body (converted by pandoc).
    #>
    param (
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Title,
        [string]$Body = '',
        [switch]$WithImage
    )
    $dir = Split-Path -Parent $Path
    [void][IO.Directory]::CreateDirectory($dir)
    $md = "$Title`n`nMonday, 1 January 2024`n`n10:00 AM`n`n$Body`n"
    if ($WithImage) {
        [IO.File]::WriteAllBytes((Join-Path $dir 'pic.png'), $script:LLTinyPng)
        $md += "`n![](pic.png)`n"
    }
    $mdPath = "$Path.md"
    [IO.File]::WriteAllText($mdPath, $md, (New-Object System.Text.UTF8Encoding $false))
    [void](Invoke-LLPandoc -Arguments @('-f', 'markdown', '-t', 'docx', '-o', $Path, $mdPath) -WorkingDirectory $dir)
    Remove-Item -LiteralPath $mdPath
}

function New-LLTestPage {
    param ([string]$Key, [string]$Name, [int]$Order, [int]$Level = 1, [string]$Status = 'exported')
    [pscustomobject]@{
        id = "{$Key}"; key = $Key; name = $Name; order = $Order; level = $Level
        created = '2024-01-01T10:00:00Z'; modified = '2024-01-02T10:00:00Z'
        status = $Status; docx = "docx/$Key.docx"; attachments = @()
    }
}

function ConvertTo-LLTestManifestObject {
    # Round-trip through JSON so tests see exactly what Import-LLManifest returns.
    param ($Manifest)
    ($Manifest | ConvertTo-Json -Depth 20) | ConvertFrom-Json
}
