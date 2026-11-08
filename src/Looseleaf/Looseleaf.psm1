#Requires -Version 5.1
# Looseleaf module loader: dot-sources Private\ then Public\ and exports Public functions.
# Convention: one function per file, file name = function name.

Set-StrictMode -Version 3.0

$private = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private\*.ps1') -File -ErrorAction SilentlyContinue)
$public = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public\*.ps1') -File -ErrorAction SilentlyContinue)

foreach ($file in @($private + $public)) {
    try {
        . $file.FullName
    } catch {
        throw "Looseleaf: failed to load $($file.FullName): $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function @($public | ForEach-Object { $_.BaseName })
