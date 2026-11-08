function Import-LLManifest {
    <#
    .SYNOPSIS
        Reads manifest.json written by stage 1 and checks its schema version (docs/DESIGN.md section 4).
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Manifest not found: $Path. Run Export-LooseleafNotebook first."
    }
    $json = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    try {
        $manifest = $json | ConvertFrom-Json
    } catch {
        throw "Manifest '$Path' is not valid JSON: $($_.Exception.Message)"
    }
    $version = Get-LLValue $manifest 'schemaVersion' 0
    if ($version -ne 1) {
        throw "Unsupported manifest schemaVersion '$version' in $Path. This version of Looseleaf reads schema version 1."
    }
    $manifest
}

function Get-LLValue {
    <#
    .SYNOPSIS
        Safe property access for JSON objects and dictionaries under strict mode.
    #>
    param (
        $Object,
        [string]$Name,
        $Default = $null
    )
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name) -and $null -ne $Object[$Name]) { return $Object[$Name] }
        return $Default
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    $property.Value
}
