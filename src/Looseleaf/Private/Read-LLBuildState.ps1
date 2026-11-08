function Read-LLBuildState {
    <#
    .SYNOPSIS
        Reads <notebook>\.looseleaf\build-state.json (docs/DESIGN.md 5.2). Missing file = empty state.
    .OUTPUTS
        Hashtable: Files = @{ relPath = @{ Hash; Page } }, Media = @{ registryKey = fileName }.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    $state = @{ Files = @{}; Media = @{} }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $state }

    try {
        $json = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json
    } catch {
        throw "Build state '$Path' is unreadable: $($_.Exception.Message). Delete it to start a fresh build (files you edited will then be treated as foreign and left alone)."
    }
    $files = Get-LLValue $json 'files'
    if ($files) {
        foreach ($p in $files.PSObject.Properties) {
            $state.Files[$p.Name] = @{
                Hash = [string](Get-LLValue $p.Value 'hash' '')
                Page = [string](Get-LLValue $p.Value 'page' '')
            }
        }
    }
    $media = Get-LLValue $json 'media'
    if ($media) {
        foreach ($p in $media.PSObject.Properties) { $state.Media[$p.Name] = [string]$p.Value }
    }
    $state
}

function Save-LLBuildState {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [hashtable]$State
    )

    $files = [ordered]@{}
    foreach ($k in ($State.Files.Keys | Sort-Object)) {
        $files[$k] = [ordered]@{ hash = $State.Files[$k].Hash; page = $State.Files[$k].Page }
    }
    $media = [ordered]@{}
    foreach ($k in ($State.Media.Keys | Sort-Object)) { $media[$k] = $State.Media[$k] }

    $document = [ordered]@{
        schemaVersion = 1
        generator     = 'Looseleaf'
        builtAt       = (Get-Date).ToString('o')
        files         = $files
        media         = $media
    }
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $Path))
    $tmp = "$Path.tmp"
    [IO.File]::WriteAllText($tmp, ($document | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding $false))
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    Move-Item -LiteralPath $tmp -Destination $Path
}

function Get-LLContentHash {
    <#
    .SYNOPSIS
        SHA-256 (hex) of a byte array or of a file's content.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Bytes')]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, ParameterSetName = 'Bytes')]
        [AllowEmptyCollection()]
        [byte[]]$Bytes,

        [Parameter(Mandatory, ParameterSetName = 'File')]
        [string]$LiteralPath
    )

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        if ($PSCmdlet.ParameterSetName -eq 'File') {
            $Bytes = Invoke-LLWithRetry -Description "read $LiteralPath" -Action { [IO.File]::ReadAllBytes($LiteralPath) }
        }
        ([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-', '').ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}
