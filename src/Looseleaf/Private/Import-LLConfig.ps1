# Configuration schema: one entry per setting (docs/CONFIGURATION.md).
# Type: bool | int | double | string | string[] | path. Allowed: valid values (case-insensitive).
$script:LLConfigSchema = [ordered]@{
    # Paths
    Notebooks           = @{ Type = 'string[]'; Required = $true }
    WorkPath            = @{ Type = 'path'; Required = $true }
    VaultPath           = @{ Type = 'path'; Required = $true }
    # Stage 1
    ExpandCollapsed     = @{ Type = 'bool'; Default = $false }
    RestoreCollapsed    = @{ Type = 'bool'; Default = $true }
    ExportInkPagesAsPdf = @{ Type = 'bool'; Default = $false }
    # Stage 2: layout and naming
    OrderPrefix         = @{ Type = 'bool'; Default = $true }
    MediaScope          = @{ Type = 'string'; Default = 'Notebook'; Allowed = @('Notebook', 'Folder') }
    ImageNameStyle      = @{ Type = 'string'; Default = 'SectionPage'; Allowed = @('SectionPage', 'Page') }
    MaxImageLabelLength = @{ Type = 'int'; Default = 40; Min = 8; Max = 80 }
    MaxImageNameLength  = @{ Type = 'int'; Default = 80; Min = 24; Max = 200 }
    LinkStyle           = @{ Type = 'string'; Default = 'Relative'; Allowed = @('Relative', 'Wiki') }
    MaxNameLength       = @{ Type = 'int'; Default = 120; Min = 16; Max = 255 }
    MaxPathLength       = @{ Type = 'int'; Default = 250; Min = 100; Max = 32000 }
    NameSpaces          = @{ Type = 'string'; Default = 'Keep'; Allowed = @('Keep', 'Hyphen') }
    EmptyPages          = @{ Type = 'string'; Default = 'FolderOnly'; Allowed = @('FolderOnly', 'Keep', 'Skip') }
    OnModifiedFile      = @{ Type = 'string'; Default = 'Skip'; Allowed = @('Skip', 'Backup', 'Overwrite') }
    # Stage 2: content
    FrontMatter         = @{ Type = 'bool'; Default = $true }
    TitleHeading        = @{ Type = 'bool'; Default = $true }
    PandocFormat        = @{ Type = 'string'; Default = 'markdown-simple_tables-multiline_tables-grid_tables-smart+pipe_tables' }
    DetectCode          = @{ Type = 'bool'; Default = $true }
    MonospaceFonts      = @{
        Type = 'string[]'
        Default = @('Consolas', 'Cascadia Code', 'Cascadia Mono', 'Courier New', 'Lucida Console',
            'Source Code Pro', 'Fira Code', 'JetBrains Mono')
    }
    BoxTables           = @{ Type = 'string'; Default = 'Auto'; Allowed = @('Auto', 'CodeBlock', 'Callout', 'Unwrap', 'Table') }
    BoxCodeThreshold    = @{ Type = 'double'; Default = 0.6; Min = 0.0; Max = 1.0 }
    ComplexTables       = @{ Type = 'string'; Default = 'Html'; Allowed = @('Html') }
    IndentedParagraphs  = @{ Type = 'string'; Default = 'Unindent'; Allowed = @('Unindent', 'Quote') }
    ConvertGlyphBullets = @{ Type = 'bool'; Default = $true }
    CleanEscapes        = @{ Type = 'string'; Default = 'Safe'; Allowed = @('Safe', 'None') }
    Underline           = @{ Type = 'string'; Default = 'Html'; Allowed = @('Html', 'Drop') }
    ClipSources         = @{ Type = 'string'; Default = 'Normalize'; Allowed = @('Normalize', 'Keep') }
    ImageWidth          = @{ Type = 'string'; Default = 'Obsidian'; Allowed = @('Obsidian', 'Drop') }
    MaxImageWidth       = @{ Type = 'int'; Default = 1000; Min = 100; Max = 4000 }
    NewLine             = @{ Type = 'string'; Default = 'LF'; Allowed = @('LF', 'CRLF') }
}

function Import-LLConfig {
    <#
    .SYNOPSIS
        Loads and validates a Looseleaf configuration data file (.psd1). No code is executed.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Configuration file not found: $Path"
    }
    try {
        $data = Import-PowerShellDataFile -LiteralPath $Path
    } catch {
        throw "Configuration file '$Path' is not a valid PowerShell data file: $($_.Exception.Message)"
    }
    Resolve-LLConfig -Data $data -Source $Path
}

function Resolve-LLConfig {
    <#
    .SYNOPSIS
        Validates configuration values against the schema and fills in defaults.
    .OUTPUTS
        Ordered dictionary with every setting.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [hashtable]$Data,

        [string]$Source = 'configuration'
    )

    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($key in $Data.Keys) {
        if (-not $script:LLConfigSchema.Contains($key)) {
            $errors.Add("Unknown setting '$key'.")
        }
    }

    $config = [ordered]@{}
    foreach ($key in $script:LLConfigSchema.Keys) {
        $rule = $script:LLConfigSchema[$key]
        if (-not $Data.ContainsKey($key)) {
            if ($rule.ContainsKey('Required') -and $rule.Required) {
                $errors.Add("Missing required setting '$key'.")
            } else {
                $config[$key] = $rule.Default
            }
            continue
        }

        $value = $Data[$key]
        $problem = $null
        switch ($rule.Type) {
            'bool' {
                if ($value -isnot [bool]) { $problem = 'must be $true or $false' }
            }
            'int' {
                if ($value -isnot [int]) { $problem = 'must be a whole number' }
                elseif ($value -lt $rule.Min -or $value -gt $rule.Max) { $problem = "must be between $($rule.Min) and $($rule.Max)" }
            }
            'double' {
                if ($value -is [int] -or $value -is [double] -or $value -is [decimal]) {
                    $value = [double]$value
                    if ($value -lt $rule.Min -or $value -gt $rule.Max) { $problem = "must be between $($rule.Min) and $($rule.Max)" }
                } else {
                    $problem = 'must be a number'
                }
            }
            'string' {
                if ($value -isnot [string] -or [string]::IsNullOrWhiteSpace($value)) {
                    $problem = 'must be a non-empty string'
                } elseif ($rule.ContainsKey('Allowed')) {
                    $match = @($rule.Allowed | Where-Object { $_ -eq $value })
                    if ($match.Count -eq 0) { $problem = "must be one of: $($rule.Allowed -join ', ')" }
                    else { $value = $match[0] } # canonical casing
                }
            }
            'string[]' {
                $value = @($value)
                $bad = @($value | Where-Object { $_ -isnot [string] -or [string]::IsNullOrWhiteSpace($_) })
                if ($value.Count -eq 0 -or $bad.Count -gt 0) { $problem = 'must be one or more non-empty strings' }
            }
            'path' {
                if ($value -isnot [string] -or $value -notmatch '^([A-Za-z]:\\|\\\\[^\\]+\\)') {
                    $problem = 'must be an absolute path (e.g. D:\Folder or \\server\share)'
                } else {
                    $value = $value.TrimEnd('\')
                    if ($value -match '^[A-Za-z]:$') { $value += '\' }
                }
            }
        }
        if ($problem) {
            $errors.Add("'$key' $problem.")
        } else {
            $config[$key] = $value
        }
    }

    if ($errors.Count -gt 0) {
        throw ("Invalid configuration in {0}:`n  - {1}" -f $Source, ($errors -join "`n  - "))
    }
    $config
}
