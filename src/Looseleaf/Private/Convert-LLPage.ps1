function Convert-LLPage {
    <#
    .SYNOPSIS
        Converts one page's .docx to Markdown in its own temporary folder (docs/DESIGN.md 5.1, steps 1-4).

    .DESCRIPTION
        Copies the .docx into TempDir (marking monospace text when DetectCode is on), runs pandoc
        with filters/looseleaf.lua, and returns the Markdown (with 'LLMEDIA:<name>' image
        placeholders) and the images the filter wrote to TempDir\media.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$DocxPath,

        [Parameter(Mandatory)]
        [string]$TempDir,

        [Parameter(Mandatory)]
        [string]$ImagePrefix,

        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Config,

        [Parameter(Mandatory)]
        [string]$FilterPath
    )

    $mediaDir = Join-Path $TempDir 'media'
    [void][IO.Directory]::CreateDirectory($mediaDir)
    $workDocx = Join-Path $TempDir 'page.docx'
    $workMd = Join-Path $TempDir 'page.md'

    if ($Config.DetectCode) {
        [void](ConvertTo-LLCodeMarkedDocx -Path $DocxPath -Destination $workDocx -MonospaceFonts $Config.MonospaceFonts)
    } else {
        [IO.File]::Copy($DocxPath, $workDocx, $true)
    }

    $invariant = [Globalization.CultureInfo]::InvariantCulture
    $meta = [ordered]@{
        'looseleaf-image-prefix'       = $ImagePrefix
        'looseleaf-media-dir'          = 'media' # relative: Lua's file API is not Unicode-safe on Windows
        'looseleaf-box-tables'         = $Config.BoxTables
        'looseleaf-box-code-threshold' = ([double]$Config.BoxCodeThreshold).ToString($invariant)
        'looseleaf-indented'           = $Config.IndentedParagraphs
        'looseleaf-glyph-bullets'      = ([string]$Config.ConvertGlyphBullets).ToLowerInvariant()
        'looseleaf-underline'          = $Config.Underline
        'looseleaf-clip-sources'       = $Config.ClipSources
        'looseleaf-image-width'        = $Config.ImageWidth
        'looseleaf-max-image-width'    = ([int]$Config.MaxImageWidth).ToString($invariant)
    }
    $arguments = New-Object System.Collections.Generic.List[string]
    foreach ($a in @('-f', 'docx', '-t', $Config.PandocFormat, '--wrap=none', '--markdown-headings=atx', "--lua-filter=$FilterPath")) {
        $arguments.Add($a)
    }
    foreach ($k in $meta.Keys) {
        $arguments.Add('-M')
        $arguments.Add("$k=$($meta[$k])")
    }
    foreach ($a in @('-o', 'page.md', 'page.docx')) { $arguments.Add($a) }

    $warnings = Invoke-LLPandoc -Arguments $arguments.ToArray() -WorkingDirectory $TempDir
    $markdown = Invoke-LLWithRetry -Description "read $workMd" -Action {
        [IO.File]::ReadAllText($workMd, [Text.Encoding]::UTF8)
    }
    $images = @(Get-ChildItem -LiteralPath $mediaDir -File -ErrorAction SilentlyContinue | Sort-Object Name)

    [pscustomobject]@{
        Markdown = $markdown
        Images   = $images
        IsEmpty  = ([string]::IsNullOrWhiteSpace($markdown) -and $images.Count -eq 0)
        Warnings = $warnings
    }
}
