#Requires -Version 5.1
<#
.SYNOPSIS
    Development tool: runs the stage 2 conversion core (rule C2 pre-pass + pandoc + looseleaf.lua)
    on existing OneNote .docx files, without OneNote and without the full build.

.DESCRIPTION
    For each .docx: copies it to a temporary folder, marks monospace text (ConvertTo-LLCodeMarkedDocx),
    runs pandoc with filters/looseleaf.lua, resolves image placeholders to 'media/<name>', and writes
    <OutputPath>\<title>.md plus images to <OutputPath>\media\. Prints a summary with artefact counts.

    This is a prototype harness for validating the content rules (docs/DESIGN.md 6.5). It is not
    the stage 2 command (Build-LooseleafVault) and does not apply naming, ordering or the media registry.

.EXAMPLE
    .\tools\Invoke-LLPrototype.ps1 -DocxPath D:\Looseleaf\work\docx -OutputPath $env:TEMP\ll-proto -Limit 20
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string[]]$DocxPath,

    [Parameter(Mandatory)]
    [string]$OutputPath,

    [int]$Limit = 0,

    [string]$PandocFormat = 'markdown-simple_tables-multiline_tables-grid_tables-smart+pipe_tables',

    [switch]$NoCodeDetection
)

$ErrorActionPreference = 'Stop'
$moduleRoot = Join-Path $PSScriptRoot '..\src\Looseleaf'
foreach ($name in 'Read-LLZipXml', 'Write-LLZipXml', 'ConvertTo-LLCodeMarkedDocx') {
    . (Join-Path $moduleRoot "Private\$name.ps1")
}
$filter = (Resolve-Path (Join-Path $moduleRoot 'filters\looseleaf.lua')).Path

if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) { throw 'pandoc not found on PATH.' }

$files = @(foreach ($p in $DocxPath) {
        if (Test-Path -LiteralPath $p -PathType Container) {
            Get-ChildItem -LiteralPath $p -Filter *.docx -File | Sort-Object Name
        } else {
            Get-Item -LiteralPath $p
        }
    })
if ($Limit -gt 0) { $files = @($files | Select-Object -First $Limit) }

$mediaDir = Join-Path $OutputPath 'media'
[void][IO.Directory]::CreateDirectory($mediaDir)
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("Looseleaf-proto-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
[void][IO.Directory]::CreateDirectory($tempRoot)
$utf8NoBom = New-Object System.Text.UTF8Encoding $false

function Get-LLProtoTitle([string]$DocxFile) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($DocxFile)
    try {
        $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
        $xml = $reader.ReadToEnd(); $reader.Close()
    } finally { $zip.Dispose() }
    $m = [regex]::Match($xml, '<w:p\b.*?</w:p>')
    $title = (([regex]::Matches($m.Value, '<w:t[^>]*>([^<]*)</w:t>') | ForEach-Object { $_.Groups[1].Value }) -join '')
    [Net.WebUtility]::HtmlDecode($title).Trim()
}

function Get-LLProtoSafeName([string]$Name) {
    $n = $Name -replace '^[\s\-=_*~.]+|[\s\-=_*~.]+$', ''
    $n = $n -replace '\[', '(' -replace '\]', ')' -replace ':', ' -'
    $n = $n -replace '[<>"/\\|?*#^\x00-\x1f]', '-'
    $n = $n -replace '\(\s+', '(' -replace '\s+\)', ')' -replace '\s{2,}', ' ' -replace '-{2,}', '-'
    $n = $n.Trim(' ', '-', '.')
    if (-not $n) { $n = 'Untitled' }
    $n
}

$results = New-Object System.Collections.Generic.List[object]
$usedNames = @{}
$i = 0
try {
    foreach ($f in $files) {
        $i++
        Write-Progress -Activity 'Looseleaf prototype' -Status $f.Name -PercentComplete ([int](100 * $i / [Math]::Max(1, $files.Count)))
        $title = Get-LLProtoTitle $f.FullName
        $base = Get-LLProtoSafeName $title
        $n = 2; $unique = $base
        while ($usedNames.ContainsKey($unique.ToLowerInvariant())) { $unique = "$base-$n"; $n++ }
        $usedNames[$unique.ToLowerInvariant()] = $true
        $imagePrefix = ($unique -replace '\s+', '-' -replace '[()]', '')
        if ($imagePrefix.Length -gt 60) { $imagePrefix = $imagePrefix.Substring(0, 60).TrimEnd('-') }

        $pageTemp = Join-Path $tempRoot ('{0:D4}' -f $i)
        [void][IO.Directory]::CreateDirectory($pageTemp)
        $work = Join-Path $pageTemp 'page.docx'
        $outMd = Join-Path $pageTemp 'page.md'
        $row = [ordered]@{ Title = $title; File = "$unique.md"; Docx = $f.Name; Runs = 0; Paras = 0; Error = $null; Stderr = $null }
        try {
            if ($NoCodeDetection) {
                Copy-Item -LiteralPath $f.FullName -Destination $work
            } else {
                $mark = ConvertTo-LLCodeMarkedDocx -Path $f.FullName -Destination $work
                $row.Runs = $mark.RunsMarked; $row.Paras = $mark.ParagraphsMarked
            }
            $pandocArgs = @(
                '-f', 'docx', '-t', $PandocFormat, '--wrap=none', '--markdown-headings=atx',
                "--lua-filter=$filter",
                '-M', "looseleaf-image-prefix=$imagePrefix",
                '-M', "looseleaf-media-dir=$($mediaDir -replace '\\', '/')",
                '-o', $outMd, $work
            )
            $stderr = & pandoc @pandocArgs 2>&1
            if ($LASTEXITCODE -ne 0) { throw "pandoc exit $LASTEXITCODE`: $($stderr | Out-String)" }
            if ($stderr) { $row.Stderr = ($stderr | Out-String).Trim() }

            $md = [IO.File]::ReadAllText($outMd)
            $md = $md.Replace('LLMEDIA:', 'media/')
            $md = [regex]::Replace($md, '(\r?\n){3,}', "`n`n").Trim("`r", "`n") + "`n"
            $md = "# $title`n`n" + $md
            [IO.File]::WriteAllText((Join-Path $OutputPath "$unique.md"), $md, $utf8NoBom)
        } catch {
            $row.Error = $_.Exception.Message
        }
        $results.Add([pscustomobject]$row)
    }
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    Write-Progress -Activity 'Looseleaf prototype' -Completed
}

# --- Summary -------------------------------------------------------------------
# Split all output into prose and code-block lines (fences may sit inside callouts: '> ```').
$codeBlockLines = New-Object System.Collections.Generic.List[int]
$unclosed = 0
$proseLines = New-Object System.Collections.Generic.List[string]
foreach ($mdFile in Get-ChildItem -LiteralPath $OutputPath -Filter *.md -File) {
    $inCode = $false; $fence = ''; $n = 0
    foreach ($line in [IO.File]::ReadAllLines($mdFile.FullName)) {
        $bare = $line -replace '^(>\s?)+', ''
        if (-not $inCode -and $bare -match '^(`{3,})') { $inCode = $true; $fence = $Matches[1]; $n = 0 }
        elseif ($inCode -and $bare -eq $fence) { $inCode = $false; $codeBlockLines.Add($n) }
        elseif ($inCode) { $n++ }
        else { $proseLines.Add($line) }
    }
    if ($inCode) { $unclosed++ }
}
$prose = [regex]::Replace(($proseLines -join "`n"), '`[^`\n]+`', '')   # also drop inline code
$count = { param($pattern) ([regex]::Matches($prose, $pattern)).Count }
[pscustomobject]@{
    Pages            = $results.Count
    Failed           = @($results | Where-Object Error).Count
    PandocWarnings   = @($results | Where-Object Stderr).Count
    RunsMarkedCode   = ($results | Measure-Object Runs -Sum).Sum
    ParasMarkedCode  = ($results | Measure-Object Paras -Sum).Sum
    Images           = @(Get-ChildItem -LiteralPath $mediaDir -File).Count
    CodeBlocks       = $codeBlockLines.Count
    CodeBlockMaxLines = ($codeBlockLines | Measure-Object -Maximum).Maximum
    UnclosedFences   = $unclosed
    InlineCode       = ([regex]::Matches(($proseLines -join "`n"), '(?<!`)`[^`\n]+`(?!`)')).Count
    NestedCallouts   = & $count '(?m)^> > \[!note\]'
    Callouts         = & $count '(?m)^> \[!note\]'
    HtmlTables       = & $count '<table'
    PipeTableRows    = & $count '(?m)^\|.*\|\s*$'
    OtherHtmlTags    = & $count '</?(?!u>|br>|table|/table|thead|tbody|tr|th|td|colgroup|col\b|https?:)[a-zA-Z][a-zA-Z0-9_]*(\s[^>\n]*)?>'
    PandocAttrSpans  = & $count '\]\{[.#][^}]*\}'
    Highlights       = & $count '==[^=\n]+=='
    GlyphBulletsLeft = & $count '(?m)^\s*[●•○◦▪■]'
    SourceLines      = & $count '(?m)^Source: '
}
$results | Where-Object { $_.Error -or $_.Stderr } | Select-Object Title, Error, Stderr | Format-List
