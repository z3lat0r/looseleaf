function ConvertTo-LLCodeMarkedDocx {
    <#
    .SYNOPSIS
        Copies a .docx and marks monospace-font text with Word code styles so pandoc emits code (rule C2).

    .DESCRIPTION
        OneNote's Word export formats code with direct fonts (e.g. Consolas, Cascadia Code), which
        pandoc ignores. This function writes a copy of the .docx in which:
          - every run whose font is in -MonospaceFonts gets the character style 'Verbatim Char'
            (pandoc: inline Code), and
          - every paragraph whose visible text is entirely monospace gets the paragraph style
            'Source Code' (pandoc: CodeBlock).
        Both styles are added to word/styles.xml if missing. The source file is never modified.

    .OUTPUTS
        PSCustomObject with Path, RunsMarked, ParagraphsMarked.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Destination,

        [string[]]$MonospaceFonts = @(
            'Consolas', 'Cascadia Code', 'Cascadia Mono', 'Courier New', 'Lucida Console',
            'Source Code Pro', 'Fira Code', 'JetBrains Mono'
        )
    )

    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

    $wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $normalize = { param([string]$Name) ($Name -replace '[\s\-_]', '').ToLowerInvariant() }
    $monoSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($font in $MonospaceFonts) { [void]$monoSet.Add((& $normalize $font)) }

    [IO.File]::Copy($Path, $Destination, $true)

    $runsMarked = 0
    $paragraphsMarked = 0
    $zip = [IO.Compression.ZipFile]::Open($Destination, [IO.Compression.ZipArchiveMode]::Update)
    try {
        # --- document.xml ----------------------------------------------------
        $doc = Read-LLZipXml -Zip $zip -EntryName 'word/document.xml'
        $ns = New-Object System.Xml.XmlNamespaceManager $doc.NameTable
        $ns.AddNamespace('w', $wNs)

        foreach ($p in @($doc.SelectNodes('//w:p', $ns))) {
            $textRuns = 0
            $monoRuns = 0
            foreach ($r in @($p.SelectNodes('w:r | w:hyperlink/w:r', $ns))) {
                $text = (@($r.SelectNodes('w:t', $ns)) | ForEach-Object { $_.InnerText }) -join ''
                if ([string]::IsNullOrWhiteSpace($text)) { continue }
                $textRuns++

                $fontsNode = $r.SelectSingleNode('w:rPr/w:rFonts', $ns)
                if (-not $fontsNode) { continue }
                $font = $fontsNode.GetAttribute('ascii', $wNs)
                if (-not $font) { $font = $fontsNode.GetAttribute('hAnsi', $wNs) }
                if (-not $font -or -not $monoSet.Contains((& $normalize $font))) { continue }

                $monoRuns++
                $rPr = $r.SelectSingleNode('w:rPr', $ns)
                if (-not $rPr.SelectSingleNode('w:rStyle', $ns)) {
                    $rStyle = $doc.CreateElement('w', 'rStyle', $wNs)
                    [void]$rStyle.SetAttribute('val', $wNs, 'VerbatimChar')
                    [void]$rPr.PrependChild($rStyle)
                    $runsMarked++
                }
            }

            if ($textRuns -gt 0 -and $monoRuns -eq $textRuns) {
                $pPr = $p.SelectSingleNode('w:pPr', $ns)
                if (-not $pPr) {
                    $pPr = $doc.CreateElement('w', 'pPr', $wNs)
                    [void]$p.PrependChild($pPr)
                }
                $pStyle = $pPr.SelectSingleNode('w:pStyle', $ns)
                if (-not $pStyle) {
                    $pStyle = $doc.CreateElement('w', 'pStyle', $wNs)
                    [void]$pPr.PrependChild($pStyle)
                }
                [void]$pStyle.SetAttribute('val', $wNs, 'SourceCode')
                $paragraphsMarked++
            }
        }
        Write-LLZipXml -Zip $zip -EntryName 'word/document.xml' -Xml $doc

        # --- styles.xml ------------------------------------------------------
        if ($runsMarked -gt 0 -or $paragraphsMarked -gt 0) {
            $styles = Read-LLZipXml -Zip $zip -EntryName 'word/styles.xml'
            $sns = New-Object System.Xml.XmlNamespaceManager $styles.NameTable
            $sns.AddNamespace('w', $wNs)
            $definitions = @(
                @{ Type = 'character'; Id = 'VerbatimChar'; Name = 'Verbatim Char' }
                @{ Type = 'paragraph'; Id = 'SourceCode'; Name = 'Source Code' }
            )
            foreach ($def in $definitions) {
                if ($styles.SelectSingleNode("//w:style[@w:styleId='$($def.Id)']", $sns)) { continue }
                $style = $styles.CreateElement('w', 'style', $wNs)
                [void]$style.SetAttribute('type', $wNs, $def.Type)
                [void]$style.SetAttribute('customStyle', $wNs, '1')
                [void]$style.SetAttribute('styleId', $wNs, $def.Id)
                $nameNode = $styles.CreateElement('w', 'name', $wNs)
                [void]$nameNode.SetAttribute('val', $wNs, $def.Name)
                [void]$style.AppendChild($nameNode)
                [void]$styles.DocumentElement.AppendChild($style)
            }
            Write-LLZipXml -Zip $zip -EntryName 'word/styles.xml' -Xml $styles
        }
    } finally {
        $zip.Dispose()
    }

    [pscustomobject]@{
        Path             = $Destination
        RunsMarked       = $runsMarked
        ParagraphsMarked = $paragraphsMarked
    }
}
