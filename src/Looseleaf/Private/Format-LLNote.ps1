function Format-LLNote {
    <#
    .SYNOPSIS
        Assembles the final Markdown of a note: front matter, title heading, body, attachments
        (docs/DESIGN.md 5.1 steps 7-8, 6.4). Applies CleanEscapes and NewLine.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string]$Title,

        [AllowEmptyString()]
        [string]$Body = '',

        [System.Collections.IDictionary]$FrontMatter,

        [string[]]$AttachmentLinks = @(),

        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Config
    )

    $parts = New-Object System.Collections.Generic.List[string]
    if ($Config.FrontMatter -and $FrontMatter) {
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('---')
        foreach ($k in $FrontMatter.Keys) {
            $v = $FrontMatter[$k]
            if ($null -eq $v -or [string]$v -eq '') { continue }
            $lines.Add(('{0}: "{1}"' -f $k, (([string]$v).Replace('\', '\\').Replace('"', '\"'))))
        }
        $lines.Add('---')
        $parts.Add(($lines -join "`n"))
    }
    if ($Config.TitleHeading) {
        $parts.Add("# $($Title.Trim())")
    }

    $text = $Body.Replace("`r`n", "`n").Trim("`n")
    if ($Config.CleanEscapes -eq 'Safe') { $text = Remove-LLUnneededEscape -Markdown $text }
    if ($text) { $parts.Add($text) }

    if (@($AttachmentLinks).Count -gt 0) {
        $parts.Add("## Attachments`n`n" + ((@($AttachmentLinks) | ForEach-Object { "- $_" }) -join "`n"))
    }

    $note = ($parts -join "`n`n") + "`n"
    if ($Config.NewLine -eq 'CRLF') { $note = $note.Replace("`n", "`r`n") }
    $note
}

function Remove-LLUnneededEscape {
    <#
    .SYNOPSIS
        Rule C7 (Safe): removes backslash escapes that change nothing, outside code.
        Never unescapes '<', '&' or anything that would start Markdown syntax.
    .DESCRIPTION
        Removed: '\>' not at the start of a line, and escapes before ' " @ %.
    #>
    param ([AllowEmptyString()][string]$Markdown)

    $out = New-Object System.Collections.Generic.List[string]
    $inFence = $false
    $fence = ''
    foreach ($line in ($Markdown -split "`n")) {
        $bare = $line -replace '^(>\s?)+', ''
        if (-not $inFence -and $bare -match '^(`{3,}|~{3,})') {
            $inFence = $true; $fence = $Matches[1]; $out.Add($line); continue
        }
        if ($inFence) {
            if ($bare.TrimEnd() -eq $fence) { $inFence = $false }
            $out.Add($line); continue
        }
        # Keep quote/callout markers aside; '\>' at the start of the content must stay escaped.
        $prefix = [regex]::Match($line, '^(>\s?)*').Value
        $content = $line.Substring($prefix.Length)
        # Split on inline code spans and only touch the text between them.
        $segments = [regex]::Split($content, '(`+[^`]*`+)')
        $seenText = $false
        for ($i = 0; $i -lt $segments.Count; $i++) {
            $s = $segments[$i]
            if ($s -match '^`') { if ($s) { $seenText = $true }; continue }
            if ($seenText) {
                $s = $s.Replace('\>', '>')
            } else {
                $s = [regex]::Replace($s, '(?<=\S.*)\\>', '>') # '\>' after other text in the segment
            }
            $s = [regex]::Replace($s, '\\([''"@%])', '$1')
            if ($s.Trim()) { $seenText = $true }
            $segments[$i] = $s
        }
        $out.Add($prefix + ($segments -join ''))
    }
    $out -join "`n"
}
