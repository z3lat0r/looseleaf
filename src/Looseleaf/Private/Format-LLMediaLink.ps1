function Format-LLMediaLink {
    <#
    .SYNOPSIS
        Markdown link or embed for a file in the media folder (LinkStyle: Relative | Wiki).
    .EXAMPLE
        Format-LLMediaLink -FileName 'a b.png' -Width 800 -LinkPrefix '../media/' -LinkStyle Relative -Embed
        # ![|800](../media/a%20b.png)
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string]$FileName,

        [string]$Width = '',

        [string]$Text = '',

        [string]$LinkPrefix = 'media/',

        [ValidateSet('Relative', 'Wiki')]
        [string]$LinkStyle = 'Relative',

        # Image embed (!) instead of a plain link.
        [switch]$Embed
    )

    $bang = if ($Embed) { '!' } else { '' }
    if ($LinkStyle -eq 'Wiki') {
        $alias = if ($Embed -and $Width) { "|$Width" } elseif (-not $Embed -and $Text -and $Text -ne $FileName) { "|$Text" } else { '' }
        return "$bang[[$FileName$alias]]"
    }
    $target = ($LinkPrefix + $FileName).Replace('%', '%25').Replace(' ', '%20').Replace('(', '%28').Replace(')', '%29')
    $label = if ($Embed) { if ($Width) { "|$Width" } else { '' } } else { if ($Text) { $Text.Replace('[', '\[').Replace(']', '\]') } else { $FileName } }
    "$bang[$label]($target)"
}

function Resolve-LLMediaPlaceholder {
    <#
    .SYNOPSIS
        Replaces the filter's image placeholders '![|W](LLMEDIA:<name>)' with final links.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [AllowEmptyString()]
        [string]$Markdown,

        # filter file name -> final file name
        [hashtable]$Names,

        [string]$LinkPrefix,

        [string]$LinkStyle
    )

    $pattern = '!\[(?:\|(\d+))?\]\(LLMEDIA:([^)\s]+)\)'
    $sb = New-Object System.Text.StringBuilder
    $last = 0
    foreach ($m in [regex]::Matches($Markdown, $pattern)) {
        [void]$sb.Append($Markdown, $last, $m.Index - $last)
        $name = $m.Groups[2].Value
        $final = if ($Names.ContainsKey($name)) { $Names[$name] } else { $name }
        [void]$sb.Append((Format-LLMediaLink -FileName $final -Width $m.Groups[1].Value -LinkPrefix $LinkPrefix -LinkStyle $LinkStyle -Embed))
        $last = $m.Index + $m.Length
    }
    [void]$sb.Append($Markdown, $last, $Markdown.Length - $last)
    $sb.ToString()
}
