function Get-LLSafeName {
    <#
    .SYNOPSIS
        Turns a OneNote title into a file or folder name that is valid on NTFS and safe in
        Obsidian links (docs/DESIGN.md section 6.2, rules 1-8; order prefix and clashes are
        handled by the caller).
    .EXAMPLE
        Get-LLSafeName '[ LAB 1.2 ] Demo: Scan'   # (LAB 1.2) Demo - Scan
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, Position = 0)]
        [AllowEmptyString()]
        [string]$Name,

        [ValidateSet('Keep', 'Hyphen')]
        [string]$NameSpaces = 'Keep',

        [ValidateRange(1, 255)]
        [int]$MaxLength = 120
    )

    $n = $Name -replace '[ \t\r\n]', ' '

    # 1. Decorative runs at either end: '----- Linux -----' -> 'Linux'
    $n = $n -replace '^[\s\-=_*~.]+', '' -replace '[\s\-=_*~.]+$', ''

    # 2. Readable replacements: [ ] -> ( ), ':' -> ' - '
    $n = $n.Replace('[', '(').Replace(']', ')')
    $n = $n -replace '\s*:\s*', ' - '

    # 3. Characters invalid on NTFS or breaking Obsidian links
    $n = $n -replace '[<>"/\\|?*#^\x00-\x1F]', '-'

    # 4. Collapse and trim
    $n = $n -replace '\(\s+', '(' -replace '\s+\)', ')'
    $n = $n -replace '\s{2,}', ' ' -replace '-{2,}', '-' -replace '(\s-)+\s', ' - '
    $n = $n.Trim(' ', '-').TrimEnd('.', ' ')

    # 5. Spaces
    if ($NameSpaces -eq 'Hyphen') {
        $n = ($n -replace '\s+', '-') -replace '-{2,}', '-'
    }

    # 6. Reserved device names
    if ($n -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') {
        $n = "_$n"
    }

    # 7. Empty
    if ([string]::IsNullOrWhiteSpace($n)) {
        $n = 'Untitled'
    }

    # 8. Length
    if ($n.Length -gt $MaxLength) {
        $n = $n.Substring(0, $MaxLength).TrimEnd(' ', '-', '.')
        if (-not $n) { $n = 'Untitled' }
    }
    $n
}

function Get-LLImageLabel {
    <#
    .SYNOPSIS
        ASCII-only label for image file names (no spaces, no brackets). Accents are removed,
        other non-ASCII characters become '-'. ASCII keeps names safe for every tool,
        including pandoc's Lua runtime on Windows.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, Position = 0)]
        [AllowEmptyString()]
        [string]$Name,

        [int]$MaxLength = 40
    )

    $n = Get-LLSafeName -Name $Name -NameSpaces Hyphen -MaxLength 255
    $decomposed = $n.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $decomposed.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($category -eq [Globalization.UnicodeCategory]::NonSpacingMark) { continue }
        [void]$sb.Append($ch)
    }
    $n = $sb.ToString() -replace '&', 'and'
    $n = $n -replace '[^A-Za-z0-9._-]', '-' -replace '-{2,}', '-'
    $n = $n.Trim('-', '.', '_')
    if (-not $n) { $n = 'untitled' }
    if ($n.Length -gt $MaxLength) {
        $n = $n.Substring(0, $MaxLength).TrimEnd('-', '.', '_')
    }
    $n
}
