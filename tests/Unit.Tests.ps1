# Unit tests: naming, configuration, planning, media registry, links, escaping.
# Run in Windows PowerShell:  Invoke-Pester .\tests

. "$PSScriptRoot\TestHelpers.ps1"

Describe 'Get-LLSafeName (DESIGN 6.2)' {
    $cases = @(
        @{ In = '----- Linux -----'; Out = 'Linux' }
        @{ In = '( Windows )'; Out = '(Windows)' }
        @{ In = '[ LAB 1.2 ] Scan the Server'; Out = '(LAB 1.2) Scan the Server' }
        @{ In = 'Demo: Setup'; Out = 'Demo - Setup' }
        @{ In = 'a/b\c|d?e*f#g^h"i<j>k'; Out = 'a-b-c-d-e-f-g-h-i-j-k' }
        @{ In = 'CON'; Out = '_CON' }
        @{ In = 'aux.txt'; Out = '_aux.txt' }
        @{ In = 'Trailing dot.'; Out = 'Trailing dot' }
        @{ In = '2023-10-05 Meeting'; Out = '2023-10-05 Meeting' }
        @{ In = ''; Out = 'Untitled' }
        @{ In = '   '; Out = 'Untitled' }
        @{ In = 'Non' + [char]0x00A0 + 'breaking'; Out = 'Non breaking' }
        @{ In = 'Café – naïve'; Out = 'Café – naïve' }
    )
    foreach ($c in $cases) {
        It "'$($c.In)' -> '$($c.Out)'" {
            Get-LLSafeName -Name $c.In | Should BeExactly $c.Out
        }
    }
    It 'replaces spaces with NameSpaces = Hyphen' {
        Get-LLSafeName -Name 'Kick-off  Notes' -NameSpaces Hyphen | Should BeExactly 'Kick-off-Notes'
    }
    It 'shortens to MaxLength' {
        Get-LLSafeName -Name 'abcdefghij klmno' -MaxLength 11 | Should BeExactly 'abcdefghij'
    }
}

Describe 'Get-LLShortenedName' {
    It 'cuts at a word boundary and drops an unclosed bracket' {
        Get-LLShortenedName -Name '09-Exploiting Windows MS17-010 SMB Vulnerability (EternalBlue)' -MaxLength 53 |
            Should BeExactly '09-Exploiting Windows MS17-010 SMB Vulnerability'
    }
    It 'drops dangling separators' {
        Get-LLShortenedName -Name 'Common Web Application Threats & Risks' -MaxLength 32 | Should BeExactly 'Common Web Application Threats'
    }
    It 'cuts inside a word when there is no usable space' {
        Get-LLShortenedName -Name 'Supercalifragilistic' -MaxLength 10 | Should BeExactly 'Supercalif'
    }
    It 'leaves short names alone' {
        Get-LLShortenedName -Name 'Short' -MaxLength 10 | Should BeExactly 'Short'
    }
}

Describe 'Get-LLImageLabel' {
    It 'is ASCII only, without spaces or brackets' {
        Get-LLImageLabel 'Café & Crème (Demo)' | Should BeExactly 'Cafe-and-Creme-Demo'
    }
    It 'falls back to untitled for non-Latin titles' {
        Get-LLImageLabel ([string][char]0x65E5 + [char]0x672C) | Should BeExactly 'untitled'
    }
    It 'respects MaxLength' {
        (Get-LLImageLabel ('x' * 100) -MaxLength 20).Length | Should Be 20
    }
}

Describe 'Resolve-LLConfig' {
    $minimal = @{ Notebooks = @('My Notebook'); WorkPath = 'D:\Work'; VaultPath = 'D:\Vault' }

    It 'fills in defaults' {
        $cfg = Resolve-LLConfig -Data $minimal
        $cfg.BoxTables | Should BeExactly 'Auto'
        $cfg.OrderPrefix | Should BeExactly 'Auto'
        $cfg.MaxImageWidth | Should Be 1000
    }
    It 'returns allowed values in canonical casing' {
        $data = $minimal.Clone(); $data.BoxTables = 'codeblock'
        (Resolve-LLConfig -Data $data).BoxTables | Should BeExactly 'CodeBlock'
    }
    It 'rejects unknown settings' {
        $data = $minimal.Clone(); $data.BoxTabels = 'Auto'
        { Resolve-LLConfig -Data $data } | Should Throw "Unknown setting 'BoxTabels'"
    }
    It 'rejects values of the wrong type' {
        $data = $minimal.Clone(); $data.OrderPrefix = 'yes'
        { Resolve-LLConfig -Data $data } | Should Throw 'OrderPrefix'
    }
    It 'rejects values outside the allowed set' {
        $data = $minimal.Clone(); $data.LinkStyle = 'Absolute'
        { Resolve-LLConfig -Data $data } | Should Throw 'LinkStyle'
    }
    It 'rejects relative paths' {
        $data = $minimal.Clone(); $data.VaultPath = '.\vault'
        { Resolve-LLConfig -Data $data } | Should Throw 'absolute path'
    }
    It 'requires Notebooks, WorkPath and VaultPath' {
        { Resolve-LLConfig -Data @{} } | Should Throw 'Missing required setting'
    }
    It 'accepts the example configuration' {
        $example = Import-PowerShellDataFile -LiteralPath (Join-Path $script:LLTestRoot '..\looseleaf.config.example.psd1')
        { Resolve-LLConfig -Data $example } | Should Not Throw
    }
}

Describe 'ConvertTo-LLProcessArgument' {
    It 'leaves simple arguments alone' { ConvertTo-LLProcessArgument 'abc' | Should BeExactly 'abc' }
    It 'quotes spaces' { ConvertTo-LLProcessArgument 'a b' | Should BeExactly '"a b"' }
    It 'escapes quotes' { ConvertTo-LLProcessArgument 'a"b' | Should BeExactly '"a\"b"' }
    It 'doubles trailing backslashes inside quotes' { ConvertTo-LLProcessArgument 'C:\a b\' | Should BeExactly '"C:\a b\\"' }
    It 'quotes empty strings' { ConvertTo-LLProcessArgument '' | Should BeExactly '""' }
}

Describe 'Get-LLPagePlan (DESIGN 6.1)' {
    $config = Resolve-LLConfig -Data @{ Notebooks = @('NB'); WorkPath = 'C:\W'; VaultPath = 'C:\V'; OrderPrefix = 'Always' }
    $notebook = ConvertTo-LLTestManifestObject @{
        name  = 'NB'
        items = @(
            @{ type = 'section'; name = 'Section: One'; order = 1; status = 'ok'; pages = @(
                    (New-LLTestPage 'p1' 'Intro' 1),
                    (New-LLTestPage 'p2' '----- Linux -----' 2),
                    (New-LLTestPage 'p3' 'Permissions' 3 -Level 2),
                    (New-LLTestPage 'p4' 'Empty leaf' 4),
                    (New-LLTestPage 'p5' 'Parent with content' 5),
                    (New-LLTestPage 'p6' 'Child' 6 -Level 2),
                    (New-LLTestPage 'p7' 'Grandchild' 7 -Level 3)
                )
            }
            @{ type = 'sectionGroup'; name = 'Group'; order = 2; items = @(
                    @{ type = 'section'; name = 'Inner'; order = 1; status = 'ok'; pages = @((New-LLTestPage 'p8' 'Deep page' 1)) }
                    @{ type = 'section'; name = 'Secret'; order = 2; status = 'locked'; pages = @() }
                )
            }
        )
    }
    $empty = @{ p2 = $true; p4 = $true }

    Context 'EmptyPages = FolderOnly (default)' {
        $plan = Get-LLPagePlan -Notebook $notebook -Config $config -NotebookDir 'C:\V\NB' -EmptyKeys $empty
        $byKey = @{}; foreach ($e in $plan.Entries) { $byKey[$e.Key] = $e }

        It 'prefixes and cleans section and page names' {
            $byKey.p1.RelDir | Should BeExactly '01-Section - One'
            $byKey.p1.FileName | Should BeExactly '01-Intro.md'
        }
        It 'turns an empty parent page into a folder only' {
            $byKey.p2.FileName | Should BeNullOrEmpty
            $byKey.p3.RelDir | Should BeExactly '01-Section - One\02-Linux'
            $byKey.p3.FileName | Should BeExactly '01-Permissions.md'
        }
        It 'drops an empty page without subpages and renumbers the rest' {
            $byKey.ContainsKey('p4') | Should Be $false
            $byKey.p5.FileName | Should BeExactly '03-Parent with content.md'
        }
        It 'nests subpages two levels deep' {
            $byKey.p6.RelDir | Should BeExactly '01-Section - One\03-Parent with content'
            $byKey.p7.RelDir | Should BeExactly '01-Section - One\03-Parent with content\01-Child'
        }
        It 'handles section groups' {
            $byKey.p8.RelDir | Should BeExactly '02-Group\01-Inner'
        }
        It 'keeps the output of locked sections' {
            $plan.KeepPrefixes -contains '02-Group\02-Secret' | Should Be $true
        }
        It 'records the source path for front matter' {
            $byKey.p6.SourcePath | Should BeExactly 'NB/Section: One/Parent with content'
        }
    }

    Context 'EmptyPages = Skip' {
        $cfg = Resolve-LLConfig -Data @{ Notebooks = @('NB'); WorkPath = 'C:\W'; VaultPath = 'C:\V'; EmptyPages = 'Skip'; OrderPrefix = 'Always' }
        $plan = Get-LLPagePlan -Notebook $notebook -Config $cfg -NotebookDir 'C:\V\NB' -EmptyKeys $empty
        $byKey = @{}; foreach ($e in $plan.Entries) { $byKey[$e.Key] = $e }
        It 'moves subpages of an empty page up one level' {
            $byKey.p3.RelDir | Should BeExactly '01-Section - One'
            $byKey.p3.FileName | Should BeExactly '02-Permissions.md'
        }
    }

    Context 'EmptyPages = Keep, OrderPrefix off' {
        $cfg = Resolve-LLConfig -Data @{ Notebooks = @('NB'); WorkPath = 'C:\W'; VaultPath = 'C:\V'; EmptyPages = 'Keep'; OrderPrefix = $false }
        $nb = ConvertTo-LLTestManifestObject @{ name = 'NB'; items = @(
                @{ type = 'section'; name = 'S'; order = 1; pages = @((New-LLTestPage 'a' 'Introduction' 1), (New-LLTestPage 'b' 'Introduction' 2)) }
            )
        }
        $plan = Get-LLPagePlan -Notebook $nb -Config $cfg -NotebookDir 'C:\V\NB' -EmptyKeys @{ a = $true }
        It 'writes title-only notes for empty pages' {
            $plan.Entries[0].FileName | Should BeExactly 'Introduction.md'
        }
        It 'de-duplicates equal names in one folder' {
            $plan.Entries[1].FileName | Should BeExactly 'Introduction-2.md'
        }
    }

    Context 'OrderPrefix = Auto (default)' {
        $cfg = Resolve-LLConfig -Data @{ Notebooks = @('NB'); WorkPath = 'C:\W'; VaultPath = 'C:\V' }
        $nb = ConvertTo-LLTestManifestObject @{ name = 'NB'; items = @(
                @{ type = 'section'; name = '1 - Basics'; order = 1; pages = @(
                        (New-LLTestPage 'a' 'Zeta' 1),
                        (New-LLTestPage 'b' 'Alpha' 2)
                    )
                }
                @{ type = 'section'; name = '2 - Advanced'; order = 2; pages = @(
                        (New-LLTestPage 'c' '1.9 Nine' 1),
                        (New-LLTestPage 'd' '1.10 Ten' 2)
                    )
                }
                @{ type = 'section'; name = 'Exam'; order = 3; pages = @((New-LLTestPage 'e' 'Only page' 1)) }
            )
        }
        $plan = Get-LLPagePlan -Notebook $nb -Config $cfg -NotebookDir 'C:\V\NB' -EmptyKeys @{}
        $byKey = @{}; foreach ($e in $plan.Entries) { $byKey[$e.Key] = $e }
        It 'keeps names that already sort in OneNote order (numbers by value, then letters)' {
            $byKey.a.RelDir | Should BeExactly '1 - Basics'
            $byKey.e.RelDir | Should BeExactly 'Exam'
            $byKey.e.FileName | Should BeExactly 'Only page.md'
            $byKey.c.FileName | Should BeExactly '1.9 Nine.md'
            $byKey.d.FileName | Should BeExactly '1.10 Ten.md'
        }
        It 'prefixes a folder whose names would sort differently' {
            $byKey.a.FileName | Should BeExactly '01-Zeta.md'
            $byKey.b.FileName | Should BeExactly '02-Alpha.md'
        }
        It 'accepts $true as Always' {
            (Resolve-LLConfig -Data @{ Notebooks = @('NB'); WorkPath = 'C:\W'; VaultPath = 'C:\V'; OrderPrefix = $true }).OrderPrefix | Should BeExactly 'Always'
        }
    }

    Context 'MaxPathLength' {
        $cfg = Resolve-LLConfig -Data @{ Notebooks = @('NB'); WorkPath = 'C:\W'; VaultPath = 'C:\V'; MaxPathLength = 100 }
        $nb = ConvertTo-LLTestManifestObject @{ name = 'NB'; items = @(
                @{ type = 'section'; name = 'S'; order = 1; pages = @((New-LLTestPage 'a' ('Long title ' * 20) 1)) }
            )
        }
        $plan = Get-LLPagePlan -Notebook $nb -Config $cfg -NotebookDir 'C:\V\NB' -EmptyKeys @{}
        It 'shortens names to fit and warns' {
            (Join-Path 'C:\V\NB' (Join-Path $plan.Entries[0].RelDir $plan.Entries[0].FileName)).Length | Should Not BeGreaterThan 100
            $plan.Warnings.Count | Should BeGreaterThan 0
        }
        It 'keeps names unique when shortening makes them equal' {
            $cfg2 = Resolve-LLConfig -Data @{ Notebooks = @('NB'); WorkPath = 'C:\W'; VaultPath = 'C:\V'; MaxPathLength = 100; OrderPrefix = 'Never' }
            $long = 'Web Application Technologies and a Very Long Descriptive Title That Goes On and On - Part '
            $nb2 = ConvertTo-LLTestManifestObject @{ name = 'NB'; items = @(
                    @{ type = 'section'; name = 'S'; order = 1; pages = @((New-LLTestPage 'a' ($long + '1') 1), (New-LLTestPage 'b' ($long + '2') 2)) }
                )
            }
            $p2 = Get-LLPagePlan -Notebook $nb2 -Config $cfg2 -NotebookDir 'C:\V\NB' -EmptyKeys @{}
            $p2.Entries[0].FileName | Should Not Be $p2.Entries[1].FileName
            $p2.Entries[1].FileName | Should Match '-2\.md$'
        }
    }
}

Describe 'Resolve-LLMediaName (DESIGN 6.3)' {
    It 'keeps names stable across builds and never shares a name' {
        $old = @{ 'a|img|01' = 'Sec-Page-01.png' }
        $new = @{}
        $owned = @{ 'sec-page-01.png' = 'a|img|01' }
        Resolve-LLMediaName -RegistryKey 'b|img|01' -Desired 'Sec-Page-01.png' -OldMedia $old -NewMedia $new -Owned $owned | Should BeExactly 'Sec-Page-01-2.png'
        Resolve-LLMediaName -RegistryKey 'a|img|01' -Desired 'Other-01.png' -OldMedia $old -NewMedia $new -Owned $owned | Should BeExactly 'Sec-Page-01.png'
        Resolve-LLMediaName -RegistryKey 'c|img|01' -Desired 'Sec-Page-01.png' -OldMedia $old -NewMedia $new -Owned $owned | Should BeExactly 'Sec-Page-01-3.png'
    }
}

Describe 'Format-LLMediaLink' {
    It 'writes relative embeds with Obsidian width and encoded spaces' {
        Format-LLMediaLink -FileName 'a b(1).png' -Width 800 -LinkPrefix '../media/' -LinkStyle Relative -Embed |
            Should BeExactly '![|800](../media/a%20b%281%29.png)'
    }
    It 'writes wiki embeds' {
        Format-LLMediaLink -FileName 'x.png' -Width 300 -LinkStyle Wiki -Embed | Should BeExactly '![[x.png|300]]'
    }
    It 'writes attachment links' {
        Format-LLMediaLink -FileName 'scan.pdf' -Text 'Scan [v2].pdf' -LinkPrefix 'media/' | Should BeExactly '[Scan \[v2\].pdf](media/scan.pdf)'
    }
    It 'replaces filter placeholders' {
        Resolve-LLMediaPlaceholder -Markdown "a ![|50](LLMEDIA:p-01.png) b ![](LLMEDIA:p-02.png)" -Names @{ 'p-01.png' = 'q-01.png' } -LinkPrefix 'media/' -LinkStyle Relative |
            Should BeExactly 'a ![|50](media/q-01.png) b ![](media/p-02.png)'
    }
}

Describe 'Remove-LLUnneededEscape (rule C7)' {
    It 'unescapes > inside text' { Remove-LLUnneededEscape 'a \>\> b' | Should BeExactly 'a >> b' }
    It 'keeps \> at the start of a line' { Remove-LLUnneededEscape '\> not a quote' | Should BeExactly '\> not a quote' }
    It 'keeps \> at the start of quoted content' { Remove-LLUnneededEscape '> \> x' | Should BeExactly '> \> x' }
    It 'never unescapes <' { Remove-LLUnneededEscape 'ssh \<user\>@host' | Should BeExactly 'ssh \<user>@host' }
    It 'leaves code blocks alone' {
        $fence = '```'
        $block = "$fence`na \> b`n$fence"
        Remove-LLUnneededEscape $block | Should BeExactly $block
    }
    It 'leaves inline code alone' { Remove-LLUnneededEscape 'x `a \> b` y \> z' | Should BeExactly 'x `a \> b` y > z' }
}

Describe 'Format-LLNote' {
    $config = Resolve-LLConfig -Data @{ Notebooks = @('NB'); WorkPath = 'C:\W'; VaultPath = 'C:\V' }
    $note = Format-LLNote -Title 'Say "hi"' -Body "Body`n" -FrontMatter ([ordered]@{ title = 'Say "hi"'; source = 'NB\S' }) -AttachmentLinks @('[a.pdf](media/a.pdf)') -Config $config
    It 'writes escaped YAML front matter' { $note | Should Match '(?m)^title: "Say \\"hi\\""$' }
    It 'escapes backslashes in YAML' { $note | Should Match '(?m)^source: "NB\\\\S"$' }
    It 'adds the title heading' { $note | Should Match '(?m)^# Say "hi"$' }
    It 'lists attachments' { $note | Should Match '(?m)^## Attachments\n\n- \[a\.pdf\]\(media/a\.pdf\)$' }
    It 'uses LF and ends with a newline' {
        $note.Contains("`r") | Should Be $false
        $note.EndsWith("`n") | Should Be $true
    }
}
