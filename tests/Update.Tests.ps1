# End-to-end tests for Update-LooseleafVault (stage 2) on made-up pages. Needs pandoc.
# Run in Windows PowerShell:  Invoke-Pester .\tests

. "$PSScriptRoot\TestHelpers.ps1"

Describe 'Update-LooseleafVault' {
    if (-not $script:LLHasPandoc) {
        It 'needs pandoc' -Skip { }
        return
    }

    # Not $TestDrive: Pester 3/4 deletes files created inside a Context when it ends,
    # and these contexts build on each other.
    $root = Join-Path ([IO.Path]::GetTempPath()) ('ll-test-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $work = Join-Path $root 'work'
    $vault = Join-Path $root 'vault'
    $docx = Join-Path $work 'docx'
    [void][IO.Directory]::CreateDirectory($docx)
    [void][IO.Directory]::CreateDirectory((Join-Path $work 'attachments\p1'))
    [IO.File]::WriteAllText((Join-Path $work 'attachments\p1\notes.txt'), 'attached')

    $box = "| nmap -sV 10.0.0.1 |`n|---|"
    New-LLTestDocx -Path (Join-Path $docx 'p1.docx') -Title 'Page One' -Body "Some text.`n`n$box`n`n- item" -WithImage
    New-LLTestDocx -Path (Join-Path $docx 'p2.docx') -Title '----- Linux -----'
    New-LLTestDocx -Path (Join-Path $docx 'p3.docx') -Title 'Child' -Body 'Child text.'
    New-LLTestDocx -Path (Join-Path $docx 'p4.docx') -Title 'Demo: Colon' -Body 'Use ssh \<user\>@host' # escaped so the fixture contains literal text

    $p1 = New-LLTestPage 'p1' 'Page One' 1
    $p1.attachments = @([pscustomobject]@{ name = 'notes.txt'; file = 'attachments/p1/notes.txt' })
    $manifest = @{
        schemaVersion = 1
        notebooks     = @(@{
                name  = 'My Notebook'
                items = @(
                    @{ type = 'section'; name = 'Section A'; order = 1; status = 'ok'; pages = @(
                            $p1,
                            (New-LLTestPage 'p2' '----- Linux -----' 2),
                            (New-LLTestPage 'p3' 'Child' 3 -Level 2)
                        )
                    }
                    @{ type = 'sectionGroup'; name = 'Group'; order = 2; items = @(
                            @{ type = 'section'; name = 'Section B'; order = 1; status = 'ok'; pages = @((New-LLTestPage 'p4' 'Demo: Colon' 1)) }
                        )
                    }
                )
            })
    }
    $manifestPath = Join-Path $work 'manifest.json'
    function Save-Manifest { [IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 20), (New-Object System.Text.UTF8Encoding $false)) }
    Save-Manifest

    $configPath = Join-Path $root 'looseleaf.config.psd1'
    [IO.File]::WriteAllText($configPath, "@{ Notebooks = @('My Notebook'); WorkPath = '$work'; VaultPath = '$vault'; OrderPrefix = 'Always' }")

    $nb = Join-Path $vault 'My Notebook'
    $pageOne = Join-Path $nb '01-Section A\01-Page One.md'
    $child = Join-Path $nb '01-Section A\02-Linux\01-Child.md'
    $colon = Join-Path $nb '02-Group\01-Section B\01-Demo - Colon.md'

    Context 'WhatIf' {
        $null = Update-LooseleafVault -ConfigPath $configPath -WhatIf
        It 'changes nothing' { Test-Path -LiteralPath $vault | Should Be $false }
    }

    Context 'first build' {
        $summary = Update-LooseleafVault -ConfigPath $configPath
        $text = [IO.File]::ReadAllText($pageOne)

        It 'reports no failures' { $summary.Failed | Should Be 0 }
        It 'mirrors the OneNote hierarchy with order prefixes' {
            $pageOne | Should Exist
            $child | Should Exist
            $colon | Should Exist
        }
        It 'turns the empty divider page into a folder only' {
            Join-Path $nb '01-Section A\02-Linux.md' | Should Not Exist
        }
        It 'adds front matter and the title heading' {
            $text | Should Match '(?m)^title: "Page One"$'
            $text | Should Match '(?m)^source: "My Notebook/Section A"$'
            $text | Should Match '(?m)^# Page One$'
        }
        It 'turns a one-cell box into a code block' {
            $text | Should Match '(?m)^```\r?\nnmap -sV 10\.0\.0\.1\r?\n```$'
        }
        It 'writes the image to the shared media folder and links it relatively' {
            Join-Path $nb 'media\Section-A-Page-One-01.png' | Should Exist
            $text | Should Match '!\[(\|\d+)?\]\(\.\./media/Section-A-Page-One-01\.png\)'
        }
        It 'copies and links attachments' {
            Join-Path $nb 'media\notes.txt' | Should Exist
            $text | Should Match '(?m)^- \[notes\.txt\]\(\.\./media/notes\.txt\)$'
        }
        It 'keeps placeholders in text visible (escaped)' {
            [IO.File]::ReadAllText($colon) | Should Match 'ssh \\<user>@host'
        }
        It 'saves the build state' {
            Join-Path $nb '.looseleaf\build-state.json' | Should Exist
        }
    }

    Context 'rebuild without changes' {
        $summary = Update-LooseleafVault -ConfigPath $configPath
        It 'rewrites nothing' {
            $summary.Created | Should Be 0
            $summary.Updated | Should Be 0
            $summary.Unchanged | Should Be 3
        }
    }

    Context 'note edited in Obsidian' {
        [IO.File]::AppendAllText($pageOne, "`nMy own addition`n")
        $summary = Update-LooseleafVault -ConfigPath $configPath
        It 'is kept (OnModifiedFile = Skip)' {
            $summary.SkippedModified | Should Be 1
            [IO.File]::ReadAllText($pageOne) | Should Match 'My own addition'
        }
    }

    Context 'page deleted in OneNote' {
        $manifest.notebooks[0].items[0].pages = @($p1, (New-LLTestPage 'p2' '----- Linux -----' 2))
        Save-Manifest
        $summary = Update-LooseleafVault -ConfigPath $configPath
        It 'removes its note and the folder left empty' {
            $summary.Removed | Should Be 1
            $child | Should Not Exist
            Join-Path $nb '01-Section A\02-Linux' | Should Not Exist
        }
        It 'keeps the edited note' {
            [IO.File]::ReadAllText($pageOne) | Should Match 'My own addition'
        }
    }

    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
