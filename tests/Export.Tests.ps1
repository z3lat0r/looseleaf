# Unit tests for stage 1 parts that need no OneNote: hierarchy parsing, page keys, page inspection.
# The OneNote calls themselves are checked manually (docs/TEST-NOTEBOOK.md).

. "$PSScriptRoot\TestHelpers.ps1"

$ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
$hierarchy = @"
<?xml version="1.0"?>
<one:Notebooks xmlns:one="$ns">
  <one:Notebook name="Work Notes" nickname="Work" ID="{NB1}" path="C:\Notebooks\Work Notes\">
    <one:Section name="Inbox" ID="{S1}">
      <one:Page ID="{P1}" name="First" dateTime="2024-01-01T10:00:00.000Z" lastModifiedTime="2024-01-02T10:00:00.000Z" pageLevel="1" />
      <one:Page ID="{P2}" name="Sub" dateTime="2024-01-01T10:00:00.000Z" lastModifiedTime="2024-01-02T10:00:00.000Z" pageLevel="2" />
      <one:Page ID="{P3}" name="Deleted" pageLevel="1" isInRecycleBin="true" />
    </one:Section>
    <one:SectionGroup name="Projects" ID="{G1}">
      <one:Section name="Alpha" ID="{S2}">
        <one:Page ID="{P4}" name="Plan" pageLevel="1" />
      </one:Section>
      <one:Section name="Secret" ID="{S3}" locked="true" encrypted="true" />
    </one:SectionGroup>
    <one:SectionGroup name="OneNote_RecycleBin" ID="{RB}" isRecycleBin="true">
      <one:Section name="Deleted Pages" ID="{S4}" isInRecycleBin="true" isDeletedPages="true">
        <one:Page ID="{P5}" name="Gone" pageLevel="1" isInRecycleBin="true" />
      </one:Section>
    </one:SectionGroup>
  </one:Notebook>
  <one:Notebook name="Other" ID="{NB2}" />
</one:Notebooks>
"@

Describe 'ConvertFrom-LLHierarchy (DESIGN 3.2)' {
    $parsed = ConvertFrom-LLHierarchy -Xml $hierarchy -Notebooks @('Work', 'Missing Book')
    $nb = $parsed.Notebooks[0]

    It 'selects notebooks by display name and reports missing ones' {
        $parsed.Notebooks.Count | Should Be 1
        $nb.name | Should BeExactly 'Work'
        $parsed.Missing | Should Be 'Missing Book'
    }
    It 'also matches the folder name' {
        (ConvertFrom-LLHierarchy -Xml $hierarchy -Notebooks @('Work Notes')).Notebooks.Count | Should Be 1
    }
    It 'keeps sections and section groups in order and skips the recycle bin' {
        $nb.items.Count | Should Be 2
        $nb.items[0].type | Should Be 'section'
        $nb.items[1].type | Should Be 'sectionGroup'
        $nb.items[1].name | Should Be 'Projects'
    }
    It 'records pages in order with level and dates, skipping deleted pages' {
        $pages = $nb.items[0].pages
        $pages.Count | Should Be 2
        $pages[1].name | Should Be 'Sub'
        $pages[1].level | Should Be 2
        $pages[1].order | Should Be 2
        $pages[0].modified | Should Be '2024-01-02T10:00:00.000Z'
    }
    It 'marks locked sections' {
        $nb.items[1].items[1].status | Should Be 'locked'
        $nb.items[1].items[0].status | Should Be 'ok'
    }
    It 'gives pages a stable 16-character key' {
        $nb.items[0].pages[0].key | Should Match '^[0-9a-f]{16}$'
        $nb.items[0].pages[0].key | Should Be (Get-LLPageKey -PageId '{P1}')
    }
    It 'produces a manifest that stage 2 can read' {
        $manifest = [ordered]@{ schemaVersion = 1; notebooks = $parsed.Notebooks }
        $roundTrip = ($manifest | ConvertTo-Json -Depth 32) | ConvertFrom-Json
        $pages = New-Object System.Collections.Generic.List[object]
        Get-LLNotebookPage -Items @($roundTrip.notebooks[0].items) -SectionName '' -Into $pages
        $pages.Count | Should Be 3 # locked section contributes nothing
    }
}

Describe 'Measure-LLNotebook' {
    It 'counts sections, pages and locked sections' {
        $nb = (ConvertFrom-LLHierarchy -Xml $hierarchy -Notebooks @('Work')).Notebooks[0]
        $stats = Measure-LLNotebook -Items $nb.items
        $stats.Sections | Should Be 3
        $stats.Pages | Should Be 3
        $stats.Locked[0] | Should Be 'Secret'
    }
}

Describe 'Get-LLPageContentInfo' {
    $page = @"
<one:Page xmlns:one="$ns" ID="{P1}" name="First">
  <one:Outline><one:OEChildren>
    <one:OE collapsed="true"><one:T><![CDATA[Heading]]></one:T>
      <one:OEChildren><one:OE><one:T><![CDATA[hidden]]></one:T></one:OE></one:OEChildren>
    </one:OE>
    <one:OE><one:InsertedFile pathCache="C:\cache\a.tmp" preferredName="report.pdf" /></one:OE>
  </one:OEChildren></one:Outline>
  <one:InkDrawing><one:Position x="1" y="1" /></one:InkDrawing>
</one:Page>
"@
    $info = Get-LLPageContentInfo -Xml $page
    It 'counts collapsed paragraphs' { $info.CollapsedParagraphs | Should Be 1 }
    It 'counts ink' { $info.InkElements | Should Be 1 }
    It 'lists inserted files' {
        $info.Attachments[0].Name | Should Be 'report.pdf'
        $info.Attachments[0].PathCache | Should Be 'C:\cache\a.tmp'
    }
}
