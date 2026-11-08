# Configuration

> [!NOTE]
> **Draft.** These settings follow [DESIGN.md](DESIGN.md) and may change before version 1.0.

Looseleaf reads a **PowerShell data file** (`.psd1`) passed with `-ConfigPath`. Data files cannot run code. Start from [looseleaf.config.example.psd1](../looseleaf.config.example.psd1):

```powershell
Copy-Item .\looseleaf.config.example.psd1 .\looseleaf.config.psd1
```

`looseleaf.config.psd1` is git-ignored because it contains your local paths and notebook names.

## Validation rules

- Unknown keys stop the run with an error, which catches typos.
- Missing keys take the default shown below.
- Paths must be absolute.
- Values outside the allowed set or range stop the run.

## Paths

| Setting | Type | Default | Description |
| --- | --- | --- | --- |
| `Notebooks` | string[] | *(required)* | Notebook names as shown in OneNote. Each must be open. |
| `WorkPath` | string | *(required)* | Stage 1 output folder (`manifest.json`, `docx/`, `attachments/`, `pdf/`, `logs/`). Keep it outside the vault, or inside a dot folder, so Obsidian doesn't index `.docx` files. |
| `VaultPath` | string | *(required)* | Obsidian vault root. Each notebook is built into `<VaultPath>\<Notebook>\`. |

## Stage 1: export

| Setting | Type | Default | Description |
| --- | --- | --- | --- |
| `ExpandCollapsed` | bool | `$false` | Expands collapsed paragraphs before publishing. Without it, OneNote's Word export leaves out the hidden content. **This writes to your notebook** and changes the page's modified time. Replaces the Onetastic macro. |
| `RestoreCollapsed` | bool | `$true` | Collapses the paragraphs again after export. Only applies when `ExpandCollapsed` is on. |
| `ExportInkPagesAsPdf` | bool | `$false` | Also publishes `pdf/<key>.pdf` for pages containing ink, which the Markdown conversion loses. |

## Stage 2: layout and naming

| Setting | Type | Default | Allowed | Description |
| --- | --- | --- | --- | --- |
| `OrderPrefix` | bool | `$true` | | Adds `NN-` to file and folder names so Obsidian sorts them in OneNote order. |
| `MediaScope` | string | `Notebook` | `Notebook`, `Folder` | One shared `media/` folder per notebook, or a `media/` folder next to each note. |
| `ImageNameStyle` | string | `SectionPage` | `SectionPage`, `Page` | `<Section>-<Page>-NN.ext` or `<Page>-NN.ext`. Names are made unique either way. |
| `MaxImageLabelLength` | int | `40` | 8–80 | Maximum length of each section or page label inside an image name. |
| `MaxImageNameLength` | int | `80` | 24–200 | Maximum length of the whole image file name, extension included. |
| `LinkStyle` | string | `Relative` | `Relative`, `Wiki` | `![](../../media/x.png)`, which works in any Markdown viewer, or `![[x.png]]`, Obsidian's own embed syntax. |
| `MaxNameLength` | int | `120` | 16–255 | Maximum length of a single file or folder name. |
| `MaxPathLength` | int | `250` | 100–32000 | Full paths longer than this are shortened, with a warning. Raise it only if Windows long paths are enabled. |
| `NameSpaces` | string | `Keep` | `Keep`, `Hyphen` | Keep spaces in file and folder names, or replace them with `-`. Image names never contain spaces. |
| `EmptyPages` | string | `FolderOnly` | `FolderOnly`, `Keep`, `Skip` | Title-only pages (headings and dividers): just a folder for their subpages, a title-only note, or skipped (subpages move up one level). See [DESIGN.md §6.1](DESIGN.md#61-layout). |
| `OnModifiedFile` | string | `Skip` | `Skip`, `Backup`, `Overwrite` | What a rebuild does with a note you edited after the last build. |

## Stage 2: content

All rules are described in [DESIGN.md §6.5](DESIGN.md#65-content-rules) (C1–C12).

| Setting | Type | Default | Allowed | Description |
| --- | --- | --- | --- | --- |
| `FrontMatter` | bool | `$true` | | Adds YAML front matter (`title`, `created`, `modified`, `onenote-id`, `source`). |
| `TitleHeading` | bool | `$true` | | Adds `# <Page title>` below the front matter. |
| `PandocFormat` | string | `markdown-simple_tables-multiline_tables-grid_tables-smart+pipe_tables` | any pandoc writer | pandoc output format and extensions (`-t`). `-smart` stops pandoc writing `\"` for quotes. |
| `DetectCode` | bool | `$true` | | Turns text in a monospace font into inline code or code blocks (C2). |
| `MonospaceFonts` | string[] | `Consolas`, `Cascadia Code`, `Cascadia Mono`, `Courier New`, `Lucida Console`, `Source Code Pro`, `Fira Code`, `JetBrains Mono` | | Fonts treated as code. Matched case-insensitively; hyphens and spaces are ignored. |
| `BoxTables` | string | `Auto` | `Auto`, `CodeBlock`, `Callout`, `Unwrap`, `Table` | Single-cell tables (C1). Boxes holding headings, tables or images are always unwrapped. `Auto`: a code block if at least `BoxCodeThreshold` of the text is monospace, or if the box doesn't read like prose; otherwise a callout. |
| `BoxCodeThreshold` | double | `0.6` | 0.0–1.0 | Share of monospace characters that makes a box a code block. |
| `ComplexTables` | string | `Html` | `Html` | Tables that pipe tables can't represent (merged cells, block content) (C3). `Flatten` is planned. |
| `IndentedParagraphs` | string | `Unindent` | `Unindent`, `Quote` | Indented OneNote paragraphs: unwrap the block quotes pandoc creates, or keep them as quotes (C4). |
| `ConvertGlyphBullets` | bool | `$true` | | Turns paragraphs starting with `●`, `•`, `○`, `▪`, `◦`, `+` or `–` into list items (C5). |
| `CleanEscapes` | string | `Safe` | `Safe`, `None` | Removes only backslash escapes that change nothing. `<`, `>` and `&` are never unescaped (C7). |
| `Underline` | string | `Html` | `Html`, `Drop` | Underlined text: `<u>text</u>` or plain text (C8). Highlighted text always becomes `==text==`. |
| `ClipSources` | string | `Normalize` | `Normalize`, `Keep` | OneNote's `From <url>` / `Pasted from <url>` lines (C9). `FrontMatter` is planned. |
| `ImageWidth` | string | `Obsidian` | `Obsidian`, `Drop` | Write OneNote image sizes as Obsidian width (`![\|800](…)`) or drop them (C10). |
| `MaxImageWidth` | int | `1000` | 100–4000 | Upper limit for the image width in pixels. |
| `NewLine` | string | `LF` | `LF`, `CRLF` | Line endings of the generated Markdown. |

## Command-line switches (not in the config file)

| Switch | Commands | Description |
| --- | --- | --- |
| `-WhatIf` | all | Shows what would be created, changed or deleted, without doing it. Replaces upstream's `$dryRun`. |
| `-Force` | Export | Re-publishes every page, even unchanged ones. |
| `-StopOnError` | Export, Build | Stops at the first failed page instead of continuing. |
| `-Verbose` | all | Shows detailed progress. |

## Mapping from ON2MD / ConvertOneNote2MarkDown

| Old setting | Looseleaf |
| --- | --- |
| `$dryRun` | `-WhatIf` |
| `$notesdestpath` | `VaultPath` |
| `$targetNotebook` | `Notebooks` |
| `$usedocx`, `$keepdocx` | Automatic: `.docx` files are kept in `WorkPath` and reused when the page is unchanged; `-Force` re-exports |
| `$docxNamingConvention` | Fixed: `<key>.docx` (hash of the page ID) |
| `$prefixFolders` | Fixed: subpages go in a folder named after the parent page |
| `$medialocation` | `MediaScope` |
| `$conversion` | `PandocFormat` |
| `$headerTimestampEnabled` | `FrontMatter`, `TitleHeading` |
| `$keepspaces` | Always on: empty paragraphs and extra blank lines are removed (rule C6) |
| `$keepescape` | `CleanEscapes` (the "remove all" option was dropped because it turned `<placeholders>` into HTML) |
| `$newlineCharacter` | `NewLine` |
| `$exportPdf` | `ExportInkPagesAsPdf` (ink pages only) |
| `$mdFileNameAndFolderNameMaxLength` | `MaxNameLength` |
| `$mediaFileNameStyle` | `ImageNameStyle` (legacy style 3 dropped) |
| `$mediaSectionMaxLength`, `$mediaPageMaxLength` | `MaxImageLabelLength` |
| `$mediaFileNameMaxLength` | `MaxImageNameLength` |
| `$mediaLinkStyle` | `LinkStyle` (vault-root `media/` replaced by `Wiki`) |
| `$stripHtmlAfterConversion` | Always on, done by the Lua filter on the document structure; see `BoxTables` and `ComplexTables` |
