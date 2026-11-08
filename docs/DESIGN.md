# Looseleaf design

| | |
| --- | --- |
| **Status** | Draft. Describes intended behaviour before implementation. |
| **Version** | 0.1 (2026-11-08) |
| **Scope** | Architecture, data formats, output layout and behaviour rules. Setting-by-setting reference: [CONFIGURATION.md](CONFIGURATION.md). |

## 1. Goals and non-goals

### Goals

1. Export local OneNote notebooks to Markdown that works in Obsidian without manual cleanup.
2. Never require a cloud service, an account or an upload.
3. Handle large notebooks (500+ pages) reliably: resumable, with per-page failures that don't stop the run, and a clear summary.
4. Produce a predictable, stable vault: the same input and settings give the same file names, image names and order.
5. Be maintainable: documented behaviour, tests for everything that doesn't need OneNote, and no dependency on another project's internals.

### Non-goals

- Syncing back to OneNote, or two-way sync.
- Cross-platform export. Stage 1 needs the OneNote COM API, which exists only on Windows. Stage 2 is plain PowerShell and pandoc and may run elsewhere, but that isn't a target.
- Converting ink and handwriting to text.
- Supporting *OneNote for Windows 10* (the Store app), which has no COM API.

## 2. Architecture

```text
                     +-----------------------------------+
 OneNote desktop --> | Stage 1: Export-LooseleafNotebook | -->  <WorkPath>/
 (COM, same          | preflight -> hierarchy -> docx    |        manifest.json
  privilege level)   +-----------------------------------+        docx/<key>.docx
                                                                  attachments/<key>/...
                                                                  logs/

                     +-----------------------------------+
 <WorkPath>/ ------> | Stage 2: Build-LooseleafVault     | -->  <VaultPath>/<Notebook>/
                     | pandoc + Lua filter -> files      |        NN-Section/NN-Page.md
                     +-----------------------------------+        media/
                                                                  .looseleaf/build-state.json
```

The **manifest** is the contract between the two stages. Stage 2 reads only the manifest and the files it points to, never OneNote. This separation is the main design decision:

- Every Markdown fix can be developed and tested against saved `.docx` files.
- A failed or interrupted export can be resumed without redoing finished pages.
- Page order is known when the export runs and is written to the manifest, so it never has to be reconstructed from console output (as ON2MD's `Repair-PageOrder` did).

A third public command, `Test-LooseleafEnvironment`, runs the preflight checks (§7) without changing anything.

## 3. Stage 1: Export

### 3.1 Connecting to OneNote

- Uses **late-bound COM**: `New-Object -ComObject OneNote.Application`. It does **not** load the `Microsoft.Office.Interop.OneNote` assembly from the GAC. That assembly is why upstream refuses PowerShell 7.1 and later. Late binding should work on 5.1 and 7.x, but 7.x still has to be verified.
- The COM object is released (`[Runtime.InteropServices.Marshal]::ReleaseComObject`) when the run ends, including after errors.

### 3.2 Walking the hierarchy

1. `GetHierarchy('', hsPages)` returns the XML for all open notebooks.
2. Only the notebooks named in `Notebooks` are processed.
3. Items in the recycle bin (`isInRecycleBin` / `isDeletedPages`) are skipped.
4. Section groups and sections are recorded **in document order**, and pages are recorded in order with their `pageLevel` (1–3).
5. Locked sections (`locked="true"`) are recorded with `status: "locked"` and skipped, with a warning.

### 3.3 Exporting a page

For each page:

1. Compute the page **key**: the first 16 hex characters of SHA-1(page ID). It's stable, file-name safe and short.
2. **Skip** the page if `docx/<key>.docx` exists and the manifest from the previous run has the same `modified` timestamp, unless `-Force` is used. This is what makes runs resumable.
3. Count collapsed paragraphs (`one:OE[@collapsed='true']`) and ink (`one:InkDrawing`, `one:InkWord`) in the page XML (`GetPageContent`).
4. If `ExpandCollapsed` is on and the page has collapsed paragraphs, expand them with `UpdatePageContent`, sending only the changed outline elements, and remember them so they can be restored.
5. Publish the page to `docx/<key>.docx.partial` (`Publish(id, path, pfWord)`), then rename it to `<key>.docx`. A crash therefore never leaves a half-written file that looks complete.
6. Copy each inserted file (`one:InsertedFile/@pathCache`) to `attachments/<key>/<safe name>`. This has to happen in stage 1, because OneNote's cache paths are only valid while OneNote is running.
7. If `ExportInkPagesAsPdf` is on and the page contains ink, also publish `pdf/<key>.pdf`.
8. If paragraphs were expanded and `RestoreCollapsed` is on, collapse them again.
9. Update the page's entry in the in-memory manifest.

> [!CAUTION]
> Steps 4 and 8 **write to your notebook**. They change each affected page's last-modified time, and that change syncs. Both are off by default and must be tested on a copy of a notebook before you use them on real notes.

### 3.4 Writing the manifest

The manifest is saved after every N pages and at the end. Each save writes to `manifest.json.tmp` and then renames it over the old file, so an interruption loses at most N pages of progress.

## 4. Manifest format (`manifest.json`, schema version 1)

```json
{
  "schemaVersion": 1,
  "generator": "Looseleaf 0.1.0",
  "exportedAt": "2026-11-08T14:00:00+01:00",
  "notebooks": [
    {
      "id": "{A1B2...}{1}{B0}",
      "name": "My Notebook",
      "items": [
        {
          "type": "sectionGroup",
          "id": "{...}",
          "name": "Projects",
          "order": 1,
          "items": [ ]
        },
        {
          "type": "section",
          "id": "{...}",
          "name": "Website Redesign",
          "order": 2,
          "status": "ok",
          "pages": [
            {
              "id": "{...}{1}{E1...}",
              "key": "3f9a1c2b7d004e11",
              "name": "Kick-off Notes",
              "order": 1,
              "level": 1,
              "created": "2024-01-10T09:12:00Z",
              "modified": "2024-02-02T18:40:00Z",
              "status": "exported",
              "docx": "docx/3f9a1c2b7d004e11.docx",
              "pdf": null,
              "collapsedParagraphs": 0,
              "inkElements": 0,
              "attachments": [
                { "name": "brief.pdf", "file": "attachments/3f9a1c2b7d004e11/brief.pdf" }
              ],
              "error": null
            }
          ]
        }
      ]
    }
  ]
}
```

- Pages are kept as a **flat ordered list with `level`**, exactly as OneNote stores them. Stage 2 works out the subpage nesting from it.
- Page `status` is one of `exported`, `unchanged` (skipped, reused from an earlier run), `failed` (with `error`), or `locked` (inherited from the section).
- All paths are relative to `WorkPath`, so the work folder can be moved.
- A newer `schemaVersion` than the build understands makes stage 2 stop with a clear error.

## 5. Stage 2: Build

### 5.1 For each page (in manifest order)

1. Create a fresh **temporary folder for this page only** (under `%TEMP%\Looseleaf\<run>\<key>`). It's deleted afterwards whether the page succeeds or fails. No images from one page can leak into another (ON2MD bug 2).
2. **Pre-process a copy of the `.docx`** in the temporary folder (rule C2 in §6.5). Runs set in a monospace font get Word's `Verbatim Char` character style. OneNote uses direct font formatting (for example Consolas or Cascadia Code), and pandoc ignores fonts, so without this step code would come out as plain text. The cached `.docx` itself is never changed.
3. Run pandoc:
   `pandoc -f docx -t <PandocFormat> --wrap=none --markdown-headings=atx --lua-filter=filters/looseleaf.lua -M looseleaf-image-prefix=<prefix> -M looseleaf-media-dir=<temp>/media -o <temp>/page.md <temp>/page.docx`
   No `--extract-media`: the filter writes each image itself and links it as `LLMEDIA:<name>`, a placeholder that step 5 replaces with the final link (`LinkStyle`). Code blocks are written as fences by the filter, because pandoc writes code blocks without attributes in indented form.
4. The **Lua filter** (planned; prototype first) applies the content rules in §6.5, working on pandoc's document structure rather than on text:
   - removes OneNote's own title/date/time block at the top of the document, which is the job of the "first 6 lines" removal in upstream and ON2MD
   - renames images in **order of appearance** through `pandoc.mediabag` (`<prefix>-01.png`, `<prefix>-02.png`, …), which fixes ON2MD bug 4
   - converts image sizes to Obsidian `|width` sizing, or drops them (`ImageWidth`)
   - leaves external `http(s)://` images unchanged (ON2MD bug 3)
   - converts single-cell "box" tables, unwraps indentation-only block quotes, turns glyph bullets into lists, and handles tables, underline, highlight and clip-source lines (§6.5)
5. Move the extracted images into the notebook's `media/` folder through the **media registry** (§6.3).
6. Copy attachments to `media/` the same way, and turn references to them into links (ON2MD bug 5).
7. Post-process the text, in a documented and tested order: safe escape cleanup (`CleanEscapes`), removal of empty and space-only paragraphs, collapsing runs of blank lines, and the line-ending style (`NewLine`).
8. Add the YAML front matter and the title heading (§6.4).
9. Write the file as **UTF-8 without BOM** to its final path. If the file was edited since the last build, follow `OnModifiedFile` (§5.2).

### 5.2 Owning the output and cleaning up

- Stage 2 owns `<VaultPath>/<Notebook>/`. `build-state.json` lists every file the build created and the image registry.
- When it rebuilds, files from the previous build that are no longer produced (a page was renamed or deleted in OneNote) are removed. **Files Looseleaf did not create are never touched**, so notes you add yourself in Obsidian are safe.
- **Files you edited after export are protected.** People keep working in the vault (pasting images, fixing text), so `build-state.json` stores a hash of every file as written. If a file's current hash differs, it was edited after the build, and `OnModifiedFile` decides what happens:
  - `Skip` (default): leave it alone and warn
  - `Backup`: rename your version to `<name>.edited-<timestamp>.md`, then write the new one
  - `Overwrite`: replace it

  Stale-file removal follows the same rule, so an edited file is never deleted.
- `-WhatIf` lists every create, overwrite and delete without doing any of them.

## 6. Output layout and naming

### 6.1 Layout

```text
<VaultPath>/
└── My Notebook/                        notebook folder
    ├── media/                          shared images and attachments (MediaScope = Notebook)
    ├── .looseleaf/build-state.json     ignored by Obsidian (dot folder)
    ├── 01-Projects/                    section group
    │   └── 01-Website Redesign/        section
    │       ├── 01-Kick-off Notes.md
    │       ├── 02-Research.md          page with subpages and content…
    │       ├── 02-Research/            …subpages go in a folder named after the parent page
    │       │   ├── 01-Competitors.md
    │       │   └── 02-User Interviews.md
    │       └── 03-Linux/               empty parent page: folder only (EmptyPages = FolderOnly)
    │           └── 01-Permissions.md
    └── 02-Reference/
```

**Empty pages.** Many notebooks use title-only pages as headings or dividers for their subpages (for example `----- Linux -----` or `( Windows )`). In the reference notebook, 91 of 502 pages had no content at all. `EmptyPages` controls them:

- `FolderOnly` (default): an empty page with subpages becomes just a folder, and an empty page without subpages is skipped and logged
- `Keep`: write a title-only note
- `Skip`: skip all empty pages; their subpages move up one level

### 6.2 File and folder names

Applied, in this order, to every notebook, section group, section, page and media name:

1. **Trim decorative runs.** Remove leading and trailing runs of `- = _ * ~ .` and whitespace used as visual dividers, so `----- Linux -----` becomes `Linux`. Brackets around a whole title are kept, so `( Windows )` becomes `(Windows)`.
2. **Readable replacements:** `[` → `(`, `]` → `)`, and `:` becomes a spaced hyphen. So `[ LAB 1.2 ] Scan` becomes `(LAB 1.2) Scan`, and `Demo: Setup` becomes `Demo - Setup`.
3. Replace the remaining characters that are invalid on NTFS (`< > " / \ | ? *` and control characters) or that break Obsidian links (`# ^ |`) with `-`.
4. Collapse runs of whitespace and `-`. Trim leading and trailing whitespace, `-`, and trailing `.`. Remove spaces just inside brackets (`( CTF 1.1 )` becomes `(CTF 1.1)`).
5. Spaces are kept by default (`NameSpaces = Keep`), or replaced with `-` (`Hyphen`). Image names never contain spaces, whatever this setting is.
6. Prefix reserved device names (`CON`, `PRN`, `AUX`, `NUL`, `COM1–9`, `LPT1–9`, with or without an extension) with `_`.
7. An empty result becomes `Untitled`.
8. Shorten each name to `MaxNameLength`. Then, if the full path is longer than `MaxPathLength` (default 250, below Windows' traditional 260-character MAX_PATH), shorten the longest parts and warn.
9. **Order prefix** (`OrderPrefix = $true`): `NN-` is prepended **after** cleaning, using two digits, or three if a folder has more than 99 items. The prefix is only ever added by Looseleaf, never parsed back out of a OneNote name, so a page called `2023-10-05 Meeting` keeps its date (ON2MD bug 9).
10. Two items in the same folder that end up with the same name get `-2`, `-3`, and so on. Common page titles such as `Introduction` often repeat across sections; that's fine, because they're in different folders.

### 6.3 Image names and the media registry

- Format: `<Section>-<Page>-<NN>.<ext>` (`ImageNameStyle = SectionPage`) or `<Page>-<NN>.<ext>` (`Page`). Each part is shortened (`MaxImageLabelLength`) and the whole name is limited to `MaxImageNameLength`.
- The registry in `build-state.json` maps each `(page key, image index)` to a file name.
  - On a rebuild, a page's images keep the names they had.
  - A new name that clashes with a name **owned by another page** gets `-2`, `-3`, and so on. A file is **never overwritten across pages** (ON2MD bug 1).
- Links: `LinkStyle = Relative` (default, portable: `../../media/x.png`) or `Wiki` (`![[x.png]]`, Obsidian-native, which works because names are unique).

### 6.4 Page content header

```markdown
---
title: Kick-off Notes
created: 2024-01-10T09:12:00Z
modified: 2024-02-02T18:40:00Z
onenote-id: "{...}{1}{E1...}"
source: My Notebook/Projects/Website Redesign
---

# Kick-off Notes
```

`FrontMatter` and `TitleHeading` can each be turned off. OneNote's own title/date block is always removed (§5.1), so the title never appears twice (ON2MD bug 6). The heading is always followed by a blank line, so it can't merge with the first line of the page (ON2MD bug 12).

### 6.5 Content rules

These rules come from analysing a real 500-page notebook exported to `.docx` (see §12.1). Each is general, configurable where reasonable, and covered by a made-up test fixture.

| # | Pattern found in OneNote `.docx` | Rule | Setting |
| --- | --- | --- | --- |
| C1 | **Single-cell tables used as boxes** for commands and notes. 1,092 of 1,270 tables had one cell. pandoc writes them as HTML `<table>`, because the cell contains line breaks. | A box holding headings, tables or images is a **layout container** and is unwrapped. Otherwise, with `Auto`: it becomes a fenced **code block** if it's mostly monospace (C2) **or doesn't read like prose**. It becomes an Obsidian **callout** (`> [!note]`) if it contains lists or real links, or if at least half its lines are sentences (8+ words ending in `. ! ?`). Separate boxes stay separate code blocks. | `BoxTables = Auto` (`Auto`, `CodeBlock`, `Callout`, `Unwrap`, `Table`), `BoxCodeThreshold = 0.6` |
| C2 | **Code formatted by font only:** Consolas or Cascadia Code set directly on runs (about 24,000 runs in 148 pages). pandoc ignores fonts. | The pre-processing step (§5.1 step 2) gives monospace runs the `Verbatim Char` style. Inline monospace becomes `` `code` ``, and consecutive monospace paragraphs become one fenced block. | `MonospaceFonts` (list), `DetectCode = $true` |
| C3 | **Multi-cell tables with line breaks or lists in cells**: pipe tables can't hold them. | Line breaks in cells become `<br>`, which Obsidian renders inside pipe tables. Only tables with merged cells or block content that can't be flattened stay as HTML. | `ComplexTables = Html` (`Html`, `Flatten`) |
| C4 | **Indentation turns into block quotes.** OneNote outline indentation becomes Word indentation, which pandoc reads as `>` quotes, often several levels deep. | Block quotes created only by indentation are unwrapped. Nested indentation under a list item becomes nested list items. | `IndentedParagraphs = Unindent` (`Unindent`, `Quote`) |
| C5 | **Typed bullet glyphs** (`●`, `•`, `○`, `▪`, `◦`, `+`, `–`) instead of real lists, usually from pasted content. | Paragraphs starting with a bullet glyph and a space become list items. The nesting level comes from the indentation. | `ConvertGlyphBullets = $true` |
| C6 | **Empty and space-only paragraphs** (non-breaking spaces), trailing `\` hard breaks, and runs of 5–10 blank lines. | Remove them, and keep at most one blank line between blocks. | always on |
| C7 | **Angle-bracket placeholders in text**, such as `smbclient -L <target_ip>`. If the escaping is removed, Obsidian reads `<target_ip>` as an HTML tag, hides it, and can mis-format the text after it. ON2MD's "remove all backslashes" did exactly this. | Never unescape `<`, `>`, `&`. Text inside code is left as written. Placeholders outside code are written as `\<target_ip\>`. Only escapes that change nothing are removed. | `CleanEscapes = Safe` (`Safe`, `None`) |
| C8 | **Underline and highlight**: pandoc writes `[text]{.underline}` and `[text]{.mark}`, which Obsidian shows as literal text. | Underline becomes `<u>text</u>` (or is dropped), and highlight becomes `==text==`. | `Underline = Html` (`Html`, `Drop`) |
| C9 | **Clip-source lines**: OneNote adds `From <https://…>` or `Pasted from <…>` (wording depends on the OneNote language) when you paste from a browser. They often come out as broken HTML. | Detect the pattern, including localised wordings, and write `Source: <url>`, or move the URLs to the front matter `sources:` list. | `ClipSources = Normalize` (`Normalize`, `FrontMatter`, `Keep`) |
| C10 | **Image sizes in inches** (`{width="10.77in"}`, and ON2MD alt text such as `9.91667x2.19792`). | Convert to pixels at 96 DPI, capped at `MaxImageWidth`, and write Obsidian's `![\|800](…)`. Alt text stays empty or comes from OneNote's alt text. | `ImageWidth = Obsidian` (`Obsidian`, `Drop`), `MaxImageWidth = 1000` |
| C11 | **Web-pasted HTML** inside cells (`<ul class="incremental">`, `<a href>`, `<strong>`), which ON2MD left behind when it stripped table tags. | Handled at the document-structure level by C1 and C3. No HTML is ever produced by deleting tags with regex. | — |
| C12 | **Multiple `.docx` files for one page** (the page was edited between runs). | The cache is keyed by page ID only (§3.3), so the newer export replaces the older one. | — |

## 7. Preflight checks (`Test-LooseleafEnvironment`, also run at the start of stage 1)

| Check | Result | Message hint |
| --- | --- | --- |
| Running on Windows | **Error** | Stage 1 needs Windows. |
| `pandoc` on `PATH`, version ≥ 3.0 | **Error** (stage 2) | Install pandoc for the current user or all users and restart the shell. |
| OneNote desktop COM class registered | **Error** | OneNote for Windows 10 (Store) has no COM API; install OneNote desktop. |
| PowerShell and `ONENOTE.EXE` at the same privilege level | **Error** | Run both normally (not as Administrator). |
| OneNote running | **Warning** | It will be started automatically; keep it open. |
| Requested notebooks are open | **Error** | Open the notebook in OneNote first. |
| Locked sections | **Warning** | Unlock them in OneNote or they'll be skipped. |
| Pages with collapsed paragraphs | **Warning** | Content would be missing; use `ExpandCollapsed`. |
| Pages with ink | **Warning** | Ink is lost; consider `ExportInkPagesAsPdf`. |
| `WorkPath` / `VaultPath` writable, paths absolute | **Error** | |
| Long paths likely to exceed `MaxPathLength` | **Warning** | |

## 8. Error handling, logging and exit codes

- **Per-page isolation:** a page that fails is recorded (`status: failed`, `error`) and the run continues. `-StopOnError` aborts at the first failure instead.
- **End-of-run summary:** pages exported, unchanged, failed and locked, plus the failed pages with their reasons. Saved to `logs/<timestamp>-summary.json`.
- **Logging:** a plain-text UTF-8 log in `logs/<timestamp>-<stage>.log`, written by Looseleaf itself. You never need `*>` redirection or copy-pasted console output.
- **Streams:** progress via `Write-Progress`, details via `-Verbose`, problems via `Write-Warning` / `Write-Error`. `Write-Host` is used only for the final summary.
- **Locked files and antivirus:** notes about security topics (exploits, reverse shells, payloads) can make antivirus software lock or quarantine `.docx` and `.md` files while they're being written. This was observed in the reference data. Sharing violations are retried with back-off (3 attempts). If the file is still blocked, the page fails with the hint "file locked — possibly by antivirus; consider excluding WorkPath and VaultPath".
- **Exit codes** when run as a script: `0` = success, `1` = finished but some pages failed, `2` = preflight or configuration error (nothing was done).

## 9. Configuration

- The configuration is a **PowerShell data file** (`.psd1`), loaded with `Import-PowerShellDataFile`, which runs no code. ON2MD's `config.ps1` was executed as a script with `Invoke-Expression`.
- The settings are validated against a schema in the code (types, allowed values, ranges, absolute paths). Unknown keys stop the run, which catches typos.
- `-WhatIf` replaces upstream's `$dryRun` setting.
- Every setting is listed in [CONFIGURATION.md](CONFIGURATION.md).

## 10. Compatibility

- **Windows PowerShell 5.1** is the baseline: every feature must work there. **PowerShell 7.x** is best-effort and tested when practical.
- PowerShell 5.1 differences to keep in mind:
  - `New-Item` has no `-LiteralPath`. Use `[IO.Directory]::CreateDirectory()`.
  - Use `[IO.File]::WriteAllText` with an explicit `UTF8Encoding($false)`. `Set-Content -Encoding UTF8` writes a BOM.
  - `>` and `*>` redirection write UTF-16LE (this is why the log is written by Looseleaf itself).
- pandoc 3.x. The minimum version is checked by the preflight.

## 11. Testing strategy

| Layer | How | Needs OneNote? |
| --- | --- | --- |
| Name cleanup, order prefixes, path shortening | Pester unit tests with table-driven cases | No |
| Media registry (clashes, stability across rebuilds) | Pester, in-memory | No |
| Manifest reading and writing, schema version | Pester against `tests/fixtures/*.json` | No |
| Lua filter and content rules C1–C12 | Pester runs pandoc on small, made-up `tests/fixtures/*.docx` files (one per rule) and compares the result with the expected `.md` | No (needs pandoc) |
| `.docx` pre-processing (monospace) | Pester on fixture `.docx` files with Consolas/Cascadia runs | No |
| Protection of edited files | Pester: build, modify a file, rebuild, check `OnModifiedFile` behaviour | No |
| Full stage 2 build | Fixture manifest + docx → temp vault, checking the file tree and contents | No |
| Stage 1 | Manual checklist against a dedicated **test notebook**: see [TEST-NOTEBOOK.md](TEST-NOTEBOOK.md) | Yes |

Fixtures are always **made up**. Personal notebooks are only ever used for local comparison runs and are never committed.

## 12. Lessons from ON2MD

ON2MD was the wrapper layered over ConvertOneNote2MarkDown that Looseleaf replaces. Its known defects and the design decisions that prevent each one:

| # | ON2MD defect | Prevented by |
| --- | --- | --- |
| 1 | Image names shortened to 50 characters could clash across pages, and `Move-Item -Force` overwrote images silently | Media registry with ownership; never overwrite across pages (§6.3) |
| 2 | pandoc extracted into the shared `media/` folder, so leftover files went to the wrong page | Temporary folder per page (§5.1) |
| 3 | External image URLs were rewritten to `media/` | Lua filter works on the document structure, not regex, and skips `http(s)` (§5.1) |
| 4 | Images numbered by file-name sort (`image10` before `image2`) | Numbered in order of appearance in the document (§5.1) |
| 5 | Attachments copied but never linked | Attachments handled explicitly (§5.1 step 5) |
| 6 | Duplicate title when the timestamp header was on | OneNote header always removed; one header source (§6.4) |
| 7 | `dryRun` failed on every page | `-WhatIf` / `ShouldProcess` throughout |
| 8 | PDF export silently ignored | `ExportInkPagesAsPdf` implemented in stage 1 |
| 9 | Page-order repair removed leading digits from real names | Order recorded in the manifest; prefixes never parsed back (§6.2) |
| 10 | Page order recovered from a UTF-16 console log | Manifest written by stage 1 |
| 11 | Failures swallowed; exit code 0 | Summary + exit codes (§8) |
| 12 | With the timestamp header off, upstream's heading had no newline after it, so the title merged with the first line (`## DNS ReconDNSRecon…`) | Heading always followed by a blank line (§6.4) |
| 13 | Removing all `\` escapes turned `\<target_ip\>` into HTML tags that Obsidian hides | `CleanEscapes = Safe` (rule C7) |
| 14 | Stripping `<table>`/`<tr>`/`<td>` with regex left about 4,600 stray `<strong>`, `<li>` and `<a>` tags | Box and table rules on the document structure (C1, C3, C11) |
| — | Replaced upstream functions at runtime; tied to PowerShell 5.1 by the GAC Interop assembly | Standalone code, late-bound COM (§3.1) |
| — | Config executed as code | `.psd1` data file + schema validation (§9) |
| — | Mixed encodings, duplicate files, no history | `.editorconfig`, `.gitattributes`, git from day one |

### 12.1 Reference data analysis (2026-11-08)

The content rules in §6.5 come from comparing four versions of one personal notebook (5 sections, about 324 MB of `.one` files, 502 pages). The notebook itself isn't part of the repository.

| Version | Pages | Images | Observations |
| --- | --- | --- | --- |
| `.docx` cache (ON2MD stage 1) | 502 | 1,725 | 91 empty pages; 1 page exported twice; 251 pages contain tables |
| ON2MD Markdown output | 410 | 1,726 | Empty pages dropped or moved by hand; about 4,600 leftover HTML tags; heading merged into the first line |
| Hand-cleaned version (regex in an editor) | 480 | 1,687 | Re-converted with a different pandoc format (`{.underline}` spans); images added afterwards in Obsidian (`![[Pasted image …]]`) |
| Obsidian Importer vault | 159 (×2 runs) | 439 | Only 3 of 5 sections, consistent with Importer failing on large notebooks; about 4,500 `<br>` tags and 2,000 escaped characters |

### 12.2 Prototype results (2026-11-08)

The conversion core was run on all 502 reference pages with `tools/Invoke-LLPrototype.ps1`: the C2 pre-processing step, pandoc 3.9 and `filters/looseleaf.lua`. It took about 4 minutes on Windows PowerShell 5.1.

| Measure | ON2MD output | Prototype |
| --- | --- | --- |
| Pages converted | 501 (+ manual regex fixes) | 501. One page failed because antivirus locked its output file (§8). |
| Images | 1,726 | 1,725, numbered in order of appearance |
| Code blocks | almost none; commands left in HTML/quotes | 1,264 (largest: 328 lines of scan output) |
| Callouts | — | 2 (both genuine prose boxes) |
| Stray HTML tags outside tables | about 4,600 | 0 |
| HTML tables (merged cells / block content) | — | 26, kept on purpose (C3) |
| `{.underline}`/`{.mark}` spans, typed bullets, indentation quotes | many | 0 |

Known gaps:

- Glyph-bullet lists come out flat, because the nesting level is lost when indentation quotes are unwrapped.
- Windows UNC paths (`\\host\share`) inside command boxes are auto-linked by pandoc as `file://`.
- About 100 placeholders in prose are written escaped (`\<enter\>`), which is correct but hard to read in source mode.

## 13. Open questions

### Resolved by the prototype

- ~~Lua filter versus PowerShell post-processing~~: the filter writes images itself, without `--extract-media`, and uses placeholders for links (§5.1). This is simpler than renaming through `pandoc.mediabag` and doesn't depend on pandoc's extraction behaviour.
- ~~Monospace detection~~: pandoc 3.9 reads `Verbatim Char` runs as inline code and `Source Code` paragraphs as code blocks. For box classification, the monospace share alone wasn't enough: in the reference notebook, 512 command boxes used Calibri. The "reads like prose" test (C1) handles this.

### Still open

1. **Nested glyph lists:** recover the nesting level from indentation before unwrapping quotes (C4 + C5).
2. **Callout heuristics:** the 8-word sentence rule may need tuning for languages with long compound words. Make it configurable?
3. **Partial `UpdatePageContent`:** is sending only the changed `OE` elements safe on pages with ink or embedded files? Test on a copy of a notebook.
4. **Subpage folders:** keep the parent page as `NN-Page.md` next to an `NN-Page/` folder (as planned), or move the parent inside its own folder as a folder note?
5. **Incremental stage 2:** a full rebuild is the default. Would content hashing to skip unchanged pages be worth the extra complexity for 500+ page notebooks?
6. **Callout type for non-code boxes:** use one fixed `[!note]` for all, or pick the type from the box's background colour?
