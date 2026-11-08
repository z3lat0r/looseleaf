# Changelog

All notable changes to Looseleaf are recorded here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- **Stage 1: `Export-LooseleafNotebook`.** Exports the configured notebooks through late-bound COM (no Office Interop assembly):
  - reads the hierarchy in OneNote order, skips the recycle bin, and marks locked sections
  - publishes each page to `docx\<key>.docx` through a temporary name, and resumes by skipping unchanged pages (`-Force` re-exports all)
  - copies attachments, counts collapsed paragraphs and ink, and can publish ink pages as PDF (`ExportInkPagesAsPdf`), which stage 2 links
  - saves the manifest every 25 pages
- **Preflight: `Test-LooseleafEnvironment`.** Checks Windows, PowerShell, pandoc, the configuration, the paths, the OneNote COM registration, whether OneNote runs at the same privilege level (token check), whether the notebooks are open, and locked sections.
- **Stage 2: `Update-LooseleafVault`.** Builds the vault from `<WorkPath>\manifest.json`:
  - mirrors the OneNote hierarchy (section groups, sections, subpages as folders) with `NN-` order prefixes, applies the naming rules and `EmptyPages`, and shortens names to fit `MaxPathLength`
  - media registry: stable, unique image and attachment names
  - front matter, title heading and an attachments list
  - leaves notes edited after the last build alone (`OnModifiedFile`) and removes stale output
  - retries files locked by antivirus; supports `-WhatIf` and `-StopOnError`; writes logs to `<WorkPath>\logs`
- Configuration loading with a schema: unknown keys, wrong types and relative paths are rejected.
- Pester tests (83) for the **built-in Pester 3.4**: unit tests, end-to-end build tests on made-up pages, and stage 1 parsing tests on sample OneNote XML.
- Project scaffold: module layout, configuration template, lint settings, and repository conventions (`.editorconfig`, `.gitattributes`, `.gitignore`).
- Documentation: README, design document, configuration reference, and troubleshooting guide.
- Design: content rules C1–C12 (box tables, monospace detection, indentation, glyph bullets, safe escaping, underline and highlight, clip sources, image sizes), handling of empty pages, protection of edited files, and retries for locked files. Derived from analysing a 500-page reference notebook.
- Stage 1 test-notebook checklist (`docs/TEST-NOTEBOOK.md`).
- Conversion core prototype:
  - `ConvertTo-LLCodeMarkedDocx` marks monospace text with Word code styles (rule C2).
  - `filters/looseleaf.lua` implements rules C1, C4–C6 and C8–C10, renames images in order of appearance and writes fenced code blocks.
  - `tools/Invoke-LLPrototype.ps1` is a harness for running them on existing `.docx` files.
  - Validated on 502 reference pages (`docs/DESIGN.md` §12.2).

### Fixed

- **Table values could disappear.** Values in a monospace font inside multi-cell tables became code blocks, the table was then written as HTML, and the code blocks were dropped from it. Code in table cells is now written as inline code in a Markdown table, and empty table rows are removed.
- Stage 2 temporary files now go under `<WorkPath>\tmp` instead of `%TEMP%`, so one antivirus exclusion for `WorkPath` covers them.

### Changed

- `CreateObsidianSettings` (default on): a new vault gets `.obsidian\app.json` with settings that match Looseleaf's output. Existing vault settings are never touched.
- `OrderPrefix` is now `Auto` | `Always` | `Never` (default `Auto`; `$true`/`$false` still work). `Auto` skips prefixes in folders whose names already sort into OneNote order, which avoids double numbering such as `01-1 - Basics`.
- Names shortened for `MaxPathLength` are cut at a word boundary, and an unclosed bracket is dropped.
- The stage 2 command is `Update-LooseleafVault` instead of `Build-LooseleafVault`: `Build` isn't an approved verb in Windows PowerShell 5.1.
- No development dependencies: tests target the Pester version that ships with Windows, and PSScriptAnalyzer is optional.
- Settings that aren't implemented yet were removed from the allowed values: `ComplexTables = Flatten` and `ClipSources = FrontMatter`.
