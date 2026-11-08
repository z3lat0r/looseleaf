# Changelog

All notable changes to Looseleaf are recorded here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- **Stage 2: `Update-LooseleafVault`.** Builds the vault from `<WorkPath>\manifest.json`:
  - mirrors the OneNote hierarchy (section groups, sections, subpages as folders) with `NN-` order prefixes, applies the naming rules and `EmptyPages`, and shortens names to fit `MaxPathLength`
  - media registry: stable, unique image and attachment names
  - front matter, title heading and an attachments list
  - leaves notes edited after the last build alone (`OnModifiedFile`) and removes stale output
  - retries files locked by antivirus; supports `-WhatIf` and `-StopOnError`; writes logs to `<WorkPath>\logs`
- Configuration loading with a schema: unknown keys, wrong types and relative paths are rejected.
- Pester tests (72) for the **built-in Pester 3.4**: unit tests and end-to-end build tests on made-up pages.
- Project scaffold: module layout, configuration template, lint settings, and repository conventions (`.editorconfig`, `.gitattributes`, `.gitignore`).
- Documentation: README, design document, configuration reference, and troubleshooting guide.
- Design: content rules C1–C12 (box tables, monospace detection, indentation, glyph bullets, safe escaping, underline and highlight, clip sources, image sizes), handling of empty pages, protection of edited files, and retries for locked files. Derived from analysing a 500-page reference notebook.
- Stage 1 test-notebook checklist (`docs/TEST-NOTEBOOK.md`).
- Conversion core prototype:
  - `ConvertTo-LLCodeMarkedDocx` marks monospace text with Word code styles (rule C2).
  - `filters/looseleaf.lua` implements rules C1, C4–C6 and C8–C10, renames images in order of appearance and writes fenced code blocks.
  - `tools/Invoke-LLPrototype.ps1` is a harness for running them on existing `.docx` files.
  - Validated on 502 reference pages (`docs/DESIGN.md` §12.2).

### Changed

- The stage 2 command is `Update-LooseleafVault` instead of `Build-LooseleafVault`: `Build` isn't an approved verb in Windows PowerShell 5.1.
- No development dependencies: tests target the Pester version that ships with Windows, and PSScriptAnalyzer is optional.
- Settings that aren't implemented yet were removed from the allowed values: `ComplexTables = Flatten` and `ClipSources = FrontMatter`.
