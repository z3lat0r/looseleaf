# Changelog

All notable changes to Looseleaf are recorded here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- Project scaffold: module layout, configuration template, lint settings, and repository conventions (`.editorconfig`, `.gitattributes`, `.gitignore`).
- Documentation: README, design document, configuration reference, and troubleshooting guide.
- Design: content rules C1–C12 (box tables, monospace detection, indentation, glyph bullets, safe escaping, underline and highlight, clip sources, image sizes), handling of empty pages, protection of edited files, and retries for locked files. Derived from analysing a 500-page reference notebook.
- Stage 1 test-notebook checklist (`docs/TEST-NOTEBOOK.md`).
- Conversion core prototype:
  - `ConvertTo-LLCodeMarkedDocx` marks monospace text with Word code styles (rule C2).
  - `filters/looseleaf.lua` implements rules C1, C4–C6 and C8–C10, renames images in order of appearance and writes fenced code blocks.
  - `tools/Invoke-LLPrototype.ps1` is a harness for running them on existing `.docx` files.
  - Validated on 502 reference pages (`docs/DESIGN.md` §12.2).
