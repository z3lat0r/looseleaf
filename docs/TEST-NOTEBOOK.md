# Test notebook for stage 1

Stage 1 talks to OneNote, so it can't be covered by automated tests. Instead, build a small throwaway notebook once in OneNote desktop, and run every stage 1 change against it. Each item below maps to a rule or check in [DESIGN.md](DESIGN.md).

> [!IMPORTANT]
> Use a **separate notebook** (for example `Looseleaf Test`) stored locally. `ExpandCollapsed` writes to the notebook, so never test it on real notes.

## Structure

- [ ] Two **section groups**, one of them nested inside the other
- [ ] One section **outside** any group
- [ ] A section with more than 99 pages (copy one page repeatedly), to test three-digit order prefixes
- [ ] Subpages at level 2 and level 3 under one parent page
- [ ] An **empty parent page** (title only) with subpages (`EmptyPages`)
- [ ] An **empty page without subpages**
- [ ] A **divider page** titled `----- Linux -----`
- [ ] A page in the **recycle bin** (should be skipped)
- [ ] A **password-protected section**, left locked (should be skipped with a warning)

## Page titles (naming rules, DESIGN.md §6.2)

- [ ] `[ LAB 1.2 ] Brackets and spaces`
- [ ] `( Windows )`
- [ ] `Demo: Colon in title`
- [ ] `Slash / and backslash \ and pipe |`
- [ ] `Hash # caret ^ question ? star * quotes " <angle>`
- [ ] `CON` and `aux.txt` (reserved device names)
- [ ] `Trailing dot.` and `Trailing spaces   `
- [ ] `2023-10-05 Meeting` (leading date must survive)
- [ ] A 200-character title (length limits)
- [ ] Two pages titled `Introduction` in the same section (clash handling)
- [ ] Non-ASCII: `Café – naïve 日本語 😀`

## Content (rules C1–C12, DESIGN.md §6.5)

- [ ] **C1/C2:** a single-cell table containing commands in **Consolas** (should become a code block)
- [ ] **C1:** a single-cell table containing ordinary text in Calibri (should become a callout)
- [ ] **C2:** an inline command in Cascadia Code inside a normal sentence
- [ ] **C3:** a 3×3 table with line breaks in cells; a table with merged cells; a table with a bullet list in a cell
- [ ] **C4:** paragraphs indented 1, 2 and 3 levels; a quote with OneNote's *Quote* style
- [ ] **C5:** lines typed as `● item`, `• item`, `+ item`
- [ ] **C6:** several empty lines in a row, and lines containing only spaces
- [ ] **C7:** the text `ssh <user>@<host> -p <port>` written **outside** a code box
- [ ] **C8:** underlined text; highlighted text (several colours)
- [ ] **C9:** content pasted from a web page, so OneNote adds a `From <url>` line
- [ ] **C10:** a large screenshot, and one resized smaller in OneNote
- [ ] **C11:** a bulleted list with links pasted from a web page into a table cell
- [ ] Headings 1–3, bold and italic, numbered and nested lists, checkboxes (To Do tags), other tags (Important, Question)
- [ ] Hyperlinks: web URLs, and a link to another OneNote page

## Images, attachments and special objects

- [ ] A page with 12 or more images (numbering `-01`…`-12`, order of appearance)
- [ ] Two pages whose titles share the first 50 characters, both with images (no overwriting, ON2MD bug 1)
- [ ] An attached file (PDF) and a file printout
- [ ] **Ink:** a handwritten drawing and handwriting (`ExportInkPagesAsPdf`)
- [ ] **Collapsed paragraphs:** an outline with a collapsed level (`ExpandCollapsed`, `RestoreCollapsed`)
- [ ] An image with OneNote alt text
- [ ] An embedded external image (web URL) if OneNote keeps one

## Behaviour to verify on each run

- [ ] Preflight reports the locked section, collapsed paragraphs and ink pages
- [ ] A second run skips all unchanged pages
- [ ] After editing one page in OneNote, only that page is re-exported
- [ ] Interrupting the run (Ctrl+C) and starting it again resumes correctly
- [ ] `-WhatIf` changes nothing
- [ ] Running PowerShell as Administrator while OneNote runs normally gives the privilege-mismatch error
- [ ] After a build, edit one note in Obsidian and rebuild: the note is kept (`OnModifiedFile = Skip`)
