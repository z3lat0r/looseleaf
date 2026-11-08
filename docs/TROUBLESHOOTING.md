# Troubleshooting

Run `Test-LooseleafEnvironment -ConfigPath .\looseleaf.config.psd1` first. It detects most of the problems below without changing anything.

## Do I need to run as Administrator?

**No.** OneNote is controlled through COM, and Windows only connects two processes running at the **same privilege level**. Upstream ConvertOneNote2MarkDown recommends running *both* PowerShell and OneNote as Administrator, but that's just one way to make them match.

**Recommended:** start OneNote normally, and run PowerShell normally (not "Run as administrator").

Symptoms of a mismatch (for example, an elevated PowerShell and a normal OneNote):

- `Exception calling "GetHierarchy"…` or notebooks missing from the hierarchy
- `Exception calling "Publish"…: "Class not registered"`
- `Unable to find type [Microsoft.Office.InterOp.OneNote.HierarchyScope]` (upstream only)

**Fix:** close both, then start both normally. Looseleaf's preflight check reports this mismatch directly.

## "Class not registered" even though the privilege levels match

- **Wrong OneNote:** *OneNote for Windows 10* (the Microsoft Store app) has no COM API. Install **OneNote desktop** (OneNote 2016 or the OneNote in Microsoft 365).
- **Broken Office registration:** run an Office *Quick Repair* (Settings → Apps → Microsoft 365 → Modify).

## Content missing from collapsed paragraphs

OneNote's Word export leaves out paragraphs that are collapsed. You have two options:

- Set `ExpandCollapsed = $true` (and `RestoreCollapsed = $true`). Looseleaf expands them before export and collapses them afterwards. This changes each affected page's modified time.
- Or expand them yourself. The [Onetastic](https://getonetastic.com) macro *Expand all Paragraphs in Notebook* shipped with ConvertOneNote2MarkDown does the same thing.

The preflight check counts affected pages, so you can see the scale before you decide.

## Locked (password-protected) sections are skipped

The COM API can't read locked sections. Unlock them in OneNote (click the section and enter the password), then run the export again before OneNote locks them again (see OneNote *Options → Passwords*).

## Handwriting and drawings are missing

Ink doesn't survive the conversion to Markdown. Options:

- `ExportInkPagesAsPdf = $true` to keep a PDF of each affected page next to the export
- Or, in OneNote, use *Draw → Ink to Text* or *Ink to Shape*, or take a screenshot of the drawing and paste it in, before exporting.

## "File is being used by another process" / files disappear (antivirus)

Notes about security topics (exploit code, reverse shells, payloads, password-cracking commands) can trigger antivirus software. Microsoft Defender may lock a `.docx` or `.md` file while it scans it, or quarantine it. Looseleaf retries locked files a few times, then marks the page as failed with a hint.

If it keeps happening:

- Check *Windows Security → Virus & threat protection → Protection history* for quarantined files.
- If you trust the content, add folder exclusions for `WorkPath` and your vault (*Manage settings → Exclusions*). Only do this if your organisation's policy allows it. Looseleaf keeps all its temporary files under `WorkPath\tmp`, so no other folder is involved.
- Excluding the vault matters too: Obsidian, sync tools and backups read those notes later, and an antivirus scan can still quarantine a note after Looseleaf has written it.

## pandoc not found

Install pandoc from <https://pandoc.org/installing.html>, then **open a new PowerShell window** so the updated `PATH` is picked up. Check it with `pandoc --version` (3.x is required).

## Path too long

Windows' traditional limit is 260 characters for a full path. Looseleaf shortens names to stay under `MaxPathLength` (250 by default) and warns when it does. To allow longer paths, enable Windows long paths (Group Policy *Enable Win32 long paths*, or the registry value `LongPathsEnabled`) and raise `MaxPathLength`. Some tools, including older versions of Obsidian's sync clients, may still have trouble with very long paths.

## Export was interrupted

Run `Export-LooseleafNotebook` again. Pages already exported and unchanged since are skipped. Use `-Force` to re-export everything.

## PowerShell 7 problems

Windows PowerShell 5.1 (`powershell.exe`) is the supported baseline. If something fails only in PowerShell 7 (`pwsh.exe`), run the same command in `powershell.exe` and please report the difference.
