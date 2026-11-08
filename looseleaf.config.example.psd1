#
# Looseleaf configuration (PowerShell data file - no code is executed).
# Copy to looseleaf.config.psd1 and edit. Full reference: docs/CONFIGURATION.md
#
# DRAFT: setting names follow docs/DESIGN.md and may change before 1.0.
#
@{
    # --- Paths ---------------------------------------------------------------

    # Notebooks to export, by name as shown in OneNote.
    Notebooks           = @('My Notebook')

    # Stage 1 output: manifest.json, docx/, attachments/, logs/. Absolute path.
    WorkPath            = 'D:\Looseleaf\work'

    # Obsidian vault root. Each notebook becomes a folder inside it. Absolute path.
    VaultPath           = 'D:\Obsidian\MyVault'

    # --- Stage 1: export -----------------------------------------------------

    # Expand collapsed paragraphs before export (WRITES to your notebook).
    ExpandCollapsed     = $false

    # Collapse them again after export (only when ExpandCollapsed is on).
    RestoreCollapsed    = $true

    # Also publish a PDF of every page that contains ink.
    ExportInkPagesAsPdf = $false

    # --- Stage 2: layout and naming ------------------------------------------

    # NN- prefixes to keep OneNote order in Obsidian:
    # Auto = only in folders whose names don't already sort correctly | Always | Never
    OrderPrefix         = 'Auto'

    # Notebook = one shared media/ folder per notebook; Folder = media/ next to each note.
    MediaScope          = 'Notebook'

    # SectionPage = <Section>-<Page>-NN.ext ; Page = <Page>-NN.ext
    ImageNameStyle      = 'SectionPage'
    MaxImageLabelLength = 40
    MaxImageNameLength  = 80

    # Relative = ![](../../media/x.png) ; Wiki = ![[x.png]]
    LinkStyle           = 'Relative'

    MaxNameLength       = 120
    MaxPathLength       = 250

    # Keep = 'Kick-off Notes.md' ; Hyphen = 'Kick-off-Notes.md'
    NameSpaces          = 'Keep'

    # Title-only pages: FolderOnly (folder for subpages) | Keep (title-only note) | Skip
    EmptyPages          = 'FolderOnly'

    # Notes you edited after the last build: Skip | Backup | Overwrite
    OnModifiedFile      = 'Skip'

    # New vault (no .obsidian folder yet): create matching Obsidian settings
    CreateObsidianSettings = $true

    # --- Stage 2: content (rules C1-C12, docs/DESIGN.md 6.5) ------------------

    FrontMatter         = $true
    TitleHeading        = $true

    # pandoc output format and extensions.
    PandocFormat        = 'markdown-simple_tables-multiline_tables-grid_tables-smart+pipe_tables'

    # Monospace text -> inline code / code blocks.
    DetectCode          = $true
    MonospaceFonts      = @('Consolas', 'Cascadia Code', 'Cascadia Mono', 'Courier New',
                            'Lucida Console', 'Source Code Pro', 'Fira Code', 'JetBrains Mono')

    # Single-cell tables: Auto | CodeBlock | Callout | Unwrap | Table
    BoxTables           = 'Auto'
    BoxCodeThreshold    = 0.6

    # Tables pipe tables can't hold: Html (rendered by Obsidian) | Flatten
    ComplexTables       = 'Html'

    # Indented paragraphs: Unindent | Quote
    IndentedParagraphs  = 'Unindent'

    # '● item' / '• item' paragraphs -> real list items
    ConvertGlyphBullets = $true

    # Safe = remove only escapes that change nothing (never < > &) ; None
    CleanEscapes        = 'Safe'

    # Underline: Html (<u>) | Drop. Highlight always becomes ==text==.
    Underline           = 'Html'

    # OneNote 'From <url>' lines: Normalize | FrontMatter | Keep
    ClipSources         = 'Normalize'

    # Image size: Obsidian (![|800](...)) | Drop
    ImageWidth          = 'Obsidian'
    MaxImageWidth       = 1000

    # LF or CRLF
    NewLine             = 'LF'
}
