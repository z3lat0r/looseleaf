@{
    RootModule           = 'Looseleaf.psm1'
    ModuleVersion        = '0.0.1'
    GUID                 = '537e05b8-50b8-4c94-b481-dfe2707ee907'
    Author               = 'z3lat0r'
    Copyright            = '(c) 2026 z3lat0r. Licensed under GPL-3.0.'
    Description          = 'Export OneNote notebooks to Obsidian-ready Markdown, locally (OneNote desktop COM + pandoc).'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')

    FunctionsToExport    = @(
        'Test-LooseleafEnvironment'
        'Export-LooseleafNotebook'
        'Build-LooseleafVault'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()

    PrivateData          = @{
        PSData = @{
            Tags       = @('OneNote', 'Obsidian', 'Markdown', 'pandoc', 'Export', 'Windows')
            LicenseUri = 'https://www.gnu.org/licenses/gpl-3.0.html'
        }
    }
}
