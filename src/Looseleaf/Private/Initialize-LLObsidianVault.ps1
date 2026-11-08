function Initialize-LLObsidianVault {
    <#
    .SYNOPSIS
        Makes VaultPath an Obsidian vault by creating .obsidian\app.json with settings that match
        Looseleaf's output. Does nothing if the folder already has a .obsidian folder, so the
        settings of an existing vault are never changed.

    .OUTPUTS
        $true if the settings were created, otherwise $false.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Config,

        [Parameter(Mandatory)]
        [System.Management.Automation.PSCmdlet]$Cmdlet
    )

    $obsidianDir = Join-Path $Config.VaultPath '.obsidian'
    if (Test-Path -LiteralPath $obsidianDir) { return $false }

    $settings = [ordered]@{
        # Relative = Markdown links ![](../media/x.png); Wiki = ![[x.png]]
        useMarkdownLinks  = ($Config.LinkStyle -eq 'Relative')
        newLinkFormat     = 'relative'
        alwaysUpdateLinks = $true
    }
    # Where images pasted in Obsidian go: next to the notes Looseleaf wrote.
    if ($Config.MediaScope -eq 'Folder') {
        $settings.attachmentFolderPath = './media'
    } elseif (@($Config.Notebooks).Count -eq 1) {
        $notebookFolder = Get-LLSafeName -Name @($Config.Notebooks)[0] -NameSpaces $Config.NameSpaces -MaxLength $Config.MaxNameLength
        $settings.attachmentFolderPath = "$notebookFolder/media"
    }

    if (-not $Cmdlet.ShouldProcess($obsidianDir, 'Create Obsidian vault settings')) { return $false }
    [void][IO.Directory]::CreateDirectory($obsidianDir)
    [IO.File]::WriteAllText((Join-Path $obsidianDir 'app.json'), ($settings | ConvertTo-Json), (New-Object System.Text.UTF8Encoding $false))
    $true
}
