function Write-LLZipXml {
    <#
    .SYNOPSIS
        Replaces an XML entry in a ZipArchive opened in Update mode (UTF-8, no BOM).
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory)]
        [string]$EntryName,

        [Parameter(Mandatory)]
        [xml]$Xml
    )

    $existing = $Zip.GetEntry($EntryName)
    if ($existing) { $existing.Delete() }

    $entry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding $false
        $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
        try {
            $Xml.Save($writer)
        } finally {
            $writer.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}
