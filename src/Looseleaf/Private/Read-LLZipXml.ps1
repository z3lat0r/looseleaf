function Read-LLZipXml {
    <#
    .SYNOPSIS
        Reads an XML entry from an open ZipArchive, preserving whitespace.
    #>
    [CmdletBinding()]
    [OutputType([xml])]
    param (
        [Parameter(Mandatory)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory)]
        [string]$EntryName
    )

    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { throw "Entry '$EntryName' not found in archive." }

    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $stream = $entry.Open()
    try {
        $xml.Load($stream)
    } finally {
        $stream.Dispose()
    }
    , $xml
}
