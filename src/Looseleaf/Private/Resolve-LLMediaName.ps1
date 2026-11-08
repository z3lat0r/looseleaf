function Resolve-LLMediaName {
    <#
    .SYNOPSIS
        Media registry (docs/DESIGN.md 6.3): a stable, unique file name for one image or attachment.

    .DESCRIPTION
        RegistryKey identifies the item ('<pageKey>|img|<NN>' or '<pageKey>|att|<name>').
        - Known key from the last build: keeps its previous name.
        - New key: uses Desired, or Desired-2, -3 ... if another item owns that name.
        A name is never shared by two items, so files are never overwritten across pages.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string]$RegistryKey,

        [Parameter(Mandatory)]
        [string]$Desired,

        [Parameter(Mandatory)]
        [hashtable]$OldMedia,

        [Parameter(Mandatory)]
        [hashtable]$NewMedia,

        # name (lower case) -> registry key that owns it
        [Parameter(Mandatory)]
        [hashtable]$Owned
    )

    if ($NewMedia.ContainsKey($RegistryKey)) { return $NewMedia[$RegistryKey] }

    if ($OldMedia.ContainsKey($RegistryKey)) {
        $previous = $OldMedia[$RegistryKey]
        $owner = $Owned[$previous.ToLowerInvariant()]
        if (-not $owner -or $owner -eq $RegistryKey) {
            $Owned[$previous.ToLowerInvariant()] = $RegistryKey
            $NewMedia[$RegistryKey] = $previous
            return $previous
        }
    }

    $base = [IO.Path]::GetFileNameWithoutExtension($Desired)
    $ext = [IO.Path]::GetExtension($Desired)
    $candidate = $Desired
    $n = 2
    while ($Owned.ContainsKey($candidate.ToLowerInvariant()) -and $Owned[$candidate.ToLowerInvariant()] -ne $RegistryKey) {
        $candidate = "$base-$n$ext"
        $n++
    }
    $Owned[$candidate.ToLowerInvariant()] = $RegistryKey
    $NewMedia[$RegistryKey] = $candidate
    $candidate
}
