@{
    Severity     = @('Error', 'Warning', 'Information')
    ExcludeRules = @()
    Rules        = @{
        PSUseCompatibleSyntax = @{
            Enable         = $true
            TargetVersions = @('5.1', '7.4')
        }
        PSAvoidUsingCmdletAliases = @{ Enable = $true }
        PSUseConsistentIndentation = @{
            Enable          = $true
            IndentationSize = 4
            Kind            = 'space'
        }
        PSUseConsistentWhitespace = @{ Enable = $true }
        PSPlaceOpenBrace = @{
            Enable     = $true
            OnSameLine = $true
        }
        PSUseBOMForUnicodeEncodedFile = @{ Enable = $true }
    }
}
