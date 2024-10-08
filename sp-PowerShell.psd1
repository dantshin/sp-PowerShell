@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'sp-PowerShell.psm1'

    # Version number of this module.
    ModuleVersion = '1.0.0'

    # ID used to uniquely identify this module
    GUID = '086bcd07-d6c5-4813-9795-b3176e461f8b'

    # Author of this module
    Author = 'UNDP'

    # Company or vendor of this module
    CompanyName = 'UNDP'

    # Copyright statement for this module
    Copyright = '(c) UNDP. All rights reserved.'

    # Description of the functionality provided by this module
    Description = 'PowerShell cmdlets for managing SPO'

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '5.0'

    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @(
        'PnP.PowerShell'
    )

    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @()

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    # ScriptsToProcess = @()

    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    NestedModules = @(
        'List/ListItem.ps1',
        'Site/Site.ps1'
    )

    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport = @(
        'Get-SiteInfo',
        'Set-SiteLogo',
        'Remove-ListItemVersionHistory'
    )

    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    # CmdletsToExport = @()

    # Variables to export from this module
    # VariablesToExport = '*'

    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    # AliasesToExport = @()

    # List of all modules packaged with this module
    # ModuleList = @()

    # List of all files packaged with this module
    # FileList = @()

    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData = @{
        PSData = @{
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags = @('PowerShell', 'SharePoint', 'SPO')

            # A URL to the license for this module.
            LicenseUri = 'https://github.com/undp/sp-PowerShell/blob/master/LICENSE.md'

            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/undp/sp-PowerShell'

            # A URL to an icon representing this module.
            # IconUri = ''

            # ReleaseNotes of this module
            # ReleaseNotes = ''
        } # End of PSData hashtable
    } # End of PrivateData hashtable

    # HelpInfo URI of this module
    # HelpInfoURI = ''
}
