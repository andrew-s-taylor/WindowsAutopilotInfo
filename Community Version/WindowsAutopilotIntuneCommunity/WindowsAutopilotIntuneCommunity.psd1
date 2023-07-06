#
# Module manifest for module 'WindowsAutoPilotIntuneCommunity'
#
# Generated by: mniehaus - amended by Andrew Taylor
#
# Generated on: 4/13/2018 - Updated on 14/06/2023
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'WindowsAutoPilotIntuneCommunity'

# Version number of this module.
ModuleVersion = '2.1'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = '0adc2f8f-18eb-43fe-ab34-328c5877a528'

# Author of this module
Author = 'Andrew Taylor & Michael Niehaus'

# Company or vendor of this module
CompanyName = 'Intune Community'

# Copyright statement for this module
Copyright = 'GPL'

# Description of the functionality provided by this module
Description = 'Sample module to manage AutoPilot devices using the Intune Graph API'

# Minimum version of the Windows PowerShell engine required by this module
# PowerShellVersion = ''

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @('Microsoft.Graph.Groups','Microsoft.Graph.Authentication')

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = '*'

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = @()

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        # Tags = @()

        # A URL to the license for this module.
        # LicenseUri = 'https://github.com/andrew-s-taylor/WindowsAutopilotInfo/blob/main/LICENSE'

        # A URL to the main website for this project.
        # ProjectUri = 'https://github.com/andrew-s-taylor/WindowsAutopilotInfo/community'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        ReleaseNotes = @'
Version 1.0: First releast of community fork
Version 1.1: Fixed Get-Organization issue for domain
Version 2.0: Updated to work with v2 SDK
Version 2.1: Authentication fix
'@
    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}
