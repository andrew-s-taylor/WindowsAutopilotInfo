##Install Module
Write-Host "Installing WindowsAutoPilotIntune Module"

write-host "Getting Module Path"
$modulepath = $env:PSModulePath.Split(';')[1]
write-host "Module Path is $modulepath"
$folderpath = Join-Path $modulepath "WindowsAutopilotIntune"
write-host "Creating Folder $folderpath"
if(!(Test-Path $folderpath))
{
    New-Item -Path $folderpath -ItemType Directory
}
write-host "Folder Created"

write-host "Downloading Module Files"
$psm1uri = "https://raw.githubusercontent.com/andrew-s-taylor/WindowsAutopilotInfo/main/WindowsAutoPilotIntune.psm1"

$psm1path = Join-Path $folderpath "WindowsAutoPilotIntune.psm1"

Invoke-WebRequest -Uri $psm1uri -OutFile $psm1path
write-host "Downloaded WindowsAutoPilotIntune.psm1"

$psd1uri = "https://raw.githubusercontent.com/andrew-s-taylor/WindowsAutopilotInfo/main/WindowsAutoPilotIntune.psd1"

$psd1path = Join-Path $folderpath "WindowsAutoPilotIntune.psd1"

Invoke-WebRequest -Uri $psd1uri -OutFile $psd1path

write-host "Downloaded WindowsAutoPilotIntune.psd1"

##Installing Required modules
        # Get NuGet
        $provider = Get-PackageProvider NuGet -ErrorAction Ignore
        if (-not $provider) {
            Write-Host "Installing provider NuGet"
            Find-PackageProvider -Name NuGet -ForceBootstrap -IncludeDependencies
        }
                # Get Microsoft Graph Groups if needed
        if ($AddToGroup) {
            $module = Import-Module microsoft.graph.groups -PassThru -ErrorAction Ignore
            if (-not $module) {
                Write-Host "Installing module MS Graph Groups"
                Install-Module microsoft.graph.groups -Force
            }
            Import-Module microsoft.graph.groups -Scope Global

        }
        # Get Graph Authentication module (and dependencies)
        $module = Import-Module microsoft.graph.authentication -PassThru -ErrorAction Ignore
        if (-not $module) {
            Write-Host "Installing module microsoft.graph.authentication"
            Install-Module microsoft.graph.authentication -Force
        }
        Import-Module microsoft.graph.authentication -Scope Global




write-host "Importing Module"
Import-Module WindowsAutoPilotIntune -Global
write-host "Module Imported"

##Install Script

write-host "Installing get-windowsautopilotinfo.ps1 Script"
$scriptpath = $modulepath.Replace("Modules","Scripts")

$scripturi = "https://raw.githubusercontent.com/andrew-s-taylor/WindowsAutopilotInfo/main/get-windowsautopilotinfo.ps1"

write-host "Script Path is $scriptpath"
$scriptfilepath = Join-Path $scriptpath "get-windowsautopilotinfo.ps1"

Invoke-WebRequest -Uri $scripturi -OutFile $scriptfilepath
write-host "Downloaded get-windowsautopilotinfo.ps1"
