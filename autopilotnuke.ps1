<#PSScriptInfo
 
.VERSION 2.4
 
.GUID b608a45b-6cd0-405e-bfb2-aa11450821b5
 
.AUTHOR Alexey Semibratov
 
.COMPANYNAME
 
.COPYRIGHT Alexey Semibratov
 
.TAGS
 
.LICENSEURI
 
.PROJECTURI
 
.ICONURI
 
.EXTERNALMODULEDEPENDENCIES
 
.REQUIREDSCRIPTS
 
.EXTERNALSCRIPTDEPENDENCIES
 
.RELEASENOTES
Version 2.4: Switched to MgGraph SDK and added support for app reg
Version 2.1: Bugfix
Version 2.0: Bugfix
Version 1.9: Bugfix
Version 1.8: Streamlined all logic with found Intune/AAD devices, changed output of found objects to a table
Version 1.7: Fixed a situation where there can be multiple Intune devices
Version 1.6: Added assigned user and tag - we will capture the old values, and will allow to change those if needed
Version 1.5: Some change in language around on-prem domain. Added wait for sync if it was less then 10 minutes ago. Fixed a bug when there is no AP devices, but we still want to delete Intune/AAD/AD devices.
Version 1.2: Added more documentation and set of required rights. Now if the device is not found in Autopilot, but exists in Intune (by serial number), it still cleans it from AD DS and AAD
Version 1.1: Invoke-AutopilotSync, when called too soon, error out
Version 1.0: Original public version.
 
#>

<#
 
.SYNOPSIS
Interactive script that helps to provision Autopilot machines. Identifies and fixes issues by removing the computer from Intune, AAD, AD and Autopilot, then adds it.
 
MIT LICENSE
 
Copyright (c) 2021 Alexey Semibratov
 
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 
 
.DESCRIPTION
Runs from OOBE screen, connects to Azure AD, Intune and optionally to AD DS, finds all objects for the serial number of the machine it is running on, then deletes it from everywhere, then adds it to Autopilot again.
Asks for deletion of each object
 Usage:
 - The script can work from running Windows 10, but be careful removing native Azure AD joined Intune Devices - you can lock yourself out, if you do not know local administrator's password
 - Intended usage - from OOBE (Out of Box Experience)
 - While in OOBE, hits Shift+F10
 - Powershell.exe
 - Install-Script AutopilotNuke
 - Accept all prompts
 - & 'C:\Program Files\WindowsPowerShell\Scripts\AutopilotNuke.ps1'
 - The script will:
        Download and install all required modules (accept all prompts)
        Show you the Serial Number of the machine
        Prompt to connect you to Azure AD and Intune Graph
        Ask you if you want to connect to local AD (ADDS, NT Domain) so it could delete old records from there. Enter the local FQDN (domain.com, contoso.local) of your AD Domain
        If you entered local AD domain, it will ask you for the username and password, for the username, use <NetbiosName>\User format
        Search in Autopilot for the serial number
        Show you all objects in Intune and AAD related to that Serial Number
        Ask if you want to delete in from Intune then deletes
        Ask if you want to delete in from Autopilot then deletes
        Loop through all AAD and AD (if it was selected) objects and ask to delete them
        Ask if you want to add it to AP then adds
 
Minimum security rights needed:
• To authorize Intune Graph, you will need global admin, but this is just one time. Ask your GA to run:
    Install-PackageProvider -Name NuGet
    Install-Module AzureAD
     Install-Module WindowsAutopilotIntune
    Install-Module Microsoft.Graph.Intune
    Connect-AzureAD
    Connect-MSGraph
    Accept the consent prompt
• Custom role with the following permissions required in Intune:
    Managed devices
        Read
        Delete
        Update
        Enrollment programs
        Create device
        Delete device
        Read device
        Sync device
    Assigned to All Devices (did not try scoping it with RBAC, but should work in theory)
• Cloud device administrator role required in Azure AD
• AD DS rights similar to Intune Connector rights: https://docs.microsoft.com/en-us/mem/autopilot/windows-autopilot-hybrid#:~:text=The%20Intune%20Connector%20for%20your,the%20rights%20to%20create%20computers.
 
 
#> 
[CmdletBinding()]
param(
    [Parameter(Mandatory = $False)] [String] $TenantId = "",
    [Parameter(Mandatory = $False)] [String] $AppId = "",
    [Parameter(Mandatory = $False)] [String] $AppSecret = ""
)

Write-Host "Downloading and installing all required modules, please accept all prompts"

        # Get NuGet
        $provider = Get-PackageProvider NuGet -ErrorAction Ignore
        if (-not $provider) {
            Write-Host "Installing provider NuGet"
            Find-PackageProvider -Name NuGet -ForceBootstrap -IncludeDependencies
        }
        
        # Get Graph Authentication module (and dependencies)
        $module = Import-Module microsoft.graph.authentication -PassThru -ErrorAction Ignore
        if (-not $module) {
            Write-Host "Installing module microsoft.graph.authentication"
            Install-Module microsoft.graph.authentication -Force
        }
        Import-Module microsoft.graph.authentication -Scope Global

            $module = Import-Module microsoft.graph.groups -PassThru -ErrorAction Ignore
            if (-not $module) {
                Write-Host "Installing module MS Graph Groups"
                Install-Module microsoft.graph.groups -Force
            }
            Import-Module microsoft.graph.groups -Scope Global


        $module2 = Import-Module Microsoft.Graph.Identity.DirectoryManagement -PassThru -ErrorAction Ignore
        if (-not $module2) {
            Write-Host "Installing module MS Graph Identity Management"
            Install-Module Microsoft.Graph.Identity.DirectoryManagement -Force
        }
        Import-Module microsoft.graph.Identity.DirectoryManagement -Scope Global

        $module3 = Import-Module WindowsAutopilotInfoCommunity -PassThru -ErrorAction Ignore
        if (-not $module3) {
            Write-Host "Installing module WindowsAutopilotInfoCommunity"
            Install-Module WindowsAutopilotInfoCommunity -Force
        }
        Import-Module WindowsAutopilotInfoCommunity -Scope Global


$session = New-CimSession
$DomainIP = $null
$de = $null
$autopilotDevices = $null
$aadDevices = $null
$intuneDevices = $null
$localADfqdn = $null
$DomainIP = $null
$de = $null
$relatedIntuneDevice=$null
$FoundAADDevices=$null

$groupTag=""
$userPrincipalName=""
$displayName=""
$newdisplayName=""

$serial = (Get-CimInstance -CimSession $session -Class Win32_BIOS).SerialNumber


Write-Host "Will be processing device with serial number: " -NoNewline
Write-Host $serial -ForegroundColor Green

Write-Host "Connecting to Intune Graph"

if ($AppId -ne "") {
    $body = @{
        grant_type    = "client_credentials";
        client_id     = $AppId;
        client_secret = $AppSecret;
        scope         = "https://graph.microsoft.com/.default";
    }

    $response = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
    $accessToken = $response.access_token

    $accessToken

    Select-MgProfile -Name Beta
    $graph = Connect-MgGraph  -AccessToken $accessToken 
    Write-Host "Connected to Intune tenant $TenantId using app-based authentication (Azure AD authentication not supported)"
}
else {
    $graph = Connect-MgGraph -scopes Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All
    Write-Host "Connected to Intune tenant $($graph.TenantId)"
    if ($AddToGroup) {
        $aadId = Connect-MgGraph -scopes Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All
        Write-Host "Connected to Azure AD tenant $($aadId.TenantId)"
    }
}

Write-Host "Loading all objects. This can take a while on large tenants"
$aadDevices = Get-MgDevice -All $true
##$intuneDevices = Get-IntuneManagedDevice -Filter "contains(operatingsystem, 'Windows')" | Get-MSGraphAllPages
$uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices"
$response = Invoke-MgGraphRequest -Uri $uri -Method Get -OutputType PSObject
         $devices = $response.value 
    
     $devicesNextLink = $response."@odata.nextLink"

     while ($null -ne $devicesNextLink) {
         $devicesResponse = (Invoke-MgGraphRequest -Uri $devicesNextLink -Method Get -OutputType PSObject)
         $devicesNextLink = $devicesResponse."@odata.nextLink"
         if ($serialhasspaces -eq 1) {
             $devices += $devicesResponse.value | Where-Object {$_.serialNumber -eq "$($serial)"}
         }
         else {
             $devices += $devicesResponse.value
         }
     }

     $intunedevices = $devices | Where-Object {$_.operatingSystem -eq "Windows"}

##$autopilotDevices = Get-AutopilotDevice | Get-MSGraphAllPages
$autopilotDevices = Get-AutopilotDevice


$localADfqdn = Read-Host -Prompt 'If you want to *DELETE* this computer from your local Active Directory domain and have Domain Controllers in line of sight, please enter the DNS of your AD DS domain (ie domain.local or contoso.com), otherwise, to skip AD DS deletion, hit "Enter"'
if($localADfqdn -ne "" -and $localADfqdn -ne $null)
{
    $DomainIP = (Test-Connection -ComputerName $localADfqdn -Count 1 -ErrorAction SilentlyContinue).IPV4Address.IPAddressToString
}


# Let's connect to on-prem AD

if($DomainIP -ne $null)
{

    Write-Host Connecting to $DomainIP
    Write-Host "Please provide the username and the password (DOMAIN\UserName)"
    $ADUserName = Read-Host -Prompt 'Username'
    $ADPassword = Read-Host -Prompt 'Password' -AsSecureString
    $ADPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ADPassword))
    $de = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainIP", $ADUserName, $ADPassword) -ErrorAction Stop
    Write-Host Connected to $de.distinguishedName   
}


$currentAutopilotDevice = $autopilotDevices | Where-Object {$_.serialNumber -eq $serial}

if ($currentAutopilotDevice -ne $null)
{

    # Find the objects linked to the Autopilot device

    Write-Verbose $currentAutopilotDevice |  Format-List -Property *
    
    [array]$relatedIntuneDevice = $intuneDevices | Where-Object {
    $_.serialNumber -eq $currentAutopilotDevice.serialNumber -or 
    $_.serialNumber -eq $currentAutopilotDevice.serialNumber.replace(' ','') -or 
    $_.id -eq $currentAutopilotDevice.managedDeviceId -or 
    $_.azureADDeviceId -eq $currentAutopilotDevice.azureActiveDirectoryDeviceId}       
   
    [array]$FoundAADDevices = $aadDevices | Where-Object { 
        $_.DeviceId -eq $currentAutopilotDevice.azureActiveDirectoryDeviceId -or 
        $_.DeviceId -iin $relatedIntuneDevice.azureADDeviceId -or 
        $_.DevicePhysicalIds -match $currentAutopilotDevice.Id
        }

    # Display a summary for this device and found related Intune /AAD devices

    Write-Host "User:" $currentAutopilotDevice.userPrincipalName
    Write-Host "Group Tag:" $currentAutopilotDevice.groupTag

    $userPrincipalName = $currentAutopilotDevice.userPrincipalName
    $groupTag = $currentAutopilotDevice.groupTag

    Write-Host "Found Related Intune Devices:"

    $relatedIntuneDevice | Format-Table -Property deviceName, id, userID, enrolledDateTime, LastSyncDateTime, operatingSystem, osVersion, deviceEnrollmentType

    Write-Host "Found Related AAD Devices:"

    $FoundAADDevices | Format-Table -Property DisplayName, ObjectID, DeviceID, AccountEnabled, ApproximateLastLogonTimeStamp, DeviceTrustType, DirSyncEnabled, LastDirSyncTime -AutoSize  


    if($relatedIntuneDevice -ne $null){
        foreach($relIntuneDevice in $relatedIntuneDevice)        {
            $displayName=$relIntuneDevice.deviceName
            if($Host.UI.PromptForChoice('Delete Intune Device', 'Do you want to *DELETE* ' + $relIntuneDevice.deviceName +' from the Intune?', @('&Yes'; '&No'), 1) -eq 0){
                $deviceid = $relIntuneDevice.id
                $url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$deviceid"
                $response = Invoke-MgGraphRequest -Uri $url -Method Delete -OutputType PSObject
            #Remove-IntuneManagedDevice -managedDeviceId $relIntuneDevice.id -ErrorAction Continue
            }
        }

    }


   
    if($Host.UI.PromptForChoice('Delete Autopilot Device', 'Do you want to *DELETE* the device with serial number ' + $currentAutopilotDevice.serialNumber +' from the Autopilot?', @('&Yes'; '&No'), 1) -eq 0){
    

        Remove-AutopilotDevice -id $currentAutopilotDevice.id -ErrorAction Continue
        $SecondsSinceLastSync = $null
        $SecondsSinceLastSync = (New-Timespan -Start (Get-AutopilotSyncInfo).lastSyncDateTime.ToUniversalTime()  -End (Get-Date).ToUniversalTime()).TotalSeconds
        If ($SecondsSinceLastSync -ge 610)
        {
            Invoke-AutopilotSync 
            
        }
        else
        {
            Write-Host "Last sync was" $SecondsSinceLastSync "seconds ago, will sleep for" (610-$SecondsSinceLastSync) "seconds before trying to sync."
            if($Host.UI.PromptForChoice('Autopilot Sync','Do you want to wait?', @('&Yes'; '&No'), 1) -eq 0){Start-Sleep -Seconds (610-$SecondsSinceLastSync) ; Invoke-AutopilotSync}            
        }
        while (Get-AutopilotDevice  | Where-Object {$_.serialNumber -eq $serial} -ne $null){
            Start-Sleep -Seconds 5                        
       }
       Write-Host "Deleted"

    }

}

if($relatedIntuneDevice -eq $null -and $FoundAADDevices -eq $null ){
    # this serial number was not found in Autopilot Devices, but we still want to check intune devices with this serial number and search AAD and AD DS for that one
    [array]$relatedIntuneDevice = $intuneDevices | Where-Object {$_.serialNumber -eq $serial -or $_.serialNumber -eq $serial.replace(' ','')}
    [array]$FoundAADDevices = $aadDevices | Where-Object { $_.DeviceId -eq $relatedIntuneDevice.azureADDeviceId }
    Write-Host "Found Related Intune Devices:"

    $relatedIntuneDevice | Format-Table -Property deviceName, id, userID, enrolledDateTime, LastSyncDateTime, operatingSystem, osVersion, deviceEnrollmentType

    Write-Host "Found Related AAD Devices:"

    $FoundAADDevices | Format-Table -Property DisplayName, ObjectID, DeviceID, AccountEnabled, ApproximateLastLogonTimeStamp, DeviceTrustType, DirSyncEnabled, LastDirSyncTime -AutoSize  


    if($relatedIntuneDevice -ne $null){
        foreach($relIntuneDevice in $relatedIntuneDevice)        {
            $displayName=$relIntuneDevice.deviceName
            if($Host.UI.PromptForChoice('Delete Intune Device', 'Do you want to *DELETE* ' + $relIntuneDevice.deviceName +' from the Intune?', @('&Yes'; '&No'), 1) -eq 0){
                $deviceid = $relIntuneDevice.id
                $url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$deviceid"
                $response = Invoke-MgGraphRequest -Uri $url -Method Delete -OutputType PSObject
            #Remove-IntuneManagedDevice -managedDeviceId $relIntuneDevice.id -ErrorAction Stop
            }
        }

    }

}



foreach($aadDevice in $FoundAADDevices){
    if($de -ne $null){            
        $escapedguid = “\” + ((([GUID]$aadDevice.deviceID).ToByteArray() |% {“{0:x}” -f $_}) -join '\')
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($de,"(&(objectCategory=Computer)(ObjectGUID=$escapedguid))")
        $obj = $searcher.FindOne()
        if ($obj -ne $null){
            $objdel = $obj.GetDirectoryEntry()
            if($Host.UI.PromptForChoice('Delete Active Directory Device', 'Do you want to *DELETE* the device with the name ' + $objdel.Name +' from AD DS?', @('&Yes'; '&No'), 1) -eq 0){
            $objdel.DeleteTree()
            }
                
        }
       
    }
    if($Host.UI.PromptForChoice('Delete Azure Active Directory Device', 'Do you want to *DELETE* the device with the name ' + $aadDevice.DisplayName +' from Azure AD?', @('&Yes'; '&No'), 1) -eq 0){
        
        Remove-mgdevice -DeviceId $aadDevice.ObjectID -ErrorAction SilentlyContinue
    }
    
}


# Get the hash (if available)
$devDetail = (Get-CimInstance -CimSession $session -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'")
if ($devDetail)
{
    $hash = $devDetail.DeviceHardwareData
    if($Host.UI.PromptForChoice('Add Autopilot Device', 'Do you want to *ADD* the device with serial number ' + $serial +' to Autopilot?', @('&Yes'; '&No'), 1) -eq 0){
        
        $newuserPrincipalName = Read-Host -Prompt "Change assigned user [$userPrincipalName] (type a new value or hit enter to keep the old one)"
        if (![string]::IsNullOrWhiteSpace($newuserPrincipalName)){ $userPrincipalName = $newuserPrincipalName }

        $newgroupTag = Read-Host -Prompt "Change group tag [$groupTag] (type a new value or hit enter to keep the old one)"
        if (![string]::IsNullOrWhiteSpace($newgroupTag)){ $groupTag = $newgroupTag }
        

        Add-AutopilotImportedDevice -serialNumber $serial -hardwareIdentifier $hash -groupTag $groupTag -assignedUser $userPrincipalName        

        $SecondsSinceLastSync = $null
        $SecondsSinceLastSync = (New-Timespan -Start (Get-AutopilotSyncInfo).lastSyncDateTime.ToUniversalTime()  -End (Get-Date).ToUniversalTime()).TotalSeconds
        If ($SecondsSinceLastSync -ge 610)
        {
            Invoke-AutopilotSync            
        }
        else
        {
            Write-Host "Last sync was" $SecondsSinceLastSync "seconds ago, will sleep for" (610-$SecondsSinceLastSync) "seconds before trying to sync."
            if($Host.UI.PromptForChoice('Autopilot Sync','Do you want to wait?', @('&Yes'; '&No'), 0) -eq 0){Start-Sleep -Seconds (610-$SecondsSinceLastSync); Invoke-AutopilotSync}
            
        }
        
    }

}

if($Host.UI.PromptForChoice('Computer name','Do you want to configure a unique name for a device? This name will be ignored in Hybrid Azure AD joined deployments. Device name still comes from the domain join profile for Hybrid Azure AD devices. This will only work if you have not deleted the device from AP recently.', @('&Yes'; '&No'), 1) -eq 0){

    $newdisplayName = Read-Host -Prompt "[$displayName] (type a new value or hit enter to keep the old one)"
    if (![string]::IsNullOrWhiteSpace($displayName) -or ![string]::IsNullOrWhiteSpace($newdisplayName)){ 
    
        if (![string]::IsNullOrWhiteSpace($newdisplayName) ){ $displayName = $newdisplayName }
           
        $autopilotDevices = Get-AutopilotDevice

        [array]$currentAutopilotDevices = $autopilotDevices | Where-Object {$_.serialNumber -eq $serial}

        foreach($currentAutopilotDevice in $currentAutopilotDevices){
        
            Set-AutopilotDevice -id $currentAutopilotDevice.id -displayName $displayName 
        }
            
    }

}