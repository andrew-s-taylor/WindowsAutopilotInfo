<#PSScriptInfo
.VERSION 3.9
.GUID b608a45b-6cd0-405e-bfb2-aa11450821b5
.AUTHOR Alexey Semibratov - Updated by Andrew Taylor
.COMPANYNAME
.COPYRIGHT Alexey Semibratov
.TAGS
.LICENSEURI https://github.com/andrew-s-taylor/WindowsAutopilotInfo/blob/main/LICENSE
.PROJECTURI
.ICONURI
.EXTERNALMODULEDEPENDENCIES
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
Version 3.9: Fixed error with multi module import
Version 3.8: Timestamp fix
Version 3.7: Code Signed
Version 3.6: Added None option for assigned user
Version 3.5: Function update
Version 3.4: Fix in function name
Version 3.3: Changed method to grab devices
Version 3.2: Second fix
Version 3.1: Fix
Version 3.0: Updated to work with SDK v2
Version 2.9: Remove-MgDevice ObjectID switched to ID to match updated module
Version 2.8: Fixed speechmarks issue
Version 2.7: Changed Autopilot delete method
Version 2.6: Fixed mg-device command
Version 2.5: Typo
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
• This script will install the required modules
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


Function Connect-ToGraph {
    <#
.SYNOPSIS
Authenticates to the Graph API via the Microsoft.Graph.Authentication module.
 
.DESCRIPTION
The Connect-ToGraph cmdlet is a wrapper cmdlet that helps authenticate to the Intune Graph API using the Microsoft.Graph.Authentication module. It leverages an Azure AD app ID and app secret for authentication or user-based auth.
 
.PARAMETER Tenant
Specifies the tenant (e.g. contoso.onmicrosoft.com) to which to authenticate.
 
.PARAMETER AppId
Specifies the Azure AD app ID (GUID) for the application that will be used to authenticate.
 
.PARAMETER AppSecret
Specifies the Azure AD app secret corresponding to the app ID that will be used to authenticate.

.PARAMETER Scopes
Specifies the user scopes for interactive authentication.
 
.EXAMPLE
Connect-ToGraph -TenantId $tenantID -AppId $app -AppSecret $secret
 
-#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret,
        [Parameter(Mandatory = $false)] [string]$scopes
    )

    Process {
        Import-Module Microsoft.Graph.Authentication
        $version = (get-module microsoft.graph.authentication | Select-Object -expandproperty Version).major

        if ($AppId -ne "") {
            $body = @{
                grant_type    = "client_credentials";
                client_id     = $AppId;
                client_secret = $AppSecret;
                scope         = "https://graph.microsoft.com/.default";
            }
     
            $response = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token -Body $body
            $accessToken = $response.access_token
     
            $accessToken
            if ($version -eq 2) {
                write-host "Version 2 module detected"
                $accesstokenfinal = ConvertTo-SecureString -String $accessToken -AsPlainText -Force
            }
            else {
                write-host "Version 1 Module Detected"
                Select-MgProfile -Name Beta
                $accesstokenfinal = $accessToken
            }
            $graph = Connect-MgGraph  -AccessToken $accesstokenfinal 
            Write-Host "Connected to Intune tenant $TenantId using app-based authentication (Azure AD authentication not supported)"
        }
        else {
            if ($version -eq 2) {
                write-host "Version 2 module detected"
            }
            else {
                write-host "Version 1 Module Detected"
                Select-MgProfile -Name Beta
            }
            $graph = Connect-MgGraph -scopes $scopes
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
        }
    }
}    

function getdevicesandusers() {
    $alldevices = getallpagination -url "https://graph.microsoft.com/beta/devicemanagement/manageddevices"
    $outputarray = @()
    foreach ($value in $alldevices) {
        $objectdetails = [pscustomobject]@{
            DeviceID = $value.id
            DeviceName = $value.deviceName
            OSVersion = $value.operatingSystem
            PrimaryUser = $value.userPrincipalName
            operatingSystem = $value.operatingSystem
            AADID = $value.azureActiveDirectoryDeviceId
            SerialNumber = $value.serialnumber

        }
    
    
        $outputarray += $objectdetails
    
    }
    
    return $outputarray
    }

    function getallpagination () {
        <#
    .SYNOPSIS
    This function is used to grab all items from Graph API that are paginated
    .DESCRIPTION
    The function connects to the Graph API Interface and gets all items from the API that are paginated
    .EXAMPLE
    getallpagination -url "https://graph.microsoft.com/v1.0/groups"
     Returns all items
    .NOTES
     NAME: getallpagination
    #>
    [cmdletbinding()]
        
    param
    (
        $url
    )
        $response = (Invoke-MgGraphRequest -uri $url -Method Get -OutputType PSObject)
        $alloutput = $response.value
        
        $alloutputNextLink = $response."@odata.nextLink"
        
        while ($null -ne $alloutputNextLink) {
            $alloutputResponse = (Invoke-MGGraphRequest -Uri $alloutputNextLink -Method Get -outputType PSObject)
            $alloutputNextLink = $alloutputResponse."@odata.nextLink"
            $alloutput += $alloutputResponse.value
        }
        
        return $alloutput
        }

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
            Install-Module microsoft.graph.authentication -Force -ErrorAction Ignore
        }
        #Import-Module microsoft.graph.authentication -Scope Global

            $module = Import-Module microsoft.graph.groups -PassThru -ErrorAction Ignore
            if (-not $module) {
                Write-Host "Installing module MS Graph Groups"
                Install-Module microsoft.graph.groups -Force -ErrorAction Ignore
            }
            Import-Module microsoft.graph.groups -Scope Global


        $module2 = Import-Module Microsoft.Graph.Identity.DirectoryManagement -PassThru -ErrorAction Ignore
        if (-not $module2) {
            Write-Host "Installing module MS Graph Identity Management"
            Install-Module Microsoft.Graph.Identity.DirectoryManagement -Force -ErrorAction Ignore
        }
        Import-Module microsoft.graph.Identity.DirectoryManagement -Scope Global

        $module3 = Import-Module WindowsAutopilotIntuneCommunity -PassThru -ErrorAction Ignore
        if (-not $module3) {
            Write-Host "Installing module WindowsAutopilotIntuneCommunity"
            Install-Module WindowsAutopilotIntuneCommunity -Force -ErrorAction Ignore
        }
        Import-Module WindowsAutopilotIntuneCommunity -Scope Global


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
    Connect-ToGraph -Tenant $TenantId -AppId $AppId -AppSecret $AppSecret
    Write-Host "Connected to Intune tenant $TenantId using app-based authentication (Azure AD authentication not supported)"
}
else {
    $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All"
    Write-Host "Connected to Intune tenant $($graph.TenantId)"
    if ($AddToGroup) {
        $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All"
        Write-Host "Connected to Azure AD tenant $($aadId.TenantId)"
    }
}

Write-Host "Loading all objects. This can take a while on large tenants"
 $aadDevices = getallpagination -url "https://graph.microsoft.com/beta/devices"

$devices = getdevicesandusers

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
    
        $id = $currentAutopilotDevice.id
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/windowsAutopilotDeviceIdentities"    
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"
        Invoke-MGGraphRequest -Uri $uri -Method DELETE
        #Remove-AutopilotDevice -id $currentAutopilotDevice.id -ErrorAction Continue
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
        $escapedguid = "\" + ((([GUID]$aadDevice.deviceID).ToByteArray() |ForEach-Object {"{0:x}" -f $_}) -join '\')
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
        
        Remove-mgdevice -DeviceId $aadDevice.Id -ErrorAction SilentlyContinue
    }
    
}


# Get the hash (if available)
$devDetail = (Get-CimInstance -CimSession $session -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'")
if ($devDetail)
{
    $hash = $devDetail.DeviceHardwareData
    if($Host.UI.PromptForChoice('Add Autopilot Device', 'Do you want to *ADD* the device with serial number ' + $serial +' to Autopilot?', @('&Yes'; '&No'), 1) -eq 0){
        
        $newuserPrincipalName = Read-Host -Prompt "Change assigned user [$userPrincipalName] (type a new value or hit enter to keep the old one.  Enter None to not set a user)"
        if (![string]::IsNullOrWhiteSpace($newuserPrincipalName)){ $userPrincipalName = $newuserPrincipalName }

        $newgroupTag = Read-Host -Prompt "Change group tag [$groupTag] (type a new value or hit enter to keep the old one)"
        if (![string]::IsNullOrWhiteSpace($newgroupTag)){ $groupTag = $newgroupTag }

        ##If "None has been selected, don't add assigneduser"
        
        if ($userPrincipalName -eq "None") {
        Add-AutopilotImportedDevice -serialNumber $serial -hardwareIdentifier $hash -groupTag $groupTag
        }
        else {
        Add-AutopilotImportedDevice -serialNumber $serial -hardwareIdentifier $hash -groupTag $groupTag -assignedUser $userPrincipalName        
        }

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
# SIG # Begin signature block
# MIIoGQYJKoZIhvcNAQcCoIIoCjCCKAYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAUyceYFJ6/w0h2
# Ca0kDbGECibXcDmfVDNF8NP+XIzM06CCIRwwggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggauMIIElqADAgECAhAHNje3JFR82Ees/ShmKl5bMA0GCSqG
# SIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yMjAzMjMwMDAwMDBaFw0zNzAzMjIyMzU5NTlaMGMx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMy
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcg
# Q0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDGhjUGSbPBPXJJUVXH
# JQPE8pE3qZdRodbSg9GeTKJtoLDMg/la9hGhRBVCX6SI82j6ffOciQt/nR+eDzMf
# UBMLJnOWbfhXqAJ9/UO0hNoR8XOxs+4rgISKIhjf69o9xBd/qxkrPkLcZ47qUT3w
# 1lbU5ygt69OxtXXnHwZljZQp09nsad/ZkIdGAHvbREGJ3HxqV3rwN3mfXazL6IRk
# tFLydkf3YYMZ3V+0VAshaG43IbtArF+y3kp9zvU5EmfvDqVjbOSmxR3NNg1c1eYb
# qMFkdECnwHLFuk4fsbVYTXn+149zk6wsOeKlSNbwsDETqVcplicu9Yemj052FVUm
# cJgmf6AaRyBD40NjgHt1biclkJg6OBGz9vae5jtb7IHeIhTZgirHkr+g3uM+onP6
# 5x9abJTyUpURK1h0QCirc0PO30qhHGs4xSnzyqqWc0Jon7ZGs506o9UD4L/wojzK
# QtwYSH8UNM/STKvvmz3+DrhkKvp1KCRB7UK/BZxmSVJQ9FHzNklNiyDSLFc1eSuo
# 80VgvCONWPfcYd6T/jnA+bIwpUzX6ZhKWD7TA4j+s4/TXkt2ElGTyYwMO1uKIqjB
# Jgj5FBASA31fI7tk42PgpuE+9sJ0sj8eCXbsq11GdeJgo1gJASgADoRU7s7pXche
# MBK9Rp6103a50g5rmQzSM7TNsQIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB
# /wIBADAdBgNVHQ4EFgQUuhbZbU2FL3MpdpovdYxqII+eyG8wHwYDVR0jBBgwFoAU
# 7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoG
# CCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDig
# NqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9v
# dEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggIBAH1ZjsCTtm+YqUQiAX5m1tghQuGwGC4QTRPPMFPOvxj7x1Bd
# 4ksp+3CKDaopafxpwc8dB+k+YMjYC+VcW9dth/qEICU0MWfNthKWb8RQTGIdDAiC
# qBa9qVbPFXONASIlzpVpP0d3+3J0FNf/q0+KLHqrhc1DX+1gtqpPkWaeLJ7giqzl
# /Yy8ZCaHbJK9nXzQcAp876i8dU+6WvepELJd6f8oVInw1YpxdmXazPByoyP6wCeC
# RK6ZJxurJB4mwbfeKuv2nrF5mYGjVoarCkXJ38SNoOeY+/umnXKvxMfBwWpx2cYT
# gAnEtp/Nh4cku0+jSbl3ZpHxcpzpSwJSpzd+k1OsOx0ISQ+UzTl63f8lY5knLD0/
# a6fxZsNBzU+2QJshIUDQtxMkzdwdeDrknq3lNHGS1yZr5Dhzq6YBT70/O3itTK37
# xJV77QpfMzmHQXh6OOmc4d0j/R0o08f56PGYX/sr2H7yRp11LB4nLCbbbxV7HhmL
# NriT1ObyF5lZynDwN7+YAN8gFk8n+2BnFqFmut1VwDophrCYoCvtlUG3OtUVmDG0
# YgkPCr2B2RP+v6TR81fZvAT6gt4y3wSJ8ADNXcL50CN/AAvkdgIm2fBldkKmKYcJ
# RyvmfxqkhQ/8mJb2VVQrH4D6wPIOK+XW+6kvRBVK5xMOHds3OBqhK/bt1nz8MIIG
# sDCCBJigAwIBAgIQCK1AsmDSnEyfXs2pvZOu2TANBgkqhkiG9w0BAQwFADBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQw
# HhcNMjEwNDI5MDAwMDAwWhcNMzYwNDI4MjM1OTU5WjBpMQswCQYDVQQGEwJVUzEX
# MBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0
# ZWQgRzQgQ29kZSBTaWduaW5nIFJTQTQwOTYgU0hBMzg0IDIwMjEgQ0ExMIICIjAN
# BgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA1bQvQtAorXi3XdU5WRuxiEL1M4zr
# PYGXcMW7xIUmMJ+kjmjYXPXrNCQH4UtP03hD9BfXHtr50tVnGlJPDqFX/IiZwZHM
# gQM+TXAkZLON4gh9NH1MgFcSa0OamfLFOx/y78tHWhOmTLMBICXzENOLsvsI8Irg
# nQnAZaf6mIBJNYc9URnokCF4RS6hnyzhGMIazMXuk0lwQjKP+8bqHPNlaJGiTUyC
# EUhSaN4QvRRXXegYE2XFf7JPhSxIpFaENdb5LpyqABXRN/4aBpTCfMjqGzLmysL0
# p6MDDnSlrzm2q2AS4+jWufcx4dyt5Big2MEjR0ezoQ9uo6ttmAaDG7dqZy3SvUQa
# khCBj7A7CdfHmzJawv9qYFSLScGT7eG0XOBv6yb5jNWy+TgQ5urOkfW+0/tvk2E0
# XLyTRSiDNipmKF+wc86LJiUGsoPUXPYVGUztYuBeM/Lo6OwKp7ADK5GyNnm+960I
# HnWmZcy740hQ83eRGv7bUKJGyGFYmPV8AhY8gyitOYbs1LcNU9D4R+Z1MI3sMJN2
# FKZbS110YU0/EpF23r9Yy3IQKUHw1cVtJnZoEUETWJrcJisB9IlNWdt4z4FKPkBH
# X8mBUHOFECMhWWCKZFTBzCEa6DgZfGYczXg4RTCZT/9jT0y7qg0IU0F8WD1Hs/q2
# 7IwyCQLMbDwMVhECAwEAAaOCAVkwggFVMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYD
# VR0OBBYEFGg34Ou2O/hfEYb7/mF7CIhl9E5CMB8GA1UdIwQYMBaAFOzX44LScV1k
# TN8uZz/nupiuHA9PMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcD
# AzB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2lj
# ZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29t
# L0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmww
# HAYDVR0gBBUwEzAHBgVngQwBAzAIBgZngQwBBAEwDQYJKoZIhvcNAQEMBQADggIB
# ADojRD2NCHbuj7w6mdNW4AIapfhINPMstuZ0ZveUcrEAyq9sMCcTEp6QRJ9L/Z6j
# fCbVN7w6XUhtldU/SfQnuxaBRVD9nL22heB2fjdxyyL3WqqQz/WTauPrINHVUHmI
# moqKwba9oUgYftzYgBoRGRjNYZmBVvbJ43bnxOQbX0P4PpT/djk9ntSZz0rdKOtf
# JqGVWEjVGv7XJz/9kNF2ht0csGBc8w2o7uCJob054ThO2m67Np375SFTWsPK6Wrx
# oj7bQ7gzyE84FJKZ9d3OVG3ZXQIUH0AzfAPilbLCIXVzUstG2MQ0HKKlS43Nb3Y3
# LIU/Gs4m6Ri+kAewQ3+ViCCCcPDMyu/9KTVcH4k4Vfc3iosJocsL6TEa/y4ZXDlx
# 4b6cpwoG1iZnt5LmTl/eeqxJzy6kdJKt2zyknIYf48FWGysj/4+16oh7cGvmoLr9
# Oj9FpsToFpFSi0HASIRLlk2rREDjjfAVKM7t8RhWByovEMQMCGQ8M4+uKIw8y4+I
# Cw2/O/TOHnuO77Xry7fwdxPm5yg/rBKupS8ibEH5glwVZsxsDsrFhsP2JjMMB0ug
# 0wcCampAMEhLNKhRILutG4UI4lkNbcoFUCvqShyepf2gpx8GdOfy1lKQ/a+FSCH5
# Vzu0nAPthkX0tGFuv2jiJmCG6sivqf6UHedjGzqGVnhOMIIGwjCCBKqgAwIBAgIQ
# BUSv85SdCDmmv9s/X+VhFjANBgkqhkiG9w0BAQsFADBjMQswCQYDVQQGEwJVUzEX
# MBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0
# ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBMB4XDTIzMDcxNDAw
# MDAwMFoXDTM0MTAxMzIzNTk1OVowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRp
# Z2lDZXJ0LCBJbmMuMSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMzCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKNTRYcdg45brD5UsyPgz5/X
# 5dLnXaEOCdwvSKOXejsqnGfcYhVYwamTEafNqrJq3RApih5iY2nTWJw1cb86l+uU
# UI8cIOrHmjsvlmbjaedp/lvD1isgHMGXlLSlUIHyz8sHpjBoyoNC2vx/CSSUpIIa
# 2mq62DvKXd4ZGIX7ReoNYWyd/nFexAaaPPDFLnkPG2ZS48jWPl/aQ9OE9dDH9kgt
# XkV1lnX+3RChG4PBuOZSlbVH13gpOWvgeFmX40QrStWVzu8IF+qCZE3/I+PKhu60
# pCFkcOvV5aDaY7Mu6QXuqvYk9R28mxyyt1/f8O52fTGZZUdVnUokL6wrl76f5P17
# cz4y7lI0+9S769SgLDSb495uZBkHNwGRDxy1Uc2qTGaDiGhiu7xBG3gZbeTZD+BY
# QfvYsSzhUa+0rRUGFOpiCBPTaR58ZE2dD9/O0V6MqqtQFcmzyrzXxDtoRKOlO0L9
# c33u3Qr/eTQQfqZcClhMAD6FaXXHg2TWdc2PEnZWpST618RrIbroHzSYLzrqawGw
# 9/sqhux7UjipmAmhcbJsca8+uG+W1eEQE/5hRwqM/vC2x9XH3mwk8L9CgsqgcT2c
# kpMEtGlwJw1Pt7U20clfCKRwo+wK8REuZODLIivK8SgTIUlRfgZm0zu++uuRONhR
# B8qUt+JQofM604qDy0B7AgMBAAGjggGLMIIBhzAOBgNVHQ8BAf8EBAMCB4AwDAYD
# VR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZbU2FL3MpdpovdYxq
# II+eyG8wHQYDVR0OBBYEFKW27xPn783QZKHVVqllMaPe1eNJMFoGA1UdHwRTMFEw
# T6BNoEuGSWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRH
# NFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAGCCsGAQUFBwEBBIGD
# MIGAMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wWAYIKwYB
# BQUHMAKGTGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQwDQYJKoZIhvcNAQEL
# BQADggIBAIEa1t6gqbWYF7xwjU+KPGic2CX/yyzkzepdIpLsjCICqbjPgKjZ5+PF
# 7SaCinEvGN1Ott5s1+FgnCvt7T1IjrhrunxdvcJhN2hJd6PrkKoS1yeF844ektrC
# QDifXcigLiV4JZ0qBXqEKZi2V3mP2yZWK7Dzp703DNiYdk9WuVLCtp04qYHnbUFc
# jGnRuSvExnvPnPp44pMadqJpddNQ5EQSviANnqlE0PjlSXcIWiHFtM+YlRpUurm8
# wWkZus8W8oM3NG6wQSbd3lqXTzON1I13fXVFoaVYJmoDRd7ZULVQjK9WvUzF4UbF
# KNOt50MAcN7MmJ4ZiQPq1JE3701S88lgIcRWR+3aEUuMMsOI5ljitts++V+wQtaP
# 4xeR0arAVeOGv6wnLEHQmjNKqDbUuXKWfpd5OEhfysLcPTLfddY2Z1qJ+Panx+VP
# NTwAvb6cKmx5AdzaROY63jg7B145WPR8czFVoIARyxQMfq68/qTreWWqaNYiyjvr
# moI1VygWy2nyMpqy0tg6uLFGhmu6F/3Ed2wVbK6rr3M66ElGt9V/zLY4wNjsHPW2
# obhDLN9OTH0eaHDAdwrUAuBcYLso/zjlUlrWrBciI0707NMX+1Br/wd3H3GXREHJ
# uEbTbDJ8WC9nR2XlG3O2mflrLAZG70Ee8PBf4NvZrZCARK+AEEGKMIIHWzCCBUOg
# AwIBAgIQCLGfzbPa87AxVVgIAS8A6TANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQG
# EwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0
# IFRydXN0ZWQgRzQgQ29kZSBTaWduaW5nIFJTQTQwOTYgU0hBMzg0IDIwMjEgQ0Ex
# MB4XDTIzMTExNTAwMDAwMFoXDTI2MTExNzIzNTk1OVowYzELMAkGA1UEBhMCR0Ix
# FDASBgNVBAcTC1doaXRsZXkgQmF5MR4wHAYDVQQKExVBTkRSRVdTVEFZTE9SLkNP
# TSBMVEQxHjAcBgNVBAMTFUFORFJFV1NUQVlMT1IuQ09NIExURDCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBAMOkYkLpzNH4Y1gUXF799uF0CrwW/Lme676+
# C9aZOJYzpq3/DIa81oWv9b4b0WwLpJVu0fOkAmxI6ocu4uf613jDMW0GfV4dRodu
# tryfuDuit4rndvJA6DIs0YG5xNlKTkY8AIvBP3IwEzUD1f57J5GiAprHGeoc4Utt
# zEuGA3ySqlsGEg0gCehWJznUkh3yM8XbksC0LuBmnY/dZJ/8ktCwCd38gfZEO9UD
# DSkie4VTY3T7VFbTiaH0bw+AvfcQVy2CSwkwfnkfYagSFkKar+MYwu7gqVXxrh3V
# /Gjval6PdM0A7EcTqmzrCRtvkWIR6bpz+3AIH6Fr6yTuG3XiLIL6sK/iF/9d4U2P
# iH1vJ/xfdhGj0rQ3/NBRsUBC3l1w41L5q9UX1Oh1lT1OuJ6hV/uank6JY3jpm+Of
# Z7YCTF2Hkz5y6h9T7sY0LTi68Vmtxa/EgEtG6JVNVsqP7WwEkQRxu/30qtjyoX8n
# zSuF7TmsRgmZ1SB+ISclejuqTNdhcycDhi3/IISgVJNRS/F6Z+VQGf3fh6ObdQLV
# woT0JnJjbD8PzJ12OoKgViTQhndaZbkfpiVifJ1uzWJrTW5wErH+qvutHVt4/sEZ
# AVS4PNfOcJXR0s0/L5JHkjtM4aGl62fAHjHj9JsClusj47cT6jROIqQI4ejz1slO
# oclOetCNAgMBAAGjggIDMIIB/zAfBgNVHSMEGDAWgBRoN+Drtjv4XxGG+/5hewiI
# ZfROQjAdBgNVHQ4EFgQU0HdOFfPxa9Yeb5O5J9UEiJkrK98wPgYDVR0gBDcwNTAz
# BgZngQwBBAEwKTAnBggrBgEFBQcCARYbaHR0cDovL3d3dy5kaWdpY2VydC5jb20v
# Q1BTMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzCBtQYDVR0f
# BIGtMIGqMFOgUaBPhk1odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRU
# cnVzdGVkRzRDb2RlU2lnbmluZ1JTQTQwOTZTSEEzODQyMDIxQ0ExLmNybDBToFGg
# T4ZNaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29k
# ZVNpZ25pbmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5jcmwwgZQGCCsGAQUFBwEBBIGH
# MIGEMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXAYIKwYB
# BQUHMAKGUGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRHNENvZGVTaWduaW5nUlNBNDA5NlNIQTM4NDIwMjFDQTEuY3J0MAkGA1UdEwQC
# MAAwDQYJKoZIhvcNAQELBQADggIBAEkRh2PwMiyravr66Zww6Pjl24KzDcGYMSxU
# KOEU4bykcOKgvS6V2zeZIs0D/oqct3hBKTGESSQWSA/Jkr1EMC04qJHO/Twr/sBD
# CDBMtJ9XAtO75J+oqDccM+g8Po+jjhqYJzKvbisVUvdsPqFll55vSzRvHGAA6hjy
# DyakGLROcNaSFZGdgOK2AMhQ8EULrE8Riri3D1ROuqGmUWKqcO9aqPHBf5wUwia8
# g980sTXquO5g4TWkZqSvwt1BHMmu69MR6loRAK17HvFcSicK6Pm0zid1KS2z4ntG
# B4Cfcg88aFLog3ciP2tfMi2xTnqN1K+YmU894Pl1lCp1xFvT6prm10Bs6BViKXfD
# fVFxXTB0mHoDNqGi/B8+rxf2z7u5foXPCzBYT+Q3cxtopvZtk29MpTY88GHDVJsF
# MBjX7zM6aCNKsTKC2jb92F+jlkc8clCQQnl3U4jqwbj4ur1JBP5QxQprWhwde0+M
# ifDVp0vHZsVZ0pnYMCKSG5bUr3wOU7EP321DwvvEsTjCy/XDgvy8ipU6w3GjcQQF
# mgp/BX/0JCHX+04QJ0JkR9TTFZR1B+zh3CcK1ZEtTtvuZfjQ3viXwlwtNLy43vbe
# 1J5WNTs0HjJXsfdbhY5kE5RhyfaxFBr21KYx+b+evYyolIS0wR6New6FqLgcc4Ge
# 94yaYVTqMYIGUzCCBk8CAQEwfTBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGln
# aUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgQ29kZSBT
# aWduaW5nIFJTQTQwOTYgU0hBMzg0IDIwMjEgQ0ExAhAIsZ/Ns9rzsDFVWAgBLwDp
# MA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJ
# KoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQB
# gjcCARUwLwYJKoZIhvcNAQkEMSIEIHJEgrWKGri6p4T+J6BMGerLKitGp61sKwx+
# QsBqTR8gMA0GCSqGSIb3DQEBAQUABIICACaitdr082+yomX6FAn1utf3mJ8Lt86z
# JJnAK0h1eJHyK6QAleyv/GnPOf5eGyG4c/LbgWHtTynA0yFefZcZO5Yblo65CXhp
# HTx0yfr49xxWQOUiGnYenPjuvxaVqgi8Pk53TexwyjBV1tZ11u2EqWSB0Y6JuY0b
# 2YdMR8LIvZ/kNWao2piW4lJHsZTXdeUvJsOMrMBMTvatA12SYpAbqD8WHuczGlE5
# wI0+BnQ7Ip+KYDv4keSgexAlGmHaPeIf56WUmDWMzWd5ydPZ6+w1yQFN/8oSINIe
# dfn1bubfsQjuY83SxFnEDD6JUtjcqitF4otJyrxhxt/5o0eqZ4yDu4wT0bV2hFEK
# rXDLcMgyh6/oyZHvSC90O5x1Ctn/1HOrvjDnTLD1zMkZAGaCxVNusrd4pmbVEZ9P
# KoUChOgl9MtA91o9QwgOSGHc9bfM7TDJ7B4WHEvgjOPe4beDAVYkIac4GJMX0G3D
# +Ebd2toVywekBOwttn/Z7jpcoqInvq8tjO3xBs+8piyst42XnnR97xGhq/Fk9ROK
# mdwkN+654yMZaJ9zpTQwJ79WlCGLw6cH1jlehf8YrRCmU2U7nAq9g8LKojUb7ATX
# wjKhVsi4qGP7WE28FOtpPmxKvIm20GdxBOU37drONamEl8H0UWJUSRFmaXMN4E3b
# aSNS2roxgSZAoYIDIDCCAxwGCSqGSIb3DQEJBjGCAw0wggMJAgEBMHcwYzELMAkG
# A1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdp
# Q2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQQIQ
# BUSv85SdCDmmv9s/X+VhFjANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzEL
# BgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIzMTExNzA5MTAwOFowLwYJKoZI
# hvcNAQkEMSIEIAk3phpI9IJkmsB3N6e8/+VwHUQjuTnlxdGfF7EBimfsMA0GCSqG
# SIb3DQEBAQUABIICAGouWqqdcYqnfSB2OAeDmCRN5jNxXuQzFdb9P2L65LO0ISqX
# 5tVpoi70cnyX3F6OIAShbPNvdyQ4dlpVyMJ34fexbWRCsTUkbI/F+lRXV/iwOQdB
# kvFZnNxFVX4+sAlXasE0gK/C2yabcb/fW2LVYe9NRFlv+/rnpS6rijtpwdVbPsin
# iJygk7nDOkOIOVondbCdGGzLD2BCPn50lfw3WqmwzcEa0Z5BFgUiHl7rhOfysmmZ
# 0GaLdW/CKqY9LVMa3nFXE8I6vyGRxHFXjm8PcBSQWP3F7U1owQ6u03W9nU6cj4B+
# ZGbN5iYMxGZ51As54YpnhgcQdekCncsT2lfy+/mjufe490LX2KY7LEcufUZaLbbj
# dEktm2Y7FFQ8UYt8AIOfov+QG7VvkHhyX+L201tMP1wP+E9MFSKAc6ZLS9H40CQv
# B1Rzj797f/Cnd6e0hDU2GoORTLNLpV49fvmyY6bhfJR7qKZJe620AT/44fbroVCb
# xIzBwsbWfsez5ax69N8nTg81TBNwDF2DkTHZ3pdi0LAKzCHvBoGf3uusrIdemDu1
# NDVEwxLuUmOPPSUf1yZIbuwk4FyCyUepM9gvuw2A1zjP495ViqbDsA58rA26Zrkb
# l6Q2wtPR/z6uKMrXfVf1+HTxYHGPuKm8DlYCnvnu+AlT8DYcNUe/411Q3uIT
# SIG # End signature block
