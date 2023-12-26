
#region Helper methods

Function BoolToString() {
    param
    (
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $True)] [bool] $value
    )

    Process {
        return $value.ToString().ToLower()
    }
}

#endregion

#region App-based authentication
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

#region Core methods

Function Get-AutopilotDevice() {
    <#
.SYNOPSIS
Gets devices currently registered with Windows Autopilot.
 
.DESCRIPTION
The Get-AutopilotDevice cmdlet retrieves either the full list of devices registered with Windows Autopilot for the current Azure AD tenant, or a specific device if the ID of the device is specified.
 
.PARAMETER id
Optionally specifies the ID (GUID) for a specific Windows Autopilot device (which is typically returned after importing a new device)
 
.PARAMETER serial
Optionally specifies the serial number of the specific Windows Autopilot device to retrieve
 
.PARAMETER expand
Expand the properties of the device to include the Autopilot profile information
 
.EXAMPLE
Get a list of all devices registered with Windows Autopilot
 
Get-AutopilotDevice
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $True)] $id,
        [Parameter(Mandatory = $false)] $serial,
        [Parameter(Mandatory = $false)] [Switch]$expand = $false,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/windowsAutopilotDeviceIdentities"
    
        if ($id -and $expand) {
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)/$($id)?`$expand=deploymentProfile,intendedDeploymentProfile"
        }
        elseif ($id) {
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)/$id"
        }
        elseif ($serial) {
            $encoded = [uri]::EscapeDataString($serial)
            ##Check if serial contains a space
            $serialelements = $serial.Split(" ")
            if ($serialelements.Count -gt 1) {
                $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)?`$filter=contains(serialNumber,'$($serialelements[0])')"
                $serialhasspaces = 1
            }
            else {
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)?`$filter=contains(serialNumber,'$encoded')"
            }
        }
        else {
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
        }

        Write-Verbose "GET $uri"

        try {
            $response = Invoke-MgGraphRequest -Uri $uri -Method Get -OutputType PSObject
            if ($id) {
                $response
            }
            else {
                if ($serialhasspaces -eq 1) {  
                    $devices = $response.value | Where-Object {$_.serialNumber -eq "$($serial)"}
               } else {
                    $devices = $response.value 
               }
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
    
                if ($expand) {
                    $devices | Get-AutopilotDevice -Expand
                }
                else {
                    $devices
                }
            }
        }
        catch {
            Write-Error $_.Exception 
            break
        }
    }
}


Function Set-AutopilotDevice() {
    <#
.SYNOPSIS
Updates settings on an Autopilot device.
 
.DESCRIPTION
The Set-AutopilotDevice cmdlet can be used to change the updatable properties on a Windows Autopilot device object.
 
.PARAMETER id
The Windows Autopilot device id (mandatory).
 
.PARAMETER userPrincipalName
The user principal name.
 
.PARAMETER addressibleUserName
The name to display during Windows Autopilot enrollment. If specified, the userPrincipalName must also be specified.
 
.PARAMETER displayName
The name (computer name) to be assigned to the device when it is deployed via Windows Autopilot. This is presently only supported with Azure AD Join scenarios. Note that names should not exceed 15 characters. After setting the name, you need to initiate a sync (Invoke-AutopilotSync) in order to see the name in the Intune object.
 
.PARAMETER groupTag
The group tag value to set for the device.
 
.EXAMPLE
Assign a user and a name to display during enrollment to a Windows Autopilot device.
 
Set-AutopilotDevice -id $id -userPrincipalName $userPrincipalName -addressableUserName "John Doe" -displayName "CONTOSO-0001" -groupTag "Testing"
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] $id,
        [Parameter(ParameterSetName = "Prop")] $userPrincipalName = $null,
        [Parameter(ParameterSetName = "Prop")] $addressableUserName = $null,
        [Parameter(ParameterSetName = "Prop")][Alias("ComputerName", "CN", "MachineName")] $displayName = $null,
        [Parameter(ParameterSetName = "Prop")] $groupTag = $null,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/windowsAutopilotDeviceIdentities"
    
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id/UpdateDeviceProperties"

        $json = "{"
        if ($PSBoundParameters.ContainsKey('userPrincipalName')) {
            $json = $json + " userPrincipalName: `"$userPrincipalName`","
        }
        if ($PSBoundParameters.ContainsKey('addressableUserName')) {
            $json = $json + " addressableUserName: `"$addressableUserName`","
        }
        if ($PSBoundParameters.ContainsKey('displayName')) {
            $json = $json + " displayName: `"$displayName`","
        }
        if ($PSBoundParameters.ContainsKey('groupTag')) {
            $json = $json + " groupTag: `"$groupTag`""
        }
        else {
            $json = $json.Trim(",")
        }
        $json = $json + " }"

        Write-Verbose "POST $uri`n$json"

        try {
            Invoke-MGGraphRequest -Uri $uri -Method POST -Body $json -ContentType "application/json" -OutputType PSObject
        }
        catch {
            Write-Error $_.Exception 
            break
        }
    }
}

    
Function Remove-AutopilotDevice() {
    <#
.SYNOPSIS
Removes a specific device currently registered with Windows Autopilot.
 
.DESCRIPTION
The Remove-AutopilotDevice cmdlet removes the specified device, identified by its ID, from the list of devices registered with Windows Autopilot for the current Azure AD tenant.
 
.PARAMETER id
Specifies the ID (GUID) for a specific Windows Autopilot device
 
.EXAMPLE
Remove all Windows Autopilot devices from the current Azure AD tenant
 
Get-AutopilotDevice | Remove-AutopilotDevice
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] $id,
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $True)] $serialNumber,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    Begin {
        $bulkList = @()
    }

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/windowsAutopilotDeviceIdentities"    
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"

        try {
            Write-Verbose "DELETE $uri"
            Invoke-MGGraphRequest -Uri $uri -Method DELETE
        }
        catch {
            Write-Error $_.Exception 
            break
        }
        
    }
}


Function Get-AutopilotImportedDevice() {
    <#
.SYNOPSIS
Gets information about devices being imported into Windows Autopilot.
 
.DESCRIPTION
The Get-AutopilotImportedDevice cmdlet retrieves either the full list of devices being imported into Windows Autopilot for the current Azure AD tenant, or information for a specific device if the ID of the device is specified. Once the import is complete, the information instance is expected to be deleted.
 
.PARAMETER id
Optionally specifies the ID (GUID) for a specific Windows Autopilot device being imported.
 
.EXAMPLE
Get a list of all devices being imported into Windows Autopilot for the current Azure AD tenant.
 
Get-AutopilotImportedDevice
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $false)] $id = $null,
        [Parameter(Mandatory = $false)] $serial,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    if ($AppId -ne "") {
        $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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

    # Defining Variables
    $graphApiVersion = "beta"
    if ($id) {
        $uri = "https://graph.microsoft.com/$graphApiVersion/deviceManagement/importedWindowsAutopilotDeviceIdentities/$id"
    } 
    elseif ($serial) {
        # handles also serial numbers with spaces    
        $uri = "https://graph.microsoft.com/$graphApiVersion/deviceManagement/importedWindowsAutopilotDeviceIdentities/?`$filter=contains(serialNumber,'$serial')"
    }
    else {
        $uri = "https://graph.microsoft.com/$graphApiVersion/deviceManagement/importedWindowsAutopilotDeviceIdentities"
    }

    Write-Verbose "GET $uri"

    try {
        $response = Invoke-MGGraphRequest -Uri $uri -Method Get -OutputType PSObject
        if ($id) {
            $response
        }
        else {
            $devices = $response.value
    
            $devicesNextLink = $response."@odata.nextLink"
    
            while ($null -ne $devicesNextLink) {
                $devicesResponse = (Invoke-MGGraphRequest -Uri $devicesNextLink -Method Get -OutputType PSObject)
                $devicesNextLink = $devicesResponse."@odata.nextLink"
                $devices += $devicesResponse.value
            }
    
            $devices
        }
    }
    catch {
        Write-Error $_.Exception 
        break
    }

}


<#
.SYNOPSIS
Adds a new device to Windows Autopilot.
 
.DESCRIPTION
The Add-AutopilotImportedDevice cmdlet adds the specified device to Windows Autopilot for the current Azure AD tenant. Note that a status object is returned when this cmdlet completes; the actual import process is performed as a background batch process by the Microsoft Intune service.
 
.PARAMETER serialNumber
The hardware serial number of the device being added (mandatory).
 
.PARAMETER hardwareIdentifier
The hardware hash (4K string) that uniquely identifies the device.
 
.PARAMETER groupTag
An optional identifier or tag that can be associated with this device, useful for grouping devices using Azure AD dynamic groups.
 
.PARAMETER displayName
The optional name (computer name) to be assigned to the device when it is deployed via Windows Autopilot. This is presently only supported with Azure AD Join scenarios. Note that names should not exceed 15 characters. After setting the name, you need to initiate a sync (Invoke-AutopilotSync) in order to see the name in the Intune object.
 
.PARAMETER assignedUser
The optional user UPN to be assigned to the device. Note that no validation is done on the UPN specified.
 
.EXAMPLE
Add a new device to Windows Autopilot for the current Azure AD tenant.
 
Add-AutopilotImportedDevice -serialNumber $serial -hardwareIdentifier $hash -groupTag "Kiosk" -assignedUser "anna@contoso.com"
#>
Function Add-AutopilotImportedDevice() {
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true)] $serialNumber,
        [Parameter(Mandatory = $true)] $hardwareIdentifier,
        [Parameter(Mandatory = $false)] [Alias("orderIdentifier")] $groupTag = "",
        [Parameter(ParameterSetName = "Prop2")][Alias("UPN")] $assignedUser = "",
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )
    if ($AppId -ne "") {
        $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/importedWindowsAutopilotDeviceIdentities"
    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
    $json = @"
{
    "@odata.type": "#microsoft.graph.importedWindowsAutopilotDeviceIdentity",
    "groupTag": "$groupTag",
    "serialNumber": "$serialNumber",
    "productKey": "",
    "hardwareIdentifier": "$hardwareIdentifier",
    "assignedUserPrincipalName": "$assignedUser",
    "state": {
        "@odata.type": "microsoft.graph.importedWindowsAutopilotDeviceIdentityState",
        "deviceImportStatus": "pending",
        "deviceRegistrationId": "",
        "deviceErrorCode": 0,
        "deviceErrorName": ""
    }
}
"@

    Write-Verbose "POST $uri`n$json"

    try {
        Invoke-MGGraphRequest -Uri $uri -Method Post -body $json -ContentType "application/json"
    }
    catch {
        Write-Error $_.Exception 
        break
    }
    
}

    
Function Remove-AutopilotImportedDevice() {
    <#
.SYNOPSIS
Removes the status information for a device being imported into Windows Autopilot.
 
.DESCRIPTION
The Remove-AutopilotImportedDevice cmdlet cleans up the status information about a new device being imported into Windows Autopilot. This should be done regardless of whether the import was successful or not.
 
.PARAMETER id
The ID (GUID) of the imported device status information to be removed (mandatory).
 
.EXAMPLE
Remove the status information for a specified device.
 
Remove-AutopilotImportedDevice -id $id
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] $id,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/importedWindowsAutopilotDeviceIdentities"    
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"

        try {
            Write-Verbose "DELETE $uri"
            Invoke-MGGraphRequest -Uri $uri -Method DELETE
        }
        catch {
            Write-Error $_.Exception 
            break
        }

    }
        
}


Function Get-AutopilotProfile() {
    <#
.SYNOPSIS
Gets Windows Autopilot profile details.
 
.DESCRIPTION
The Get-AutopilotProfile cmdlet returns either a list of all Windows Autopilot profiles for the current Azure AD tenant, or information for the specific profile specified by its ID.
 
.PARAMETER id
Optionally, the ID (GUID) of the profile to be retrieved.
 
.EXAMPLE
Get a list of all Windows Autopilot profiles.
 
Get-AutopilotProfile
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $false)] $id,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/windowsAutopilotDeploymentProfiles"

    if ($id) {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"
    }
    else {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
    }

    Write-Verbose "GET $uri"

    try {
        $response = Invoke-MGGraphRequest -Uri $uri -Method Get -OutputType PSObject
        if ($id) {
            $response
        }
        else {
            $devices = $response.value
    
            $devicesNextLink = $response."@odata.nextLink"
    
            while ($null -ne $devicesNextLink) {
                $devicesResponse = (Invoke-MGGraphRequest -Uri $devicesNextLink -Method Get -outputType PSObject)
                $devicesNextLink = $devicesResponse."@odata.nextLink"
                $devices += $devicesResponse.value
            }
    
            $devices
        }
    }
    catch {
        Write-Error $_.Exception 
        break
    }

}


Function Get-AutopilotProfileAssignedDevice() {
    <#
.SYNOPSIS
Gets the list of devices that are assigned to the specified Windows Autopilot profile.
 
.DESCRIPTION
The Get-AutopilotProfileAssignedDevice cmdlet returns the list of Autopilot devices that have been assigned the specified Windows Autopilot profile.
 
.PARAMETER id
The ID (GUID) of the profile to be retrieved.
 
.EXAMPLE
Get a list of all Windows Autopilot profiles.
 
Get-AutopilotProfileAssignedDevices -id $id
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $True)] $id,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/windowsAutopilotDeploymentProfiles"
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id/assignedDevices"

        Write-Verbose "GET $uri"

        try {
            $response = Invoke-MGGraphRequest -Uri $uri -Method Get
            $response.Value
        }
        catch {
            Write-Error $_.Exception 
            break
        }
    }
}



Function ConvertTo-AutopilotConfigurationJSON() {
    <#
.SYNOPSIS
Converts the specified Windows Autopilot profile into a JSON format.
 
.DESCRIPTION
The ConvertTo-AutopilotConfigurationJSON cmdlet converts the specified Windows Autopilot profile, as represented by a Microsoft Graph API object, into a JSON format.
 
.PARAMETER profile
A Windows Autopilot profile object, typically returned by Get-AutopilotProfile
 
.EXAMPLE
Get the JSON representation of each Windows Autopilot profile in the current Azure AD tenant.
 
Get-AutopilotProfile | ConvertTo-AutopilotConfigurationJSON
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $True)]
        [Object] $profile,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    Begin {

        # Set the org-related info
        $script:TenantOrg = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -OutputType PSObject).value
        foreach ($domain in $script:TenantOrg.VerifiedDomains) {
            if ($domain.isDefault) {
                $script:TenantDomain = $domain.name
            }
        }
    }

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        $oobeSettings = $profile.outOfBoxExperienceSettings

        # Build up properties
        $json = @{}
        $json.Add("Comment_File", "Profile $($_.displayName)")
        $json.Add("Version", 2049)
        $json.Add("ZtdCorrelationId", $_.id)
        if ($profile."@odata.type" -eq "#microsoft.graph.activeDirectoryWindowsAutopilotDeploymentProfile") {
            $json.Add("CloudAssignedDomainJoinMethod", 1)
        }
        else {
            $json.Add("CloudAssignedDomainJoinMethod", 0)
        }
        if ($profile.deviceNameTemplate) {
            $json.Add("CloudAssignedDeviceName", $_.deviceNameTemplate)
        }

        # Figure out config value
        $oobeConfig = 8 + 256
        if ($oobeSettings.userType -eq 'standard') {
            $oobeConfig += 2
        }
        if ($oobeSettings.hidePrivacySettings -eq $true) {
            $oobeConfig += 4
        }
        if ($oobeSettings.hideEULA -eq $true) {
            $oobeConfig += 16
        }
        if ($oobeSettings.skipKeyboardSelectionPage -eq $true) {
            $oobeConfig += 1024
            if ($_.language) {
                $json.Add("CloudAssignedLanguage", $_.language)
                # Use the same value for region so that screen is skipped too
                $json.Add("CloudAssignedRegion", $_.language)
            }
        }
        if ($oobeSettings.deviceUsageType -eq 'shared') {
            $oobeConfig += 32 + 64
        }
        $json.Add("CloudAssignedOobeConfig", $oobeConfig)

        # Set the forced enrollment setting
        if ($oobeSettings.hideEscapeLink -eq $true) {
            $json.Add("CloudAssignedForcedEnrollment", 1)
        }
        else {
            $json.Add("CloudAssignedForcedEnrollment", 0)
        }

        $json.Add("CloudAssignedTenantId", $script:TenantOrg.id)
        $json.Add("CloudAssignedTenantDomain", $script:TenantDomain)
        $embedded = @{}
        $embedded.Add("CloudAssignedTenantDomain", $script:TenantDomain)
        $embedded.Add("CloudAssignedTenantUpn", "")
        if ($oobeSettings.hideEscapeLink -eq $true) {
            $embedded.Add("ForcedEnrollment", 1)
        }
        else {
            $embedded.Add("ForcedEnrollment", 0)
        }
        $ztc = @{}
        $ztc.Add("ZeroTouchConfig", $embedded)
        $json.Add("CloudAssignedAadServerData", (ConvertTo-JSON $ztc -Compress))

        # Skip connectivity check
        if ($profile.hybridAzureADJoinSkipConnectivityCheck -eq $true) {
            $json.Add("HybridJoinSkipDCConnectivityCheck", 1)
        }

        # Hard-code properties not represented in Intune
        $json.Add("CloudAssignedAutopilotUpdateDisabled", 1)
        $json.Add("CloudAssignedAutopilotUpdateTimeout", 1800000)

        # Return the JSON
        ConvertTo-JSON $json
    }

}


Function Set-AutopilotProfile() {
    <#
.SYNOPSIS
Sets Windows Autopilot profile properties on an existing Autopilot profile.
 
.DESCRIPTION
The Set-AutopilotProfile cmdlet sets properties on an existing Autopilot profile.
 
.PARAMETER id
The GUID of the profile to be updated.
 
.PARAMETER displayName
The name of the Windows Autopilot profile to create. (This value cannot contain spaces.)
 
.PARAMETER description
The description to be configured in the profile. (This value cannot contain dashes.)
 
.PARAMETER ConvertDeviceToAutopilot
Configure the value "Convert all targeted devices to Autopilot"
 
.PARAMETER AllEnabled
Enable everything that can be enabled
 
.PARAMETER AllDisabled
Disable everything that can be disabled
 
.PARAMETER OOBE_HideEULA
Configure the OOBE option to hide or not the EULA
 
.PARAMETER OOBE_EnableWhiteGlove
Configure the OOBE option to allow or not White Glove OOBE
 
.PARAMETER OOBE_HidePrivacySettings
Configure the OOBE option to hide or not the privacy settings
 
.PARAMETER OOBE_HideChangeAccountOpts
Configure the OOBE option to hide or not the change account options
 
.PARAMETER OOBE_UserTypeAdmin
Configure the user account type as administrator.
 
.PARAMETER OOBE_NameTemplate
Configure the OOBE option to apply a device name template
 
.PARAMETER OOBE_language
The language identifier (e.g. "en-us") to be configured in the profile
 
.PARAMETER OOBE_SkipKeyboard
Configure the OOBE option to skip or not the keyboard selection page
 
.PARAMETER OOBE_HideChangeAccountOpts
Configure the OOBE option to hide or not the change account options
 
.PARAMETER OOBE_SkipConnectivityCheck
Specify whether to skip Active Directory connectivity check (UserDrivenAAD only)
 
.EXAMPLE
Update an existing Autopilot profile to specify a language:
 
Set-AutopilotProfile -ID <guid> -Language "en-us"
 
.EXAMPLE
Update an existing Autopilot profile to set multiple properties:
 
Set-AutopilotProfile -ID <guid> -Language "en-us" -displayname "My testing profile" -Description "Description of my profile" -OOBE_HideEULA $True -OOBE_hidePrivacySettings $True
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] $id,
        [Parameter(ParameterSetName = 'notAll')][string] $displayName,
        [Parameter(ParameterSetName = 'notAll')][string] $description,
        [Parameter(ParameterSetName = 'notAll')][Switch] $ConvertDeviceToAutopilot,
        [Parameter(ParameterSetName = 'notAll')][string] $OOBE_language,
        [Parameter(ParameterSetName = 'notAll')][Switch] $OOBE_skipKeyboard,
        [Parameter(ParameterSetName = 'notAll')][string] $OOBE_NameTemplate,
        [Parameter(ParameterSetName = 'notAll')][Switch] $OOBE_EnableWhiteGlove,
        [Parameter(ParameterSetName = 'notAll')][Switch] $OOBE_UserTypeAdmin,
        [Parameter(ParameterSetName = 'AllEnabled', Mandatory = $true)][Switch] $AllEnabled, 
        [Parameter(ParameterSetName = 'AllDisabled', Mandatory = $true)][Switch] $AllDisabled, 
        [Parameter(ParameterSetName = 'notAll')][Switch] $OOBE_HideEULA,
        [Parameter(ParameterSetName = 'notAll')][Switch] $OOBE_hidePrivacySettings,
        [Parameter(ParameterSetName = 'notAll')][Switch] $OOBE_HideChangeAccountOpts,
        [Parameter(ParameterSetName = 'notAll')][Switch] $OOBE_SkipConnectivityCheck,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    # Get the current values
    $current = Get-AutopilotProfile -id $id

    # If this is a Hybrid AADJ profile, make sure it has the needed property
    if ($current.'@odata.type' -eq "#microsoft.graph.azureADWindowsAutopilotDeploymentProfile") {
        if (-not ($current.PSObject.Properties | where-object { $_.Name -eq "hybridAzureADJoinSkipConnectivityCheck" })) {
            $current | Add-Member -NotePropertyName hybridAzureADJoinSkipConnectivityCheck -NotePropertyValue $false
        }
    }

    # For parameters that were specified, update that object in place
    if ($PSBoundParameters.ContainsKey('displayName')) { $current.displayName = $displayName }
    if ($PSBoundParameters.ContainsKey('description')) { $current.description = $description }
    if ($PSBoundParameters.ContainsKey('ConvertDeviceToAutopilot')) { $current.extractHardwareHash = [bool]$ConvertDeviceToAutopilot }
    if ($PSBoundParameters.ContainsKey('OOBE_language')) { $current.language = $OOBE_language }
    if ($PSBoundParameters.ContainsKey('OOBE_skipKeyboard')) { $current.outOfBoxExperienceSettings.skipKeyboardSelectionPage = [bool]$OOBE_skipKeyboard }
    if ($PSBoundParameters.ContainsKey('OOBE_NameTemplate')) { $current.deviceNameTemplate = $OOBE_NameTemplate }
    if ($PSBoundParameters.ContainsKey('OOBE_EnableWhiteGlove')) { $current.enableWhiteGlove = [bool]$OOBE_EnableWhiteGlove }
    if ($PSBoundParameters.ContainsKey('OOBE_UserTypeAdmin')) {
        if ($OOBE_UserTypeAdmin) {
            $current.outOfBoxExperienceSettings.userType = "administrator"
        }
        else {
            $current.outOfBoxExperienceSettings.userType = "standard"
        }
    }
    if ($PSBoundParameters.ContainsKey('OOBE_HideEULA')) { $current.outOfBoxExperienceSettings.hideEULA = [bool]$OOBE_HideEULA }
    if ($PSBoundParameters.ContainsKey('OOBE_HidePrivacySettings')) { $current.outOfBoxExperienceSettings.hidePrivacySettings = [bool]$OOBE_HidePrivacySettings }
    if ($PSBoundParameters.ContainsKey('OOBE_HideChangeAccountOpts')) { $current.outOfBoxExperienceSettings.hideEscapeLink = [bool]$OOBE_HideChangeAccountOpts }
    if ($PSBoundParameters.ContainsKey('OOBE_SkipConnectivityCheck')) { $current.hybridAzureADJoinSkipConnectivityCheck = [bool]$OOBE_SkipConnectivityCheck }

    if ($AllEnabled) {
        $current.extractHardwareHash = $true
        $current.outOfBoxExperienceSettings.hidePrivacySettings = $true
        $current.outOfBoxExperienceSettings.hideEscapeLink = $true
        $current.hybridAzureADJoinSkipConnectivityCheck = $true
        $current.EnableWhiteGlove = $true
        $current.outOfBoxExperienceSettings.hideEULA = $true 
        $current.outOfBoxExperienceSettings.hidePrivacySettings = $true
        $current.outOfBoxExperienceSettings.hideEscapeLink = $true
        $current.outOfBoxExperienceSettings.skipKeyboardSelectionPage = $true
        $current.outOfBoxExperienceSettings.userType = "administrator"
    }
    elseif ($AllDisabled) {
        $current.extractHardwareHash = $false
        $current.outOfBoxExperienceSettings.hidePrivacySettings = $false
        $current.outOfBoxExperienceSettings.hideEscapeLink = $false
        $current.hybridAzureADJoinSkipConnectivityCheck = $false
        $current.EnableWhiteGlove = $false
        $current.outOfBoxExperienceSettings.hideEULA = $false
        $current.outOfBoxExperienceSettings.hidePrivacySettings = $false
        $current.outOfBoxExperienceSettings.hideEscapeLink = $false
        $current.outOfBoxExperienceSettings.skipKeyboardSelectionPage = $false
        $current.outOfBoxExperienceSettings.userType = "standard"
    }

    # Clean up unneeded properties
    $current.PSObject.Properties.Remove("lastModifiedDateTime")
    $current.PSObject.Properties.Remove("createdDateTime") 
    $current.PSObject.Properties.Remove("@odata.context")
    $current.PSObject.Properties.Remove("id")
    $current.PSObject.Properties.Remove("roleScopeTagIds")
    if ($AppId -ne "") {
        $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/windowsAutopilotDeploymentProfiles"
    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"
    $json = ($current | ConvertTo-JSON).ToString()
    
    Write-Verbose "PATCH $uri`n$json"

    try {
        Invoke-MGGraphRequest -Uri $uri -Method PATCH -body $json -ContentType "application/json" -OutputType PSObject
    }
    catch {
        Write-Error $_.Exception 
        break
    }

}


Function New-AutopilotProfile() {
    <#
.SYNOPSIS
Creates a new Autopilot profile.
 
.DESCRIPTION
The New-AutopilotProfile creates a new Autopilot profile.
 
.PARAMETER displayName
The name of the Windows Autopilot profile to create. (This value cannot contain spaces.)
 
.PARAMETER mode
The type of Autopilot profile to create. Choices are "UserDrivenAAD", "UserDrivenAD", and "SelfDeployingAAD".
 
.PARAMETER description
The description to be configured in the profile. (This value cannot contain dashes.)
     
.PARAMETER ConvertDeviceToAutopilot
Configure the value "Convert all targeted devices to Autopilot"
 
.PARAMETER OOBE_HideEULA
Configure the OOBE option to hide or not the EULA
 
.PARAMETER OOBE_EnableWhiteGlove
Configure the OOBE option to allow or not White Glove OOBE
 
.PARAMETER OOBE_HidePrivacySettings
Configure the OOBE option to hide or not the privacy settings
 
.PARAMETER OOBE_HideChangeAccountOpts
Configure the OOBE option to hide or not the change account options
 
.PARAMETER OOBE_UserTypeAdmin
Configure the user account type as administrator.
 
.PARAMETER OOBE_NameTemplate
Configure the OOBE option to apply a device name template
 
.PARAMETER OOBE_language
The language identifier (e.g. "en-us") to be configured in the profile
 
.PARAMETER OOBE_SkipKeyboard
Configure the OOBE option to skip or not the keyboard selection page
 
.PARAMETER OOBE_HideChangeAccountOpts
Configure the OOBE option to hide or not the change account options
 
.PARAMETER OOBE_SkipConnectivityCheck
Specify whether to skip Active Directory connectivity checks (UserDrivenAAD only)
 
.EXAMPLE
Create profiles of different types:
 
New-AutopilotProfile -mode UserDrivenAAD -displayName "My AAD profile" -description "My user-driven AAD profile" -OOBE_Quiet
New-AutopilotProfile -mode UserDrivenAD -displayName "My AD profile" -description "My user-driven AD profile" -OOBE_Quiet
New-AutopilotProfile -mode SelfDeployingAAD -displayName "My Self Deploying profile" -description "My self-deploying profile" -OOBE_Quiet
 
.EXAMPLE
Create a user-driven AAD profile:
 
New-AutopilotProfile -mode UserDrivenAAD -displayName "My testing profile" -Description "Description of my profile" -OOBE_Language "en-us" -OOBE_HideEULA -OOBE_HidePrivacySettings
 
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true)][string] $displayName,
        [Parameter(Mandatory = $true)][ValidateSet('UserDrivenAAD', 'UserDrivenAD', 'SelfDeployingAAD')][string] $mode, 
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret,
        [string] $description,
        [Switch] $ConvertDeviceToAutopilot,
        [string] $OOBE_language,
        [Switch] $OOBE_skipKeyboard,
        [string] $OOBE_NameTemplate,
        [Switch] $OOBE_EnableWhiteGlove,
        [Switch] $OOBE_UserTypeAdmin,
        [Switch] $OOBE_HideEULA,
        [Switch] $OOBE_hidePrivacySettings,
        [Switch] $OOBE_HideChangeAccountOpts,
        [Switch] $OOBE_SkipConnectivityCheck
    )

    # Adjust values as needed
    switch ($mode) {
        "UserDrivenAAD" { $odataType = "#microsoft.graph.azureADWindowsAutopilotDeploymentProfile"; $usage = "singleUser" }
        "SelfDeployingAAD" { $odataType = "#microsoft.graph.azureADWindowsAutopilotDeploymentProfile"; $usage = "shared" }
        "UserDrivenAD" { $odataType = "#microsoft.graph.activeDirectoryWindowsAutopilotDeploymentProfile"; $usage = "singleUser" }
    }

    if ($OOBE_UserTypeAdmin) {        
        $OOBE_userType = "administrator"
    }
    else {        
        $OOBE_userType = "standard"
    }        

    if ($OOBE_EnableWhiteGlove) {        
        $OOBE_HideChangeAccountOpts = $True
    }        
    if ($AppId -ne "") {
        $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/windowsAutopilotDeploymentProfiles"
    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
    if ($mode -eq "UserDrivenAD") {
        $json = @"
{
    "@odata.type": "$odataType",
    "displayName": "$displayname",
    "description": "$description",
    "language": "$OOBE_language",
    "extractHardwareHash": $(BoolToString($ConvertDeviceToAutopilot)),
    "deviceNameTemplate": "$OOBE_NameTemplate",
    "deviceType": "windowsPc",
    "enableWhiteGlove": $(BoolToString($OOBE_EnableWhiteGlove)),
    "hybridAzureADJoinSkipConnectivityCheck": $(BoolToString($OOBE_SkipConnectivityChecks)),
    "outOfBoxExperienceSettings": {
        "hidePrivacySettings": $(BoolToString($OOBE_hidePrivacySettings)),
        "hideEULA": $(BoolToString($OOBE_HideEULA)),
        "userType": "$OOBE_userType",
        "deviceUsageType": "$usage",
        "skipKeyboardSelectionPage": $(BoolToString($OOBE_skipKeyboard)),
        "hideEscapeLink": $(BoolToString($OOBE_HideChangeAccountOpts))
    }
}
"@
    }
    else {
        $json = @"
{
    "@odata.type": "$odataType",
    "displayName": "$displayname",
    "description": "$description",
    "language": "$OOBE_language",
    "extractHardwareHash": $(BoolToString($ConvertDeviceToAutopilot)),
    "deviceNameTemplate": "$OOBE_NameTemplate",
    "deviceType": "windowsPc",
    "enableWhiteGlove": $(BoolToString($OOBE_EnableWhiteGlove)),
    "outOfBoxExperienceSettings": {
        "hidePrivacySettings": $(BoolToString($OOBE_hidePrivacySettings)),
        "hideEULA": $(BoolToString($OOBE_HideEULA)),
        "userType": "$OOBE_userType",
        "deviceUsageType": "$usage",
        "skipKeyboardSelectionPage": $(BoolToString($OOBE_skipKeyboard)),
        "hideEscapeLink": $(BoolToString($OOBE_HideChangeAccountOpts))
    }
}
"@
    }

    Write-Verbose "POST $uri`n$json"

    try {
        Invoke-MGGraphRequest -Uri $uri -Method POST -Body $json -ContentType "application/json" -OutputType PSObject
    }
    catch {
        Write-Error $_.Exception 
        break
    }

}


Function Remove-AutopilotProfile() {
    <#
.SYNOPSIS
Remove a Deployment Profile
.DESCRIPTION
The Remove-AutopilotProfile allows you to remove a specific deployment profile
.PARAMETER id
Mandatory, the ID (GUID) of the profile to be removed.
.EXAMPLE
Remove-AutopilotProfile -id $id
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)] $id,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/windowsAutopilotDeploymentProfiles"
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"

        Write-Verbose "DELETE $uri"

        Try {
            Invoke-MGGraphRequest -Uri $uri -Method DELETE
        }
        catch {
            Write-Error $_.Exception 
            break
        }
    }
}


Function Get-AutopilotProfileAssignments() {
    <#
.SYNOPSIS
List all assigned devices for a specific profile ID
.DESCRIPTION
The Get-AutopilotProfileAssignments cmdlet returns the list of groups that ae assigned to a spcific deployment profile
.PARAMETER id
Type: Integer - Mandatory, the ID (GUID) of the profile to be retrieved.
.EXAMPLE
Get-AutopilotProfileAssignments -id $id
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] $id,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/windowsAutopilotDeploymentProfiles"
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id/assignments"

        Write-Verbose "GET $uri"

        try {
            $response = Invoke-MGGraphRequest -Uri $uri -Method Get
            $Group_ID = $response.Value.target.groupId
            ForEach ($Group in $Group_ID) {
                Try {
                    Get-MgGroup | where-object { $_.ObjectId -like $Group }
                }
                Catch {
                    $Group
                }            
            }
        }
        catch {
            Write-Error $_.Exception 
            break
        }

    }

}


Function Remove-AutopilotProfileAssignments() {
    <#
.SYNOPSIS
Removes a specific group assigntion for a specifc deployment profile
.DESCRIPTION
The Remove-AutopilotProfileAssignments cmdlet allows you to remove a group assignation for a deployment profile
.PARAMETER id
Type: Integer - Mandatory, the ID (GUID) of the profile
.PARAMETER groupid
Type: Integer - Mandatory, the ID of the group
.EXAMPLE
Remove-AutopilotProfileAssignments -id $id
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true)]$id,
        [Parameter(Mandatory = $true)]$groupid,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )
    if ($AppId -ne "") {
        $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/windowsAutopilotDeploymentProfiles"
    
    $full_assignment_id = $id + "_" + $groupid + "_0"

    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id/assignments/$full_assignment_id"

    Write-Verbose "DELETE $uri"

    try {
        Invoke-MGGraphRequest -Uri $uri -Method DELETE
    }
    catch {
        Write-Error $_.Exception 
        break
    }

}


Function Set-AutopilotProfileAssignedGroup() {
    <#
.SYNOPSIS
Assigns a group to a Windows Autopilot profile.
.DESCRIPTION
The Set-AutopilotProfileAssignedGroup cmdlet allows you to assign a specific group to a specific deployment profile
.PARAMETER id
Type: Integer - Mandatory, the ID (GUID) of the profile
.PARAMETER groupid
Type: Integer - Mandatory, the ID of the group
.EXAMPLE
Set-AutopilotProfileAssignedGroup -id $id -groupid $groupid
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true)]$id,
        [Parameter(Mandatory = $true)]$groupid,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )
    $full_assignment_id = $id + "_" + $groupid + "_0"  
    if ($AppId -ne "") {
        $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/windowsAutopilotDeploymentProfiles"        
    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id/assignments"        

    $json = @"
{
    "id": "$full_assignment_id",
    "target": {
        "@odata.type": "#microsoft.graph.groupAssignmentTarget",
        "groupId": "$groupid"
    }
}
"@

    Write-Verbose "POST $uri`n$json"

    try {
        Invoke-MGGraphRequest -Uri $uri -Method Post -Body $json -ContentType "application/json" -OutputType PSObject
    }
    catch {
        Write-Error $_.Exception 
        break
    }
}


Function Get-EnrollmentStatusPage() {
    <#
.SYNOPSIS
List enrollment status page
.DESCRIPTION
The Get-EnrollmentStatusPage cmdlet returns available enrollment status page with their options
.PARAMETER id
The ID (GUID) of the status page (optional)
.EXAMPLE
Get-EnrollmentStatusPage
#>

    [cmdletbinding()]
    param
    (
        [Parameter()] $id,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )
    if ($AppId -ne "") {
        $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/deviceEnrollmentConfigurations"

    if ($id) {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"
    }
    else {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
    }

    Write-Verbose "GET $uri"

    try {
        $response = Invoke-MGGraphRequest -Uri $uri -Method Get -OutputType PSObject
        if ($id) {
            $response
        }
        else {
            $response.Value | where-object { $_.'@odata.type' -eq "#microsoft.graph.windows10EnrollmentCompletionPageConfiguration" }
        }
    }
    catch {
        Write-Error $_.Exception 
        break
    }

}


Function Add-EnrollmentStatusPage() {
    <#
.SYNOPSIS
Adds a new Windows Autopilot Enrollment Status Page.
.DESCRIPTION
The Add-EnrollmentStatusPage cmdlet sets properties on an existing Autopilot profile.
.PARAMETER DisplayName
Type: String - Configure the display name of the enrollment status page
.PARAMETER description
Type: String - Configure the description of the enrollment status page
.PARAMETER HideProgress
Type: Boolean - Configure the option: Show app and profile installation progress
.PARAMETER AllowCollectLogs
Type: Boolean - Configure the option: Allow users to collect logs about installation errors
.PARAMETER Message
Type: String - Configure the option: Show custom message when an error occurs
.PARAMETER AllowUseOnFailure
Type: Boolean - Configure the option: Allow users to use device if installation error occurs
.PARAMETER AllowResetOnError
Type: Boolean - Configure the option: Allow users to reset device if installation error occurs
.PARAMETER BlockDeviceUntilComplete
Type: Boolean - Configure the option: Block device use until all apps and profiles are installed
.PARAMETER TimeoutInMinutes
Type: Integer - Configure the option: Show error when installation takes longer than specified number of minutes
.EXAMPLE
Add-EnrollmentStatusPage -Message "Oops an error occured, please contact your support" -HideProgress $True -AllowResetOnError $True
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $True)][string]$DisplayName,
        [string]$Description,        
        [bool]$HideProgress,    
        [bool]$AllowCollectLogs,
        [bool]$blockDeviceSetupRetryByUser,    
        [string]$Message,    
        [bool]$AllowUseOnFailure,
        [bool]$AllowResetOnError,    
        [bool]$BlockDeviceUntilComplete,                
        [Int]$TimeoutInMinutes,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret        
    )
    if ($AppId -ne "") {
        $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
    If ($HideProgress -eq $False) {
        $blockDeviceSetupRetryByUser = $true
    }

    If (($Description -eq $null)) {
        $Description = $EnrollmentPage_Description
    }        

    If (($DisplayName -eq $null)) {
        $DisplayName = ""
    }    

    If (($TimeoutInMinutes -eq "")) {
        $TimeoutInMinutes = "60"
    }                

    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/deviceEnrollmentConfigurations"
    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
    $json = @"
{
    "@odata.type": "#microsoft.graph.windows10EnrollmentCompletionPageConfiguration",
    "displayName": "$DisplayName",
    "description": "$description",
    "showInstallationProgress": "$hideprogress",
    "blockDeviceSetupRetryByUser": "$blockDeviceSetupRetryByUser",
    "allowDeviceResetOnInstallFailure": "$AllowResetOnError",
    "allowLogCollectionOnInstallFailure": "$AllowCollectLogs",
    "customErrorMessage": "$Message",
    "installProgressTimeoutInMinutes": "$TimeoutInMinutes",
    "allowDeviceUseOnInstallFailure": "$AllowUseOnFailure",
}
"@

    Write-Verbose "POST $uri`n$json"

    try {
        Invoke-MgGraphRequest -Uri $uri -Method Post -Body $json -ContentType "application/json" -OutputType PSObject
    }
    catch {
        Write-Error $_.Exception 
        break
    }

}


Function Set-EnrollmentStatusPage() {
    <#
.SYNOPSIS
Sets Windows Autopilot Enrollment Status Page properties.
.DESCRIPTION
The Set-EnrollmentStatusPage cmdlet sets properties on an existing Autopilot profile.
.PARAMETER id
The ID (GUID) of the profile to be updated.
.PARAMETER DisplayName
Type: String - Configure the display name of the enrollment status page
.PARAMETER description
Type: String - Configure the description of the enrollment status page
.PARAMETER HideProgress
Type: Boolean - Configure the option: Show app and profile installation progress
.PARAMETER AllowCollectLogs
Type: Boolean - Configure the option: Allow users to collect logs about installation errors
.PARAMETER Message
Type: String - Configure the option: Show custom message when an error occurs
.PARAMETER AllowUseOnFailure
Type: Boolean - Configure the option: Allow users to use device if installation error occurs
.PARAMETER AllowResetOnError
Type: Boolean - Configure the option: Allow users to reset device if installation error occurs
.PARAMETER BlockDeviceUntilComplete
Type: Boolean - Configure the option: Block device use until all apps and profiles are installed
.PARAMETER TimeoutInMinutes
Type: Integer - Configure the option: Show error when installation takes longer than specified number of minutes
.EXAMPLE
Set-EnrollmentStatusPage -id $id -Message "Oops an error occured, please contact your support" -HideProgress $True -AllowResetOnError $True
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] $id,
        [string]$DisplayName,    
        [string]$Description,        
        [bool]$HideProgress,
        [bool]$AllowCollectLogs,
        [string]$Message,    
        [bool]$AllowUseOnFailure,
        [bool]$AllowResetOnError,    
        [bool]$AllowUseOnError,    
        [bool]$BlockDeviceUntilComplete,                
        [Int]$TimeoutInMinutes,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret        
    )

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        # LIST EXISTING VALUES FOR THE SELECTING STAUS PAGE
        # Default profile values
        $EnrollmentPage_Values = Get-EnrollmentStatusPage -ID $id
        $EnrollmentPage_DisplayName = $EnrollmentPage_Values.displayName
        $EnrollmentPage_Description = $EnrollmentPage_Values.description
        $EnrollmentPage_showInstallationProgress = $EnrollmentPage_Values.showInstallationProgress
        $EnrollmentPage_blockDeviceSetupRetryByUser = $EnrollmentPage_Values.blockDeviceSetupRetryByUser
        $EnrollmentPage_allowDeviceResetOnInstallFailure = $EnrollmentPage_Values.allowDeviceResetOnInstallFailure
        $EnrollmentPage_allowLogCollectionOnInstallFailure = $EnrollmentPage_Values.allowLogCollectionOnInstallFailure
        $EnrollmentPage_customErrorMessage = $EnrollmentPage_Values.customErrorMessage
        $EnrollmentPage_installProgressTimeoutInMinutes = $EnrollmentPage_Values.installProgressTimeoutInMinutes
        $EnrollmentPage_allowDeviceUseOnInstallFailure = $EnrollmentPage_Values.allowDeviceUseOnInstallFailure

        If (!($HideProgress)) {
            $HideProgress = $EnrollmentPage_showInstallationProgress
        }    
    
        If (!($BlockDeviceUntilComplete)) {
            $BlockDeviceUntilComplete = $EnrollmentPage_blockDeviceSetupRetryByUser
        }        
        
        If (!($AllowCollectLogs)) {
            $AllowCollectLogs = $EnrollmentPage_allowLogCollectionOnInstallFailure
        }            
    
        If (!($AllowUseOnFailure)) {
            $AllowUseOnFailure = $EnrollmentPage_allowDeviceUseOnInstallFailure
        }    

        If (($Message -eq "")) {
            $Message = $EnrollmentPage_customErrorMessage
        }        
        
        If (($Description -eq $null)) {
            $Description = $EnrollmentPage_Description
        }        

        If (($DisplayName -eq $null)) {
            $DisplayName = $EnrollmentPage_DisplayName
        }    

        If (!($AllowResetOnError)) {
            $AllowResetOnError = $EnrollmentPage_allowDeviceResetOnInstallFailure
        }    

        If (($TimeoutInMinutes -eq "")) {
            $TimeoutInMinutes = $EnrollmentPage_installProgressTimeoutInMinutes
        }                

        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/deviceEnrollmentConfigurations"
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"
        $json = @"
{
    "@odata.type": "#microsoft.graph.windows10EnrollmentCompletionPageConfiguration",
    "displayName": "$DisplayName",
    "description": "$description",
    "showInstallationProgress": "$HideProgress",
    "blockDeviceSetupRetryByUser": "$BlockDeviceUntilComplete",
    "allowDeviceResetOnInstallFailure": "$AllowResetOnError",
    "allowLogCollectionOnInstallFailure": "$AllowCollectLogs",
    "customErrorMessage": "$Message",
    "installProgressTimeoutInMinutes": "$TimeoutInMinutes",
    "allowDeviceUseOnInstallFailure": "$AllowUseOnFailure"
}
"@

        Write-Verbose "PATCH $uri`n$json"

        try {
            Invoke-MgGraphRequest -Uri $uri -Method PATCH -body $json -ContentType "application/json" -OutputType PSObject
        }
        catch {
            Write-Error $_.Exception 
            break
        }

    }

}


Function Remove-EnrollmentStatusPage() {
    <#
.SYNOPSIS
Remove a specific enrollment status page
.DESCRIPTION
The Remove-EnrollmentStatusPage allows you to remove a specific enrollment status page
.PARAMETER id
Mandatory, the ID (GUID) of the profile to be retrieved.
.EXAMPLE
Remove-EnrollmentStatusPage -id $id
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)] $id,
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/deviceEnrollmentConfigurations"
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"

        Write-Verbose "DELETE $uri"

        try {
            Invoke-MgGraphRequest -Uri $uri -Method DELETE
        }
        catch {
            Write-Error $_.Exception 
            break
        }

    }

}


Function Invoke-AutopilotSync() {
    <#
.SYNOPSIS
Initiates a synchronization of Windows Autopilot devices between the Autopilot deployment service and Intune.
 
.DESCRIPTION
The Invoke-AutopilotSync cmdlet initiates a synchronization between the Autopilot deployment service and Intune.
This can be done after importing new devices, to ensure that they appear in Intune in the list of registered
Autopilot devices. See https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/intune_enrollment_windowsautopilotsettings_sync
for more information.
 
.EXAMPLE
Initiate a synchronization.
 
Invoke-AutopilotSync
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )
    if ($AppId -ne "") {
        $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/windowsAutopilotSettings/sync"
    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"

    Write-Verbose "POST $uri"

    try {
        Invoke-MgGraphRequest -Uri $uri -Method Post
    }
    catch {
        Write-Error $_.Exception 
        break
    }

}

Function Get-AutopilotSyncInfo() {
    <#
    .SYNOPSIS
    Returns details about the last Autopilot sync.
     
    .DESCRIPTION
    The Get-AutopilotSyncInfo cmdlet retrieves details about the sync status between Intune and the Autopilot service.
    See https://docs.microsoft.com/en-us/graph/api/resources/intune-enrollment-windowsautopilotsettings?view=graph-rest-beta
    for more information.
     
    .EXAMPLE
    Get-AutopilotSyncInfo
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )
    if ($AppId -ne "") {
        $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/windowsAutopilotSettings"
    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
    
    Write-Verbose "GET $uri"
    
    try {
        Invoke-MGGraphRequest -Uri $uri -Method Get -OutputType PSObject
    }
    catch {
        Write-Error $_.Exception 
        break
    }
    
}
    
#endregion


Function Import-AutopilotCSV() {
    <#
.SYNOPSIS
Adds a batch of new devices into Windows Autopilot.
 
.DESCRIPTION
The Import-AutopilotCSV cmdlet processes a list of new devices (contained in a CSV file) using a several of the other cmdlets included in this module. It is a convenient wrapper to handle the details. After the devices have been added, the cmdlet will continue to check the status of the import process. Once all devices have been processed (successfully or not) the cmdlet will complete. This can take several minutes, as the devices are processed by Intune as a background batch process.
 
.PARAMETER csvFile
The file containing the list of devices to be added.
 
.PARAMETER groupTag
An optional identifier or tag that can be associated with this device, useful for grouping devices using Azure AD dynamic groups. This value overrides an Group Tag value specified in the CSV file.
 
.EXAMPLE
Add a batch of devices to Windows Autopilot for the current Azure AD tenant.
 
Import-AutopilotCSV -csvFile C:\Devices.csv
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true)] $csvFile,
        [Parameter(Mandatory = $false)] [Alias("orderIdentifier")] $groupTag = "",
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )
    
    # Read CSV and process each device
    $devices = Import-CSV $csvFile
    $importedDevices = @()
    foreach ($device in $devices) {
        if ($groupTag -ne "") {
            $o = $groupTag
        }
        elseif ($device.'Group Tag' -ne "") {
            $o = $device.'Group Tag'
        }
        else {
            $o = $device.'OrderID'
        }
        Add-AutopilotImportedDevice -serialNumber $device.'Device Serial Number' -hardwareIdentifier $device.'Hardware Hash' -groupTag $o -assignedUser $device.'Assigned User'
    }

    # While we could keep a list of all the IDs that we added and then check each one, it is
    # easier to just loop through all of them
    $processingCount = 1
    while ($processingCount -gt 0) {
        $deviceStatuses = @(Get-AutopilotImportedDevice)
        $deviceCount = $deviceStatuses.Length

        # Check to see if any devices are still processing
        $processingCount = 0
        foreach ($device in $deviceStatuses) {
            if ($device.state.deviceImportStatus -eq "unknown") {
                $processingCount = $processingCount + 1
            }
        }
        Write-Host "Waiting for $processingCount of $deviceCount"

        # Still processing? Sleep before trying again.
        if ($processingCount -gt 0) {
            Start-Sleep 15
        }
    }

    # Display the statuses
    $deviceStatuses | ForEach-Object {
        Write-Host "Serial number $($_.serialNumber): $($_.state.deviceImportStatus) $($_.state.deviceErrorCode) $($_.state.deviceErrorName)"
    }

    # Cleanup the imported device records
    $deviceStatuses | ForEach-Object {
        Remove-AutopilotImportedDevice -id $_.id
    }
}


Function Get-AutopilotEvent() {
    <#
.SYNOPSIS
Gets Windows Autopilot deployment events.
 
.DESCRIPTION
The Get-AutopilotEvent cmdlet retrieves the list of deployment events (the data that you would see in the "Autopilot deployments" report in the Intune portal).
 
.EXAMPLE
Get a list of all Windows Autopilot events
 
Get-AutopilotEvent
#>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret
    )

    Process {
        if ($AppId -ne "") {
            $graph = Connect-ToGraph -Tenant $Tenant -AppId $AppId -AppSecret $AppSecret
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
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/autopilotEvents"
        $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"

        try {
            $response = Invoke-MgGraphRequest -Uri $uri -Method Get -OutputType PSObject
            $devices = $response.value
            $devicesNextLink = $response."@odata.nextLink"
    
            while ($null -ne $devicesNextLink) {
                $devicesResponse = (Invoke-MgGraphRequest -Uri $devicesNextLink -Method Get -OutputType PSObject)
                $devicesNextLink = $devicesResponse."@odata.nextLink"
                $devices += $devicesResponse.value
            }
    
            $devices
        }
        catch {
            Write-Error $_.Exception 
            break
        }
    }
}
# SIG # Begin signature block
# MIIoGQYJKoZIhvcNAQcCoIIoCjCCKAYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA+pBqDr8ewaM34
# 4cpYoxfnuuYTLknkvXGljfcn4BP5S6CCIRwwggWNMIIEdaADAgECAhAOmxiO+dAt
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
# gjcCARUwLwYJKoZIhvcNAQkEMSIEIDZuIY599+skX09KgTHGS7k5ltd1uy89SrEB
# lkiRHTQIMA0GCSqGSIb3DQEBAQUABIICAJ0mfSXnomEi8U1i8v7TINjVhCKpTgNn
# j3J+dOCuIrqNVxbGp9Mbrj/m+l7h42A1CLevQQUO7A4diSL2cNyE1SxuFpGqVX4x
# 1jE3ZCMW5KNtpeLwMVRr607V7DTkXfx3I2hV1yDHerL2pjq3kKnBL3fTzcRcOe4L
# wyNH9T9DYYfnbHZiLBiGywbKYDCVbMpcCQasNmA4oMmvRwt22g7EO0AW/riucEnP
# VJLpGuSm9OZHOW9CE6TJhPqs4mlA11KA2HCnwnvebYBifhVFbLjsoEOZM2t5QmqM
# zo5rLQ95eNXuFYDSpbjyZbkALzkr67QBFpTQltRPnGiNY3XHR09B8PPuG7CVUGWB
# sptCi/XszcK51L8nUuREews9pmNtL5GKXT4p3/8fZxU8gkGrmaMU2aI8yCh4eDhA
# qBLr+mGI2R7vvtm3hZO/TbrsGtAfDQ9fJvYQ4F50ssXP1vVYbQa8qvvXxFIQAr7i
# CnxyuWBp8uHDUOTbyJjHrjkaO6lw4XhkbE1TJeh+eynH7iWGO2WBqs/gsUmd/uGN
# 74Nzrngj/rKBHcx1F+kINA9JTc2DZ6E0JRftAGqdByx9+w3r/HWi1ERupWwZRlJB
# obq3+yAYZ0i46k4xJlyNjTUL1O/M0BdbAnx3KuM0xOJEDhv23smBNhrbTRwNDGm7
# 9hlbcDVQ6rLpoYIDIDCCAxwGCSqGSIb3DQEJBjGCAw0wggMJAgEBMHcwYzELMAkG
# A1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdp
# Q2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQQIQ
# BUSv85SdCDmmv9s/X+VhFjANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzEL
# BgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIzMTExNzA5MTEzNFowLwYJKoZI
# hvcNAQkEMSIEIFnCondBZcGIMPm6wAyjxMivAiCMwJLOZ3lTFg2N7nFsMA0GCSqG
# SIb3DQEBAQUABIICACX7PihfpLsnWaBqhmmOVZHxgSohQg0ElsSQ+zpNQQJJH2kM
# sfSGg1Eq+7x9n+OpGjQzFBFqcfpu+iIPMi1sx7fPBGxiRmO+Yo+4fx09ItbCZWrb
# EGg2HRLa0Va3SBCsAcimTkgqv2BLU+Kk/pQqrRGrJoq//3h9WO0tdDXVgWovCVOl
# +8ZaTUszaICDSHylJaVRLhkv7PPWQcrPhJi0MHhuKULcpKXBjBCqty8rzj0Zb1mM
# qagngsZVO6dzieRRhrZz4g/ttLRxJBneSI/OxQxUNpe1985jPAz8RQPw2DMCe1wC
# tmF4q6h6mrT0bluwlZ8sOuer+0WWtsTa1L3De6XpTQcHjWXtd99KXlnWtpbe6Vxk
# bxW0Hrb43lWsMzkiShUCSkOuKp9CIycPIkQoUKbDeUp7VYeh4fOLVbpl+HLGyrYG
# kaFSsukOASmxo05zlY0B9wzLKuItIx40Viym1Ja7GyIJNDvLzKgp3WK+3EVwQ+Ik
# 9FN5aPSY5BxtfGYfA1xEPOje51WJD3RZA27p25SdpUTsEgodhmsLS66nE93HsLnT
# Wit8t5NT6p1nJnuCZif9/BHsNoR0Zqrhiz8pyBAPgJRHLa0a8FLpMdTe08V4YW3+
# kNvJyWeOm4Hcw3VQhbojyDMrL43ffXz+BAPmEoOW9RntHJqerwltYP2h+Zui
# SIG # End signature block
