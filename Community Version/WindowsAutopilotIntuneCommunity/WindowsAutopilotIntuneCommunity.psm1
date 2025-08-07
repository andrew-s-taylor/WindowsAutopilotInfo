
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
    "hybridAzureADJoinSkipConnectivityCheck": $(BoolToString($OOBE_SkipConnectivityCheck)),
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
                    #Get-MgGroup | where-object { $_.ObjectId -like $Group }
                    $guri = "https://graph.microsoft.com/beta/groups?`$filter=id eq '$Group'"
                    (Invoke-MgGraphRequest -Uri $guri -Method GET -OutputType PSObject).value
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
            $graph = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            if ($AddToGroup) {
                $aadId = Connect-ToGraph -scopes "Group.ReadWrite.All, Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All, GroupMember.ReadWrite.All, DeviceManagementScripts.ReadWrite.All"
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
# MIIoUAYJKoZIhvcNAQcCoIIoQTCCKD0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC1gaTUlrBQRsHY
# dVL0LJmFqXYuLy+4HiBEC8efEgQcTaCCIU0wggWNMIIEdaADAgECAhAOmxiO+dAt
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
# twGpn1eqXijiuZQwggawMIIEmKADAgECAhAIrUCyYNKcTJ9ezam9k67ZMA0GCSqG
# SIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0zNjA0MjgyMzU5NTlaMGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEzODQg
# MjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDVtC9C0Cit
# eLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0JAfhS0/TeEP0F9ce2vnS
# 1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJrQ5qZ8sU7H/Lvy0daE6ZM
# swEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhFLqGfLOEYwhrMxe6TSXBC
# Mo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+FLEikVoQ11vkunKoAFdE3
# /hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh3K3kGKDYwSNHR7OhD26j
# q22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJwZPt4bRc4G/rJvmM1bL5
# OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQayg9Rc9hUZTO1i4F4z8ujo
# 7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbIYViY9XwCFjyDKK05huzU
# tw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchApQfDVxW0mdmgRQRNYmtwm
# KwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRroOBl8ZhzNeDhFMJlP/2NP
# TLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IBWTCCAVUwEgYDVR0TAQH/
# BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHwYDVR0j
# BBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0
# cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8E
# PDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVz
# dGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATANBgkq
# hkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql+Eg08yy25nRm95RysQDK
# r2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFFUP2cvbaF4HZ+N3HLIvda
# qpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1hmYFW9snjdufE5BtfQ/g+
# lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3RywYFzzDaju4ImhvTnhOE7a
# brs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5UbdldAhQfQDN8A+KVssIhdXNS
# y0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw8MzK7/0pNVwfiThV9zeK
# iwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnPLqR0kq3bPKSchh/jwVYb
# KyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatEQOON8BUozu3xGFYHKi8Q
# xAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bnKD+sEq6lLyJsQfmCXBVm
# zGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQjiWQ1tygVQK+pKHJ6l/aCn
# HwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbqyK+p/pQd52MbOoZWeE4w
# gga0MIIEnKADAgECAhANx6xXBf8hmS5AQyIMOkmGMA0GCSqGSIb3DQEBCwUAMGIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBH
# NDAeFw0yNTA1MDcwMDAwMDBaFw0zODAxMTQyMzU5NTlaMGkxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1
# c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC0eDHTCphBcr48RsAcrHXbo0Zo
# dLRRF51NrY0NlLWZloMsVO1DahGPNRcybEKq+RuwOnPhof6pvF4uGjwjqNjfEvUi
# 6wuim5bap+0lgloM2zX4kftn5B1IpYzTqpyFQ/4Bt0mAxAHeHYNnQxqXmRinvuNg
# xVBdJkf77S2uPoCj7GH8BLuxBG5AvftBdsOECS1UkxBvMgEdgkFiDNYiOTx4OtiF
# cMSkqTtF2hfQz3zQSku2Ws3IfDReb6e3mmdglTcaarps0wjUjsZvkgFkriK9tUKJ
# m/s80FiocSk1VYLZlDwFt+cVFBURJg6zMUjZa/zbCclF83bRVFLeGkuAhHiGPMvS
# GmhgaTzVyhYn4p0+8y9oHRaQT/aofEnS5xLrfxnGpTXiUOeSLsJygoLPp66bkDX1
# ZlAeSpQl92QOMeRxykvq6gbylsXQskBBBnGy3tW/AMOMCZIVNSaz7BX8VtYGqLt9
# MmeOreGPRdtBx3yGOP+rx3rKWDEJlIqLXvJWnY0v5ydPpOjL6s36czwzsucuoKs7
# Yk/ehb//Wx+5kMqIMRvUBDx6z1ev+7psNOdgJMoiwOrUG2ZdSoQbU2rMkpLiQ6bG
# RinZbI4OLu9BMIFm1UUl9VnePs6BaaeEWvjJSjNm2qA+sdFUeEY0qVjPKOWug/G6
# X5uAiynM7Bu2ayBjUwIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAd
# BgNVHQ4EFgQU729TSunkBnx6yuKQVvYv1Ensy04wHwYDVR0jBBgwFoAU7NfjgtJx
# XWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUF
# BwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGln
# aWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJo
# dHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNy
# bDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQEL
# BQADggIBABfO+xaAHP4HPRF2cTC9vgvItTSmf83Qh8WIGjB/T8ObXAZz8OjuhUxj
# aaFdleMM0lBryPTQM2qEJPe36zwbSI/mS83afsl3YTj+IQhQE7jU/kXjjytJgnn0
# hvrV6hqWGd3rLAUt6vJy9lMDPjTLxLgXf9r5nWMQwr8Myb9rEVKChHyfpzee5kH0
# F8HABBgr0UdqirZ7bowe9Vj2AIMD8liyrukZ2iA/wdG2th9y1IsA0QF8dTXqvcnT
# mpfeQh35k5zOCPmSNq1UH410ANVko43+Cdmu4y81hjajV/gxdEkMx1NKU4uHQcKf
# ZxAvBAKqMVuqte69M9J6A47OvgRaPs+2ykgcGV00TYr2Lr3ty9qIijanrUR3anzE
# wlvzZiiyfTPjLbnFRsjsYg39OlV8cipDoq7+qNNjqFzeGxcytL5TTLL4ZaoBdqbh
# OhZ3ZRDUphPvSRmMThi0vw9vODRzW6AxnJll38F0cuJG7uEBYTptMSbhdhGQDpOX
# gpIUsWTjd6xpR6oaQf/DJbg3s6KCLPAlZ66RzIg9sC+NJpud/v4+7RWsWCiKi9EO
# LLHfMR2ZyJ/+xhCx9yHbxtl5TPau1j/1MIDpMPx0LckTetiSuEtQvLsNz3Qbp7wG
# WqbIiOWCnb5WqxL3/BAPvIXKUjPSxyZsq8WhbaM2tszWkPZPubdcMIIG7TCCBNWg
# AwIBAgIQCoDvGEuN8QWC0cR2p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQG
# EwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0
# IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0Ex
# MB4XDTI1MDYwNDAwMDAwMFoXDTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEy
# NTYgUlNBNDA5NiBUaW1lc3RhbXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3
# zBlCMGMyqJnfFNZx+wvA69HFTBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8Tch
# TySA2R4QKpVD7dvNZh6wW2R6kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWj
# FDYOzDi8SOhPUWlLnh00Cll8pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2Uo
# yrN0ijtUDVHRXdmncOOMA3CoB/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjP
# KHW5KqCvpSduSwhwUmotuQhcg9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KS
# uNLoZLc1Hf2JNMVL4Q1OpbybpMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7w
# JNdoRORVbPR1VVnDuSeHVZlc4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vW
# doUoHLWnqWU3dCCyFG1roSrgHjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOg
# rY7rlRyTlaCCfw7aSUROwnu7zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K
# 096V1hE0yZIXe+giAwW00aHzrDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCf
# gPf8+3mnAgMBAAGjggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zy
# Me39/dfzkXFjGVBDz2GM6DAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezL
# TjAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsG
# AQUFBwEBBIGIMIGFMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wXQYIKwYBBQUHMAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNy
# dDBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5j
# cmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEB
# CwUAA4ICAQBlKq3xHCcEua5gQezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZ
# D9gBq9fNaNmFj6Eh8/YmRDfxT7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/
# ML9lFfim8/9yJmZSe2F8AQ/UdKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu
# +WUqW4daIqToXFE/JQ/EABgfZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4o
# bEMnxYOX8VBRKe1uNnzQVTeLni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2h
# ECZpqyU1d0IbX6Wq8/gVutDojBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasn
# M9AWcIQfVjnzrvwiCZ85EE8LUkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol
# /DJgddJ35XTxfUlQ+8Hggt8l2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgY
# xQbV1S3CrWqZzBt1R9xJgKf47CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3oc
# CVccAvlKV9jEnstrniLvUxxVZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcB
# ZU8atufk+EMF/cWuiC7POGT75qaL6vdCvHlshtjdNXOCIUjsarfNZzCCB1swggVD
# oAMCAQICEAixn82z2vOwMVVYCAEvAOkwDQYJKoZIhvcNAQELBQAwaTELMAkGA1UE
# BhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2Vy
# dCBUcnVzdGVkIEc0IENvZGUgU2lnbmluZyBSU0E0MDk2IFNIQTM4NCAyMDIxIENB
# MTAeFw0yMzExMTUwMDAwMDBaFw0yNjExMTcyMzU5NTlaMGMxCzAJBgNVBAYTAkdC
# MRQwEgYDVQQHEwtXaGl0bGV5IEJheTEeMBwGA1UEChMVQU5EUkVXU1RBWUxPUi5D
# T00gTFREMR4wHAYDVQQDExVBTkRSRVdTVEFZTE9SLkNPTSBMVEQwggIiMA0GCSqG
# SIb3DQEBAQUAA4ICDwAwggIKAoICAQDDpGJC6czR+GNYFFxe/fbhdAq8Fvy5nuu+
# vgvWmTiWM6at/wyGvNaFr/W+G9FsC6SVbtHzpAJsSOqHLuLn+td4wzFtBn1eHUaH
# bra8n7g7oreK53byQOgyLNGBucTZSk5GPACLwT9yMBM1A9X+eyeRogKaxxnqHOFL
# bcxLhgN8kqpbBhINIAnoVic51JId8jPF25LAtC7gZp2P3WSf/JLQsAnd/IH2RDvV
# Aw0pInuFU2N0+1RW04mh9G8PgL33EFctgksJMH55H2GoEhZCmq/jGMLu4KlV8a4d
# 1fxo72pej3TNAOxHE6ps6wkbb5FiEem6c/twCB+ha+sk7ht14iyC+rCv4hf/XeFN
# j4h9byf8X3YRo9K0N/zQUbFAQt5dcONS+avVF9TodZU9TrieoVf7mp5OiWN46Zvj
# n2e2Akxdh5M+cuofU+7GNC04uvFZrcWvxIBLRuiVTVbKj+1sBJEEcbv99KrY8qF/
# J80rhe05rEYJmdUgfiEnJXo7qkzXYXMnA4Yt/yCEoFSTUUvxemflUBn934ejm3UC
# 1cKE9CZyY2w/D8yddjqCoFYk0IZ3WmW5H6YlYnydbs1ia01ucBKx/qr7rR1beP7B
# GQFUuDzXznCV0dLNPy+SR5I7TOGhpetnwB4x4/SbApbrI+O3E+o0TiKkCOHo89bJ
# TqHJTnrQjQIDAQABo4ICAzCCAf8wHwYDVR0jBBgwFoAUaDfg67Y7+F8Rhvv+YXsI
# iGX0TkIwHQYDVR0OBBYEFNB3ThXz8WvWHm+TuSfVBIiZKyvfMD4GA1UdIAQ3MDUw
# MwYGZ4EMAQQBMCkwJwYIKwYBBQUHAgEWG2h0dHA6Ly93d3cuZGlnaWNlcnQuY29t
# L0NQUzAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwgbUGA1Ud
# HwSBrTCBqjBToFGgT4ZNaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# VHJ1c3RlZEc0Q29kZVNpZ25pbmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5jcmwwU6BR
# oE+GTWh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNENv
# ZGVTaWduaW5nUlNBNDA5NlNIQTM4NDIwMjFDQTEuY3JsMIGUBggrBgEFBQcBAQSB
# hzCBhDAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMFwGCCsG
# AQUFBzAChlBodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVz
# dGVkRzRDb2RlU2lnbmluZ1JTQTQwOTZTSEEzODQyMDIxQ0ExLmNydDAJBgNVHRME
# AjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBJEYdj8DIsq2r6+umcMOj45duCsw3BmDEs
# VCjhFOG8pHDioL0ulds3mSLNA/6KnLd4QSkxhEkkFkgPyZK9RDAtOKiRzv08K/7A
# QwgwTLSfVwLTu+SfqKg3HDPoPD6Po44amCcyr24rFVL3bD6hZZeeb0s0bxxgAOoY
# 8g8mpBi0TnDWkhWRnYDitgDIUPBFC6xPEYq4tw9UTrqhplFiqnDvWqjxwX+cFMIm
# vIPfNLE16rjuYOE1pGakr8LdQRzJruvTEepaEQCtex7xXEonCuj5tM4ndSkts+J7
# RgeAn3IPPGhS6IN3Ij9rXzItsU56jdSvmJlPPeD5dZQqdcRb0+qa5tdAbOgVYil3
# w31RcV0wdJh6AzahovwfPq8X9s+7uX6FzwswWE/kN3MbaKb2bZNvTKU2PPBhw1Sb
# BTAY1+8zOmgjSrEygto2/dhfo5ZHPHJQkEJ5d1OI6sG4+Lq9SQT+UMUKa1ocHXtP
# jInw1adLx2bFWdKZ2DAikhuW1K98DlOxD99tQ8L7xLE4wsv1w4L8vIqVOsNxo3EE
# BZoKfwV/9CQh1/tOECdCZEfU0xWUdQfs4dwnCtWRLU7b7mX40N74l8JcLTS8uN72
# 3tSeVjU7NB4yV7H3W4WOZBOUYcn2sRQa9tSmMfm/nr2MqJSEtMEejXsOhai4HHOB
# nveMmmFU6jGCBlkwggZVAgEBMH0waTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRp
# Z2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IENvZGUg
# U2lnbmluZyBSU0E0MDk2IFNIQTM4NCAyMDIxIENBMQIQCLGfzbPa87AxVVgIAS8A
# 6TANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkG
# CSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEE
# AYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC4RT/Cwjj19dfpTICTr9y49hjrXRqDRc/p
# U+qVCL6qyTANBgkqhkiG9w0BAQEFAASCAgAkI7ZZkaXNuGJNI++BPtx9NgKYPrqC
# WqJYxEFCvq8T+xIAYDf/Ibvf8SIrZgEfhWW/KaOS+tZlFAheqTsjXNRSC9lFXHVp
# 3xv5ZaORB/lv916Wxb0Sfjwi+kzxk76A1MHk/+7AMdMIyM0wsc8nn7JdyDhIof+k
# C5NEyIN2oQP+xaRTUAAcz78cFiliuMCEq1Zy5LHLrCplAmevuMtV2L5lKmpnZDq1
# 3CmEd5vyI8kaFYM/o0ph/9mBCAIRpym0pYWHI/iQW0Th7yLBJMbGjcJzrpdV+KTw
# DHJ8igykedXlQZ/yEVR64bE0Cohx27qV73dwKryLeMQrtsbvE3ZBa+liqeNqvgph
# E0LqQ886IPPEr7rFi8H0ZP8eevRUWtm+S609eCQOq5m1FeHcD8y2UxrmMebIPRT1
# taqMZ8B+rEtOHfxgaV0KB20oFNj5OvbVk47tKJ1e7AFrTvNjS+8KfOs3tJ+hR+W5
# kkem+XzvPnucYDbWdP4ppTp94Jt5nZ1rTxQSFLHxqnExRTXF3v6T6moj15VJ8x5f
# 1LtfeL1BHtbLRinjg6GqTdBkLYpJ5HhSSWPXoEMKGfdMgBhi4p00CFTRsg6579eW
# Nqu3p58LbHAdMTRM928sAOPKR1+NvUeqWC7LbPqVPlxTOI0p+aYa+OVzan9bdcV8
# 24HFAfKYfKdmYKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGln
# aUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAy
# NSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG
# 9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNTA4MDcxMzE4NTZa
# MC8GCSqGSIb3DQEJBDEiBCCNg4hQ+qdguRh7bhmiMaECjGdbZGuqLSxqdah9f5p2
# QzANBgkqhkiG9w0BAQEFAASCAgBZLfvHuDsF9OVQoPqxial2quFzFgjnANQ7JR3t
# lAtQeIo4dzVeyysYJ8Kak3eSZd2b1JLF9gK8aSnGRH04SAHY8rYgn0nFB0+Yv8Bi
# NIuxiR+bmoWZOBqSTrckmIiZEY6xGrvU1CTN+Ec6+kO0UfdYQOxlkDWatRMgza3n
# gD0lcAXSUSgpvwhEs9Us/EOCIDHNKwUc04pSYchn4gyfOkDr8J/oHKU5yL+7yFtH
# PPLU1mGvZIZreS50SfZvgQksLldVnPPUHEVVUzlogNPEDqfb1RRvv8li/zWbMVla
# YrBaJfSzcPxSn8Yt85cdAeAGcj1+rSoJZtNtZLNebq1GOOXeBOLX/D1b2da7tdmD
# g6LHp7wwtFfJeZFPXlfVo3xID1lo8HSwztQFWYwssx6YA3mKPlUjTnU9f+jMhH3w
# 1JfJtCZsqAz5cJia7/4aYqtXvArlO1ROYt+WE1/O2ug2JK+BLMBod0oZI4NglA/l
# OCHJrcbhj2Hfu2z5qXoFWYnhDs+/iyikRWkg2moR6evs7ViPFM6bo2XHR/gKQr0P
# OzvnImvNbCJ3ccnRMkgJ4Xa08pgBTcHpXHXhLXtpPILaU0wfRBGGWd5vHhZFTElv
# YIt/azW/bENjQWzs8+c/qMzKq2SYD0VUv1OIYs+qMpDvy5Ktfua4X0D9NncZhTCK
# 6kVDOg==
# SIG # End signature block
