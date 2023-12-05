$ClientID = "fda1052c-fe80-4345-a682-b9d25177d7fc"
$TenantID="46cd9aae-0d0e-45e1-83b5-154b8efeb92f"
$ClientSecret="3rT8Q~LMCIFzk~NenW6S8iCdcwKvjBZJIQ~HXcji"

$token_Body = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}
$token_Response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $token_Body
$token_Header = @{
    "Authorization" = "Bearer $($token_Response.access_token)"
    "Content-type"  = "application/json"
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
        [Parameter(Mandatory = $true)] $hardwareIdentifier
    )

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
        Invoke-RestMethod -Method Post -Uri $uri -Headers $token_Header -Body $json -ContentType "application/json"
    }
    catch {
        Write-Error $_.Exception 
        break
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
        [Parameter(Mandatory = $false)] $serial
    )

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
        $response = Invoke-restmethod -Uri $uri -Method Get -Headers $token_Header
        if ($id) {
            $response
        }
        else {
            $devices = $response.value
    
            $devicesNextLink = $response."@odata.nextLink"
    
            while ($null -ne $devicesNextLink) {
                $devicesResponse = (Invoke-restmethod -Uri $devicesNextLink -Method Get -Headers $token_Header)
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


$session = New-CimSession

$serial = (Get-CimInstance -CimSession $session -Class Win32_BIOS).SerialNumber

$devDetail = (Get-CimInstance -CimSession $session -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'")
$hash = $devDetail.DeviceHardwareData

$ap = Add-AutopilotImportedDevice -serialNumber $serial -hardwareIdentifier $hash

$device = Get-AutopilotImportedDevice | Where-Object {$_.serialNumber -eq "$($serial)"}

Write-Host "$($device.serialNumber): $($device.state.deviceImportStatus) $($device.state.deviceErrorCode) $($device.state.deviceErrorName)"
