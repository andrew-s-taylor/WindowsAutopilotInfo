
<#PSScriptInfo

.VERSION 5.13
.GUID b45605b6-65aa-45ec-a23c-f5291f9fb519
.AUTHOR AndrewTaylor, Michael Niehaus & Steven van Beek
.COMPANYNAME
.COPYRIGHT GPL
.TAGS
.LICENSEURI https://github.com/andrew-s-taylor/public/blob/main/LICENSE
.PROJECTURI https://github.com/andrew-s-taylor/public
.ICONURI
.EXTERNALMODULEDEPENDENCIES 
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
.RELEASENOTES
Version 5.12: Fixed issue with bearer
Version 5.12: Removed all commandlets and added bearer param
Version 5.11: Added logic around Device Registration event log
Version 5.10: Additional logic for DO downloads, MSI product names
Version 5.9: Code Signed
Version 5.7: Fixed LastLoggedState for Win32Apps and Added support for new Graph Module
Version 5.6: Fixed parameter handling
Version 5.5: Added support for a zip file
Version 5.4: Added additional ESP details
Version 5.3: Added hardware and OS version details
Version 5.2: Added device registration events
Version 5.1: Bug fixes
Version 5.0: Bug fixes
Version 4.9: Bug fixes
Version 4.8: Added Delivery Optimization results (but not when using a CAB file), ensured events are displayed even when no ESP
Version 4.7: Added ESP settings, fixed bugs
Version 4.6: Fixed typo
Version 4.5: Fixed but to properly reported Win32 app status when a Win32 app is installed during user ESP
Version 4.4: Added more ODJ info
Version 4.3: Added policy tracking
Version 4.2: Bug fixes for Windows 10 2004 (event ID changes)
Version 4.1: Renamed to Get-AutopilotDiagnostics
Version 4.0: Added sidecar installation info
Version 3.9: Bug fixes
Version 3.8: Bug fixes
Version 3.7: Modified Office logic to ensure it accurately reflected what ESP thinks the status is. Added ShowPolicies option.
Version 3.2: Fixed sidecar detection logic
Version 3.1: Fixed ODJ applied output
Version 3.0: Added the ability to process logs as well
Version 2.2: Added new IME MSI guid, new -AllSessions switch
Version 2.0: Added -online parameter to look up app and policy details
Version 1.0: Original published version
.PRIVATEDATA
#>

<# 

.DESCRIPTION
This script displays diagnostics information from the current PC or a captured set of logs. This includes details about the Autopilot profile settings; policies, apps, certificate profiles, etc. being tracked via the Enrollment Status Page; and additional information.
 
This should work with Windows 10 1903 and later (earlier versions have not been validated). This script will not work on ARM64 systems due to registry redirection from the use of x86 PowerShell.exe.
 

#> 
<#
.SYNOPSIS
Displays Windows Autopilot diagnostics information from the current PC or a captured set of logs.
 

.PARAMETER Online
Look up the actual policy and app names via the Microsoft Graph API
 
.PARAMETER AllSessions
Show all ESP progress instead of just the final details.
 
.PARAMETER CABFile
Processes the information in the specified CAB file (captured by MDMDiagnosticsTool.exe -area Autopilot -cab filename.cab) instead of from the registry.
 
.PARAMETER ZIPFile
Processes the information in the specified ZIP file (captured by MDMDiagnosticsTool.exe -area Autopilot -zip filename.zip) instead of from the registry.
 
.PARAMETER ShowPolicies
Shows the policy details as recorded in the NodeCache registry keys, in the order that the policies were received by the client.
 
.EXAMPLE
.\Get-AutopilotDiagnostics.ps1
 
.EXAMPLE
.\Get-AutopilotDiagnostics.ps1 -Online
 
.EXAMPLE
.\Get-AutopilotESPStatus.ps1 -AllSessions
 
.EXAMPLE
.\Get-AutopilotDiagnostics.ps1 -CABFile C:\Autopilot.cab -Online -AllSessions
 
.EXAMPLE
.\Get-AutopilotDiagnostics.ps1 -ZIPFile C:\Autopilot.zip
 
.EXAMPLE
.\Get-AutopilotDiagnostics.ps1 -ShowPolicies
 
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $False)] [String] $CABFile = $null,
    [Parameter(Mandatory = $False)] [String] $ZIPFile = $null,
    [Parameter(Mandatory = $False)] [Switch] $Online = $false,
    [Parameter(Mandatory = $False)] [Switch] $AllSessions = $false,
    [Parameter(Mandatory = $False)] [Switch] $ShowPolicies = $false,
    [Parameter(Mandatory = $false)] [string]$Tenant,
    [Parameter(Mandatory = $false)] [string]$AppId,
    [Parameter(Mandatory = $false)] [string]$AppSecret,
    [Parameter(Mandatory = $false)] [string]$bearer
)

Begin {
    # Process log files if needed
    $script:useFile = $false
    if ($CABFile -or $ZIPFile) {

        if (-not (Test-Path "$($env:TEMP)\ESPStatus.tmp")) {
            New-Item -Path "$($env:TEMP)\ESPStatus.tmp" -ItemType "directory" | Out-Null
        }
        Remove-Item -Path "$($env:TEMP)\ESPStatus.tmp\*.*" -Force -Recurse        
        $script:useFile = $true

        # If using a CAB file, extract the needed files from it
        if ($CABFile) {
            $fileList = @("MdmDiagReport_RegistryDump.reg", "microsoft-windows-devicemanagement-enterprise-diagnostics-provider-admin.evtx",
                "microsoft-windows-user device registration-admin.evtx", "AutopilotDDSZTDFile.json", "*.csv")

            $fileList | % {
                $null = & expand.exe "$CABFile" -F:$_ "$($env:TEMP)\ESPStatus.tmp\" 
                if (-not (Test-Path "$($env:TEMP)\ESPStatus.tmp\$_")) {
                    Write-Error "Unable to extract $_ from $CABFile"
                }
            }
        }
        else {
            # If using a ZIP file, just extract the entire contents (not as easy to do selected files)
            Expand-Archive -Path $ZIPFile -DestinationPath "$($env:TEMP)\ESPStatus.tmp\"
        }

        # Get the hardware hash information
        $csvFile = (Get-ChildItem "$($env:TEMP)\ESPStatus.tmp\*.csv").FullName
        if ($csvFile) {
            $csv = Get-Content $csvFile | ConvertFrom-Csv
            $hash = $csv.'Hardware Hash'
        }

        # Edit the path in the .reg file
        $content = Get-Content -Path "$($env:TEMP)\ESPStatus.tmp\MdmDiagReport_RegistryDump.reg"
        $content = $content -replace "\[HKEY_CURRENT_USER\\", "[HKEY_CURRENT_USER\ESPStatus.tmp\USER\"
        $content = $content -replace "\[HKEY_LOCAL_MACHINE\\", "[HKEY_CURRENT_USER\ESPStatus.tmp\MACHINE\"
        $content = $content -replace '^ "', '"'
        $content = $content -replace '^ @', '@'
        $content = $content -replace 'DWORD:', 'dword:'
        "Windows Registry Editor Version 5.00`n" | Set-Content -Path "$($env:TEMP)\ESPStatus.tmp\MdmDiagReport_Edited.reg"
        $content | Add-Content -Path "$($env:TEMP)\ESPStatus.tmp\MdmDiagReport_Edited.reg"

        # Remove the registry info if it exists
        if (Test-Path "HKCU:\ESPStatus.tmp") {
            Remove-Item -Path "HKCU:\ESPStatus.tmp" -Recurse -Force
        }

        # Import the .reg file
        $null = & reg.exe IMPORT "$($env:TEMP)\ESPStatus.tmp\MdmDiagReport_Edited.reg" 2>&1

        # Configure the (not live) constants
        $script:provisioningPath = "HKCU:\ESPStatus.tmp\MACHINE\software\microsoft\provisioning"
        $script:autopilotDiagPath = "HKCU:\ESPStatus.tmp\MACHINE\software\microsoft\provisioning\Diagnostics\Autopilot"
        $script:omadmPath = "HKCU:\ESPStatus.tmp\MACHINE\software\microsoft\provisioning\OMADM"
        $script:path = "HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics"
        $script:msiPath = "HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\EnterpriseDesktopAppManagement"
        $script:officePath = "HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\OfficeCSP"
        $script:sidecarPath = "HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\IntuneManagementExtension\Win32Apps"
        $script:enrollmentsPath = "HKCU:\ESPStatus.tmp\MACHINE\software\microsoft\enrollments"
    }
    else {
        # Configure live constants
        $script:provisioningPath = "HKLM:\software\microsoft\provisioning"
        $script:autopilotDiagPath = "HKLM:\software\microsoft\provisioning\Diagnostics\Autopilot"
        $script:omadmPath = "HKLM:\software\microsoft\provisioning\OMADM"
        $script:path = "HKLM:\Software\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics"
        $script:msiPath = "HKLM:\Software\Microsoft\EnterpriseDesktopAppManagement"
        $script:officePath = "HKLM:\Software\Microsoft\OfficeCSP"
        $script:sidecarPath = "HKLM:\Software\Microsoft\IntuneManagementExtension\Win32Apps"
        $script:enrollmentsPath = "HKLM:\Software\Microsoft\enrollments"

        $hash = (Get-WmiObject -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'").DeviceHardwareData
    }

    # Configure other constants
    $script:officeStatus = @{"0" = "None"; "10" = "Initialized"; "20" = "Download In Progress"; "25" = "Pending Download Retry";
        "30" = "Download Failed"; "40" = "Download Completed"; "48" = "Pending User Session"; "50" = "Enforcement In Progress"; 
        "55" = "Pending Enforcement Retry"; "60" = "Enforcement Failed"; "70" = "Success / Enforcement Completed"
    }
    $script:espStatus = @{"1" = "Not Installed"; "2" = "Downloading / Installing"; "3" = "Success / Installed"; "4" = "Error / Failed" }
    $script:policyStatus = @{"0" = "Not Processed"; "1" = "Processed" }

    # Configure any other global variables
    $script:observedTimeline = @()
}

Process {
    #------------------------
    # Functions
    #------------------------

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
        $response = (Invoke-MgGraphRequest -Uri $url -Method Get -OutputType PSObject)
        $alloutput = $response.value
    
        $alloutputNextLink = $response."@odata.nextLink"
    
        while ($null -ne $alloutputNextLink) {
            $alloutputResponse = (Invoke-MgGraphRequest -Uri $alloutputNextLink -Method Get -OutputType PSObject)
            $alloutputNextLink = $alloutputResponse."@odata.nextLink"
            $alloutput += $alloutputResponse.value
        }
    
        return $alloutput
    }
    
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
        [Parameter(Mandatory = $false)] [string]$scopes,
        [Parameter(Mandatory = $false)] [string]$bearer
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
        elseif ($bearer -ne "") {
            if ($version -eq 2) {
                write-host "Version 2 module detected"
                $accesstokenfinal = ConvertTo-SecureString -String $bearer -AsPlainText -Force
            }
            else {
                write-host "Version 1 Module Detected"
                Select-MgProfile -Name Beta
                $accesstokenfinal = $bearer
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

    Function RecordStatus() {
        param
        (
            [Parameter(Mandatory = $true)] [String] $detail,
            [Parameter(Mandatory = $true)] [String] $status,
            [Parameter(Mandatory = $true)] [String] $color,
            [Parameter(Mandatory = $true)] [datetime] $date
        )

        # See if there is already an entry for this policy and status
        $found = $script:observedTimeline | ? { $_.Detail -eq $detail -and $_.Status -eq $status }
        if (-not $found) {
            # Apply a fudge so that the downloading of the next app appears one second after the previous completion
            if ($status -like "Downloading*") {
                $adjustedDate = $date.AddSeconds(1)
            } else {
                $adjustedDate = $date
            }
            $script:observedTimeline += New-Object PSObject -Property @{
                "Date"   = $adjustedDate
                "Detail" = $detail
                "Status" = $status
                "Color"  = $color
            }
        }
    }

    Function AddDisplay() {
        param
        (
            [Parameter(Mandatory = $true)] [ref]$items
        )
        $items.Value | % {
            Add-Member -InputObject $_ -NotePropertyName display -NotePropertyValue $AllSessions
        }
        $items.Value[$items.Value.Count - 1].display = $true
    }
    
    Function ProcessApps() {
        param
        (
            [Parameter(Mandatory = $true, ValueFromPipeline = $True)] [Microsoft.Win32.RegistryKey] $currentKey,
            [Parameter(Mandatory = $true)] $currentUser,
            [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] [bool] $display
        )

        Begin {
            if ($display) { Write-Host "Apps:" }
        }

        Process {
            if ($display) { Write-Host " $(([datetime]$currentKey.PSChildName).ToString('u'))" }
            $currentKey.Property | % {
                if ($_.StartsWith("./Device/Vendor/MSFT/EnterpriseDesktopAppManagement/MSI/")) {
                    $msiKey = [URI]::UnescapeDataString(($_.Split("/"))[6])
                    $fullPath = "$msiPath\$currentUser\MSI\$msiKey"
                    if (Test-Path $fullPath) {
                        $status = (Get-ItemProperty -Path $fullPath).Status
                        $msiFile = (Get-ItemProperty -Path $fullPath).CurrentDownloadUrl
                    }
                    if ($status -eq "" -or $status -eq $null) {
                        $status = 0
                    } 
                    if ($msiFile -match "IntuneWindowsAgent.msi") {
                        $msiKey = "Intune Management Extensions ($($msiKey))"
                    }
                    elseif ($Online) {
                        $found = $apps | ? { $_.ProductCode -contains $msiKey }
                        $msiKey = "$($found.DisplayName) ($($msiKey))"
                    } elseif ($currentUser -eq "S-0-0-00-0000000000-0000000000-000000000-000") {
                        # Try to read the name from the uninstall registry key
                        if (Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$msiKey") {
                            $displayName = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$msiKey").DisplayName
                            $msiKey = "$displayName ($($msiKey))"
                        }
                    }
                    if ($status -eq 70) {
                        if ($display) { Write-Host " MSI $msiKey : $status ($($officeStatus[$status.ToString()]))" -ForegroundColor Green }
                        RecordStatus -detail "MSI $msiKey" -status $officeStatus[$status.ToString()] -color "Green" -date $currentKey.PSChildName
                    }
                    elseif ($status -eq 60) {
                        if ($display) { Write-Host " MSI $msiKey : $status ($($officeStatus[$status.ToString()]))" -ForegroundColor Red }
                        RecordStatus -detail "MSI $msiKey" -status $officeStatus[$status.ToString()] -color "Red" -date $currentKey.PSChildName
                    }
                    else {
                        if ($display) { Write-Host " MSI $msiKey : $status ($($officeStatus[$status.ToString()]))" -ForegroundColor Yellow }
                        RecordStatus -detail "MSI $msiKey" -status $officeStatus[$status.ToString()] -color "Yellow" -date $currentKey.PSChildName
                    }
                }
                elseif ($_.StartsWith("./Vendor/MSFT/Office/Installation/")) {
                    # Report the main status based on what ESP is tracking
                    $status = Get-ItemPropertyValue -Path $currentKey.PSPath -Name $_

                    # Then try to get the detailed Office status
                    $officeKey = [URI]::UnescapeDataString(($_.Split("/"))[5])
                    $fullPath = "$officepath\$officeKey"
                    if (Test-Path $fullPath) {
                        $oStatus = (Get-ItemProperty -Path $fullPath).FinalStatus

                        if ($oStatus -eq $null) {
                            $oStatus = (Get-ItemProperty -Path $fullPath).Status
                            if ($oStatus -eq $null) {
                                $oStatus = "None"
                            }
                        }
                    }
                    else {
                        $oStatus = "None"
                    }
                    if ($officeStatus.Keys -contains $oStatus.ToString()) {
                        $officeStatusText = $officeStatus[$oStatus.ToString()]
                    }
                    else {
                        $officeStatusText = $oStatus
                    }
                    if ($status -eq 1) {
                        if ($display) { Write-Host " Office $officeKey : $status ($($policyStatus[$status.ToString()]) / $officeStatusText)" -ForegroundColor Green }
                        RecordStatus -detail "Office $officeKey" -status "$($policyStatus[$status.ToString()]) / $officeStatusText" -color "Green" -date $currentKey.PSChildName
                    }
                    else {
                        if ($display) { Write-Host " Office $officeKey : $status ($($policyStatus[$status.ToString()]) / $officeStatusText)" -ForegroundColor Yellow }
                        RecordStatus -detail "Office $officeKey" -status "$($policyStatus[$status.ToString()]) / $officeStatusText" -color "Yellow" -date $currentKey.PSChildName
                    }
                }
                else {
                    if ($display) { Write-Host " $_ : Unknown app" }
                }
            }
        }

    }

    Function ProcessModernApps() {
        param
        (
            [Parameter(Mandatory = $true, ValueFromPipeline = $True)] [Microsoft.Win32.RegistryKey] $currentKey,
            [Parameter(Mandatory = $true)] $currentUser,
            [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] [bool] $display
        )

        Begin {
            if ($display) { Write-Host "Modern Apps:" }
        }

        Process {
            if ($display) { Write-Host " $(([datetime]$currentKey.PSChildName).ToString('u'))" }
            $currentKey.Property | % {
                $status = (Get-ItemPropertyValue -path $currentKey.PSPath -Name $_).ToString()
                if ($_.StartsWith("./User/Vendor/MSFT/EnterpriseModernAppManagement/AppManagement/")) {
                    $appID = [URI]::UnescapeDataString(($_.Split("/"))[7])
                    $type = "User UWP"
                }
                elseif ($_.StartsWith("./Device/Vendor/MSFT/EnterpriseModernAppManagement/AppManagement/")) {
                    $appID = [URI]::UnescapeDataString(($_.Split("/"))[7])
                    $type = "Device UWP"
                }
                else {
                    $appID = $_
                    $type = "Unknown UWP"
                }
                if ($status -eq "1") {
                    if ($display) { Write-Host " $type $appID : $status ($($policyStatus[$status]))" -ForegroundColor Green }
                    RecordStatus -detail "UWP $appID" -status $policyStatus[$status] -color "Green" -date $currentKey.PSChildName
                }
                else {
                    if ($display) { Write-Host " $type $appID : $status ($($policyStatus[$status]))" -ForegroundColor Yellow }
                }
            }
        }

    }

    Function ProcessSidecar() {
        param
        (
            [Parameter(Mandatory = $true, ValueFromPipeline = $True)] [Microsoft.Win32.RegistryKey] $currentKey,
            [Parameter(Mandatory = $true)] $currentUser,
            [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] [bool] $display
        )

        Begin {
            if ($display) { Write-Host "Sidecar apps:" }
            if ($null -eq $script:DOEvents -and (-not $script:useFile)) {
                $script:DOEvents = Get-DeliveryOptimizationLog | Where-Object { $_.Function -match "(DownloadStart)|(DownloadCompleted)" -and $_.Message -like "*.intunewin*" }
            }
        }

        Process {
            if ($display) { Write-Host " $(([datetime]$currentKey.PSChildName).ToString('u'))" }
            $currentKey.Property | % {
                $win32Key = [URI]::UnescapeDataString(($_.Split("/"))[9])
                $status = Get-ItemPropertyValue -path $currentKey.PSPath -Name $_
                if ($Online) {
                    $found = $apps | ? { $win32Key -match $_.Id }
                    $win32Key = "$($found.DisplayName) ($($win32Key))"
                }
                $appGuid = $win32Key.Substring(9)
                $sidecarApp = "$sidecarPath\$currentUser\$appGuid"
                $exitCode = $null
                if (Test-Path $sidecarApp) {
                    $exitCode = (Get-ItemProperty -Path $sidecarApp).ExitCode
                }
                if ($status -eq "3") {
                    if ($exitCode -ne $null) {
                        if ($display) { Write-Host " Win32 $win32Key : $status ($($espStatus[$status.ToString()]), rc = $exitCode)" -ForegroundColor Green }
                    }
                    else {
                        if ($display) { Write-Host " Win32 $win32Key : $status ($($espStatus[$status.ToString()]))" -ForegroundColor Green }
                    }
                    RecordStatus -detail "Win32 $win32Key" -status $espStatus[$status.ToString()] -color "Green" -date $currentKey.PSChildName
                }
                elseif ($status -eq "4") {
                    if ($exitCode -ne $null) {
                        if ($display) { Write-Host " Win32 $win32Key : $status ($($espStatus[$status.ToString()]), rc = $exitCode)" -ForegroundColor Red }
                    }
                    else {
                        if ($display) { Write-Host " Win32 $win32Key : $status ($($espStatus[$status.ToString()]))" -ForegroundColor Red }
                    }
                    RecordStatus -detail "Win32 $win32Key" -status $espStatus[$status.ToString()] -color "Red" -date $currentKey.PSChildName
                }
                else {
                    if ($exitCode -ne $null) {
                        if ($display) { Write-Host " Win32 $win32Key : $status ($($espStatus[$status.ToString()]), rc = $exitCode)" -ForegroundColor Yellow }
                    }
                    else {
                        if ($display) { Write-Host " Win32 $win32Key : $status ($($espStatus[$status.ToString()]))" -ForegroundColor Yellow }
                    }
                    if ($status -ne "1") {
                        RecordStatus -detail "Win32 $win32Key" -status $espStatus[$status.ToString()] -color "Yellow" -date $currentKey.PSChildName
                    }
                    if ($status -eq "2") {
                        # Try to find the DO events.
                        $script:DOEvents | Where-Object { $_.Message -ilike "*$appGuid*" } | ForEach-Object {
                            RecordStatus -detail "Win32 $win32Key" -status "DO $($_.Function.Substring(32))" -color "Yellow" -date $_.TimeCreated.ToLocalTime()
                        }
                    }
                }
            }
        }

    }

    Function ProcessPolicies() {
        param
        (
            [Parameter(Mandatory = $true, ValueFromPipeline = $True)] [Microsoft.Win32.RegistryKey] $currentKey,
            [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] [bool] $display
        )

        Begin {
            if ($display) { Write-Host "Policies:" }
        }

        Process {
            if ($display) { Write-Host " $(([datetime]$currentKey.PSChildName).ToString('u'))" }
            $currentKey.Property | % {
                $status = Get-ItemPropertyValue -path $currentKey.PSPath -Name $_
                if ($status -eq "1") {
                    if ($display) { Write-Host " Policy $_ : $status ($($policyStatus[$status.ToString()]))" -ForegroundColor Green }
                    RecordStatus -detail "Policy $_" -status $policyStatus[$status.ToString()] -color "Green" -date $currentKey.PSChildName
                }
                else {
                    if ($display) { Write-Host " Policy $_ : $status ($($policyStatus[$status.ToString()]))" -ForegroundColor Yellow }
                }
            }
        }

    }

    Function ProcessCerts() {
        param
        (
            [Parameter(Mandatory = $true, ValueFromPipeline = $True)] [Microsoft.Win32.RegistryKey] $currentKey,
            [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $True)] [bool] $display
        )

        Begin {
            if ($display) { Write-Host "Certificates:" }
        }

        Process {
            if ($display) { Write-Host " $(([datetime]$currentKey.PSChildName).ToString('u'))" }
            $currentKey.Property | % {
                $certKey = [URI]::UnescapeDataString(($_.Split("/"))[6])
                $status = Get-ItemPropertyValue -path $currentKey.PSPath -Name $_
                if ($Online) {
                    $found = $policies | ? { $certKey.Replace("_", "-") -match $_.Id }
                    $certKey = "$($found.DisplayName) ($($certKey))"
                }
                if ($status -eq "1") {
                    if ($display) { Write-Host " Cert $certKey : $status ($($policyStatus[$status.ToString()]))" -ForegroundColor Green }
                    RecordStatus -detail "Cert $certKey" -status $policyStatus[$status.ToString()] -color "Green" -date $currentKey.PSChildName
                }
                else {
                    if ($display) { Write-Host " Cert $certKey : $status ($($policyStatus[$status.ToString()]))" -ForegroundColor Yellow }
                }
            }
        }

    }

    Function ProcessNodeCache() {

        Process {
            $nodeCount = 0
            while ($true) {
                # Get the nodes in order. This won't work after a while because the older numbers are deleted as new ones are added
                # but it will work out OK shortly after provisioning. The alternative would be to get all the subkeys and then sort
                # them numerically instead of alphabetically, but that can be saved for later...
                $node = Get-ItemProperty "$provisioningPath\NodeCache\CSP\Device\MS DM Server\Nodes\$nodeCount" -ErrorAction SilentlyContinue
                if ($node -eq $null) {
                    break
                }
                $nodeCount += 1
                $node | Select NodeUri, ExpectedValue
            }
        }

    }

    Function TrimMSI() {
	param (
		[object] $event,
		[string] $sidecarProductCode
	)

    # Fix up the name
	if ($event.Properties[0].Value -eq $sidecarProductCode) {
		return "Intune Management Extension"
	} elseif ($event.Properties[0].Value.StartsWith("{{")) {
		$r = $event.Properties[0].Value.Substring(1, $event.Properties[0].Value.Length - 2)
	} else {
		$r = $event.Properties[0].Value
	}

	# See if we can find the real name
    if (Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$r") {
		$displayName = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$r").DisplayName
        return "$displayName ($($r))"
    } else {
    	return $r
    }

    }

    Function ProcessEvents() {

        Process {

            $productCode = 'IME-Not-Yet-Installed'
            if (Test-Path "$msiPath\S-0-0-00-0000000000-0000000000-000000000-000\MSI") {
                Get-ChildItem -path "$msiPath\S-0-0-00-0000000000-0000000000-000000000-000\MSI" | % {
                    $file = (Get-ItemProperty -Path $_.PSPath).CurrentDownloadUrl
                    if ($file -match "IntuneWindowsAgent.msi") {
                        $productCode = Get-ItemPropertyValue -Path $_.PSPath -Name ProductCode
                    }
                }
            }

            # Process device management events
            if ($script:useFile) {
                $events = Get-WinEvent -Path "$($env:TEMP)\ESPStatus.tmp\microsoft-windows-devicemanagement-enterprise-diagnostics-provider-admin.evtx" -Oldest | ? { ($_.Id -in 1905, 1906, 1920, 1922) -or $_.Id -in (72, 100, 107, 109, 110, 111) }
            }
            else {
                $events = Get-WinEvent -LogName Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin -Oldest | ? { ($_.Id -in 1905, 1906, 1920, 1922) -or $_.Id -in (72, 100, 107, 109, 110, 111) }
            }
            $events | % {
                $message = $_.Message
                $detail = "Sidecar"
                $color = "Yellow"
                $event = $_
                switch ($_.id) {
                    { $_ -in (110, 109) } { 
                        $detail = "Offline Domain Join"
                        switch ($event.Properties[0].Value) {
                            0 { $message = "Offline domain join not configured" }
                            1 { $message = "Waiting for ODJ blob" }
                            2 { $message = "Processed ODJ blob" }
                            3 { $message = "Timed out waiting for ODJ blob or connectivity" }
                        }
                    }
                    111 { $detail = "Offline Domain Join"; $message = "Starting wait for ODJ blob" }
                    107 { $detail = "Offline Domain Join"; $message = "Successfully applied ODJ blob" }
                    100 { $detail = "Offline Domain Join"; $message = "Could not establish connectivity"; $color = "Red" }
                    72 { $detail = "MDM Enrollment" }
                    1905 { $detail = (TrimMSI $event $productCode); $message = "Download started" }
                    1906 { $detail = (TrimMSI $event $productCode); $message = "Download finished" }
                    1920 { $detail = (TrimMSI $event $productCode); $message = "Installation started" }
                    1922 { $detail = (TrimMSI $event $productCode); $message = "Installation finished" }
                    { $_ -in (1922, 72) } { $color = "Green" }
                }
                RecordStatus -detail $detail -date $_.TimeCreated -status $message -color $color
            }

            # Process device registration events
            if ($script:useFile) {
                $events = Get-WinEvent -Path "$($env:TEMP)\ESPStatus.tmp\microsoft-windows-user device registration-admin.evtx" -Oldest | ? { $_.Id -in (306, 101) }
            }
            else {
                try {
                    $events = Get-WinEvent -LogName 'Microsoft-Windows-User Device Registration/Admin' -Oldest -ErrorAction Stop | ? { $_.Id -in (306, 101) }
                } catch [Exception] {
                    if ($_.FullyQualifiedErrorId -match "NoMatchingEventsFound") {
                     $events = @()
                    }
                }
            }
            $events | % {
                $message = $_.Message
                $detail = "Device Registration"
                $color = "Yellow"
                $event = $_
                switch ($_.id) {
                    101 { $detail = "Device Registration"; $message = "SCP discovery successful." }
                    304 { $detail = "Device Registration"; $message = "Hybrid AADJ device registration failed." }
                    306 { $detail = "Device Registration"; $message = "Hybrid AADJ device registration succeeded."; $color = 'Green' }
                }
                RecordStatus -detail $detail -date $_.TimeCreated -status $message -color $color
            }

        }
    
    }
    
    #------------------------
    # Main code
    #------------------------

    # If online, make sure we are able to authenticate
    if ($Online) {

        Write-Host "Connect to Graph!"
        #Connect to Graph
        if ($AppId -and $AppSecret -and $tenant) {

            $graph = Connect-ToGraph -Tenant $tenant -AppId $clientid -AppSecret $clientsecret
            write-output "Graph Connection Established"
            }
            elseif ($bearer) {
                $graph = Connect-ToGraph -bearer $bearer
                write-output "Graph Connection Established"
    
            }
            else {
            ##Connect to Graph
            
            $graph = Connect-ToGraph -Scopes "DeviceManagementApps.Read.All, DeviceManagementConfiguration.Read.All"
            }
        Write-Host "Connected to tenant $($graph.TenantId)"

        # Get a list of apps
        Write-Host "Getting list of apps"
        #$script:apps = Get-MgDeviceAppManagementMobileApp -All
        $appsuri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps"
        $script:apps = getallpagination -url $appsuri
        

        # Get a list of policies (for certs)
        Write-Host "Getting list of policies"
        $configuri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
        #$script:policies = Get-MgBetaDeviceManagementConfigurationPolicy -All
        $script:policies = getallpagination -url $configuri

    }

    # Display Autopilot diag details
    Write-Host ""
    Write-Host "AUTOPILOT DIAGNOSTICS" -ForegroundColor Magenta
    Write-Host ""

    $values = Get-ItemProperty "$autopilotDiagPath"
    if (-not $values.CloudAssignedTenantId) {
        Write-Host "This is not an Autopilot device.`n"
        exit 0
    }

    if (-not $script:useFile) {
        $osVersion = (Get-WmiObject win32_operatingsystem).Version
        Write-Host "OS version: $osVersion"
    }
    Write-Host "Profile: $($values.DeploymentProfileName)"
    Write-Host "TenantDomain: $($values.CloudAssignedTenantDomain)"
    Write-Host "TenantID: $($values.CloudAssignedTenantId)"
    $correlations = Get-ItemProperty "$autopilotDiagPath\EstablishedCorrelations"
    Write-Host "ZTDID: $($correlations.ZTDRegistrationID)"
    Write-Host "EntDMID: $($correlations.EntDMID)"

    Write-Host "OobeConfig: $($values.CloudAssignedOobeConfig)"

    if (($values.CloudAssignedOobeConfig -band 1024) -gt 0) {
        Write-Host " Skip keyboard: Yes 1 - - - - - - - - - -"
    }
    else {
        Write-Host " Skip keyboard: No 0 - - - - - - - - - -"
    }
    if (($values.CloudAssignedOobeConfig -band 512) -gt 0) {
        Write-Host " Enable patch download: Yes - 1 - - - - - - - - -"
    }
    else {
        Write-Host " Enable patch download: No - 0 - - - - - - - - -"
    }
    if (($values.CloudAssignedOobeConfig -band 256) -gt 0) {
        Write-Host " Skip Windows upgrade UX: Yes - - 1 - - - - - - - -"
    }
    else {
        Write-Host " Skip Windows upgrade UX: No - - 0 - - - - - - - -"
    }
    if (($values.CloudAssignedOobeConfig -band 128) -gt 0) {
        Write-Host " AAD TPM Required: Yes - - - 1 - - - - - - -"
    }
    else {
        Write-Host " AAD TPM Required: No - - - 0 - - - - - - -"
    }
    if (($values.CloudAssignedOobeConfig -band 64) -gt 0) {
        Write-Host " AAD device auth: Yes - - - - 1 - - - - - -"
    }
    else {
        Write-Host " AAD device auth: No - - - - 0 - - - - - -"
    }
    if (($values.CloudAssignedOobeConfig -band 32) -gt 0) {
        Write-Host " TPM attestation: Yes - - - - - 1 - - - - -"
    }
    else {
        Write-Host " TPM attestation: No - - - - - 0 - - - - -"
    }
    if (($values.CloudAssignedOobeConfig -band 16) -gt 0) {
        Write-Host " Skip EULA: Yes - - - - - - 1 - - - -"
    }
    else {
        Write-Host " Skip EULA: No - - - - - - 0 - - - -"
    }
    if (($values.CloudAssignedOobeConfig -band 8) -gt 0) {
        Write-Host " Skip OEM registration: Yes - - - - - - - 1 - - -"
    }
    else {
        Write-Host " Skip OEM registration: No - - - - - - - 0 - - -"
    }
    if (($values.CloudAssignedOobeConfig -band 4) -gt 0) {
        Write-Host " Skip express settings: Yes - - - - - - - - 1 - -"
    }
    else {
        Write-Host " Skip express settings: No - - - - - - - - 0 - -"
    }
    if (($values.CloudAssignedOobeConfig -band 2) -gt 0) {
        Write-Host " Disallow admin: Yes - - - - - - - - - 1 -"
    }
    else {
        Write-Host " Disallow admin: No - - - - - - - - - 0 -"
    }

    # In theory we could read these values from the profile cache registry key, but it's so bungled
    # up in the registry export that it doesn't import without some serious massaging for embedded
    # quotes. So this is easier.
    if ($script:useFile) {
        $jsonFile = "$($env:TEMP)\ESPStatus.tmp\AutopilotDDSZTDFile.json"
    }
    else {
        $jsonFile = "$($env:WINDIR)\ServiceState\wmansvc\AutopilotDDSZTDFile.json" 
    }
    if (Test-Path $jsonFile) {
        $json = Get-Content $jsonFile | ConvertFrom-Json
        $date = [datetime]$json.PolicyDownloadDate
        RecordStatus -date $date -detail "Autopilot profile" -status "Profile downloaded" -color "Yellow" 
        if ($json.CloudAssignedDomainJoinMethod -eq 1) {
            Write-Host "Scenario: Hybrid Azure AD Join"
            if (Test-Path "$omadmPath\SyncML\ODJApplied") {
                Write-Host "ODJ applied: Yes"
            }
            else {
                Write-Host "ODJ applied: No"                
            }
            if ($json.HybridJoinSkipDCConnectivityCheck -eq 1) {
                Write-Host "Skip connectivity check: Yes"
            }
            else {
                Write-Host "Skip connectivity check: No"
            }

        }
        else {
            Write-Host "Scenario: Azure AD Join"
        }
    }
    else {
        Write-Host "Scenario: Not available (JSON not found)"
    }

    # Get ESP properties
    Get-ChildItem $enrollmentsPath | ? { Test-Path "$($_.PSPath)\FirstSync" } | % {
        $properties = Get-ItemProperty "$($_.PSPath)\FirstSync"
        Write-Host "Enrollment status page:"
        Write-Host " Device ESP enabled: $($properties.SkipDeviceStatusPage -eq 0)"
        Write-Host " User ESP enabled: $($properties.SkipUserStatusPage -eq 0)"
        Write-Host " ESP timeout: $($properties.SyncFailureTimeout)"
        if ($properties.BlockInStatusPage -eq 0) {
            Write-Host " ESP blocking: No"
        }
        else {
            Write-Host " ESP blocking: Yes"
            if ($properties.BlockInStatusPage -band 1) {
                Write-Host " ESP allow reset: Yes"
            }
            if ($properties.BlockInStatusPage -band 2) {
                Write-Host " ESP allow try again: Yes"
            }
            if ($properties.BlockInStatusPage -band 4) {
                Write-Host " ESP continue anyway: Yes"
            }
        }
    }

    # Get Delivery Optimization statistics (when available)
    if (-not $script:useFile) {
        $stats = Get-DeliveryOptimizationPerfSnapThisMonth
        if ($stats.DownloadHttpBytes -ne 0) {
            $peerPct = [math]::Round( ($stats.DownloadLanBytes / $stats.DownloadHttpBytes) * 100 )
            $ccPct = [math]::Round( ($stats.DownloadCacheHostBytes / $stats.DownloadHttpBytes) * 100 )
        }
        else {
            $peerPct = 0
            $ccPct = 0
        }
        Write-Host "Delivery Optimization statistics:"
        Write-Host " Total bytes downloaded: $($stats.DownloadHttpBytes)"
        Write-Host " From peers: $($peerPct)% ($($stats.DownloadLanBytes))"
        Write-host " From Connected Cache: $($ccPct)% ($($stats.DownloadCacheHostBytes))"
    }

    # If the ADK is installed, get some key hardware hash info
    $adkPath = Get-ItemPropertyValue "HKLM:\Software\Microsoft\Windows Kits\Installed Roots" -Name KitsRoot10 -ErrorAction SilentlyContinue
    $oa3Tool = "$adkPath\Assessment and Deployment Kit\Deployment Tools\$($env:PROCESSOR_ARCHITECTURE)\Licensing\OA30\oa3tool.exe"
    if ($hash -and (Test-Path $oa3Tool)) {
        $commandLineArgs = "/decodehwhash:$hash"
        $output = & "$oa3Tool" $commandLineArgs
        [xml] $hashXML = $output | Select -skip 8 -First ($output.Count - 12)
        Write-Host "Hardware information:"
        Write-Host " Operating system build: " $hashXML.SelectSingleNode("//p[@n='OsBuild']").v
        Write-Host " Manufacturer: " $hashXML.SelectSingleNode("//p[@n='SmbiosSystemManufacturer']").v
        Write-Host " Model: " $hashXML.SelectSingleNode("//p[@n='SmbiosSystemProductName']").v
        Write-Host " Serial number: " $hashXML.SelectSingleNode("//p[@n='SmbiosSystemSerialNumber']").v
        Write-Host " TPM version: " $hashXML.SelectSingleNode("//p[@n='TPMVersion']").v
    }
    
    # Process event log info
    ProcessEvents

    # Display the list of policies
    if ($ShowPolicies) {
        Write-Host " "
        Write-Host "POLICIES PROCESSED" -ForegroundColor Magenta   
        ProcessNodeCache | Format-Table -Wrap
    }
    
    # Make sure the tracking path exists
    if (Test-Path $path) {

        # Process device ESP sessions
        Write-Host " "
        Write-Host "DEVICE ESP:" -ForegroundColor Magenta
        Write-Host " "

        if (Test-Path "$path\ExpectedPolicies") {
            [array]$items = Get-ChildItem "$path\ExpectedPolicies"
            AddDisplay ([ref]$items)
            $items | ProcessPolicies
        }
        if (Test-Path "$path\ExpectedMSIAppPackages") {
            [array]$items = Get-ChildItem "$path\ExpectedMSIAppPackages"
            AddDisplay ([ref]$items)
            $items | ProcessApps -currentUser "S-0-0-00-0000000000-0000000000-000000000-000" 
        }
        if (Test-Path "$path\ExpectedModernAppPackages") {
            [array]$items = Get-ChildItem "$path\ExpectedModernAppPackages"
            AddDisplay ([ref]$items)
            $items | ProcessModernApps -currentUser "S-0-0-00-0000000000-0000000000-000000000-000"
        }
        if (Test-Path "$path\Sidecar") {
            [array]$items = Get-ChildItem "$path\Sidecar" | ? { $_.Property -match "./Device" -and $_.Name -notmatch "LastLoggedState" }
            AddDisplay ([ref]$items)
            $items | ProcessSidecar -currentUser "00000000-0000-0000-0000-000000000000"
        }
        if (Test-Path "$path\ExpectedSCEPCerts") {
            [array]$items = Get-ChildItem "$path\ExpectedSCEPCerts"
            AddDisplay ([ref]$items)
            $items | ProcessCerts
        }

        # Process user ESP sessions
        Get-ChildItem "$path" | ? { $_.PSChildName.StartsWith("S-") } | % {
            $userPath = $_.PSPath
            $userSid = $_.PSChildName
            Write-Host " "
            Write-Host "USER ESP for $($userSid):" -ForegroundColor Magenta
            Write-Host " "
            if (Test-Path "$userPath\ExpectedPolicies") {
                [array]$items = Get-ChildItem "$userPath\ExpectedPolicies"
                AddDisplay ([ref]$items)
                $items | ProcessPolicies
            }
            if (Test-Path "$userPath\ExpectedMSIAppPackages") {
                [array]$items = Get-ChildItem "$userPath\ExpectedMSIAppPackages" 
                AddDisplay ([ref]$items)
                $items | ProcessApps -currentUser $userSid
            }
            if (Test-Path "$userPath\ExpectedModernAppPackages") {
                [array]$items = Get-ChildItem "$userPath\ExpectedModernAppPackages"
                AddDisplay ([ref]$items)
                $items | ProcessModernApps -currentUser $userSid
            }
            if (Test-Path "$userPath\Sidecar") {
                [array]$items = Get-ChildItem "$path\Sidecar" | ? { $_.Property -match "./User" }
                AddDisplay ([ref]$items)
                $items | ProcessSidecar -currentUser $userSid
            }
            if (Test-Path "$userPath\ExpectedSCEPCerts") {
                [array]$items = Get-ChildItem "$userPath\ExpectedSCEPCerts"
                AddDisplay ([ref]$items)
                $items | ProcessCerts
            }
        }
    }
    else {
        Write-Host "ESP diagnostics info does not (yet) exist."
    }

    # Display timeline
    Write-Host ""
    Write-Host "OBSERVED TIMELINE:" -ForegroundColor Magenta
    Write-Host ""
    $observedTimeline | Sort-Object -Property Date |
    Format-Table @{
        Label      = "Date"
        Expression = { $_.Date.ToString("u") } 
    }, 
    @{
        Label      = "Status"
        Expression =
        {
            switch ($_.Color) {
                'Red' { $color = "91"; break }
                'Yellow' { $color = '93'; break }
                'Green' { $color = "92"; break }
                default { $color = "0" }
            }
            $e = [char]27
            "$e[${color}m$($_.Status)$e[0m"
        }
    },
    Detail

    Write-Host ""
}

End {

    # Remove the registry info if it exists
    if (Test-Path "HKCU:\ESPStatus.tmp") {
        Remove-Item -Path "HKCU:\ESPStatus.tmp" -Recurse -Force
    }
}



# SIG # Begin signature block
# MIIoEwYJKoZIhvcNAQcCoIIoBDCCKAACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDhlebkr9G7KK9e
# /UEvJeeHgOoaIXeR18f+eT/srl9h7qCCIRYwggWNMIIEdaADAgECAhAOmxiO+dAt
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
# Vzu0nAPthkX0tGFuv2jiJmCG6sivqf6UHedjGzqGVnhOMIIGvDCCBKSgAwIBAgIQ
# C65mvFq6f5WHxvnpBOMzBDANBgkqhkiG9w0BAQsFADBjMQswCQYDVQQGEwJVUzEX
# MBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0
# ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBMB4XDTI0MDkyNjAw
# MDAwMFoXDTM1MTEyNTIzNTk1OVowQjELMAkGA1UEBhMCVVMxETAPBgNVBAoTCERp
# Z2lDZXJ0MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyNDCCAiIwDQYJ
# KoZIhvcNAQEBBQADggIPADCCAgoCggIBAL5qc5/2lSGrljC6W23mWaO16P2RHxjE
# iDtqmeOlwf0KMCBDEr4IxHRGd7+L660x5XltSVhhK64zi9CeC9B6lUdXM0s71EOc
# Re8+CEJp+3R2O8oo76EO7o5tLuslxdr9Qq82aKcpA9O//X6QE+AcaU/byaCagLD/
# GLoUb35SfWHh43rOH3bpLEx7pZ7avVnpUVmPvkxT8c2a2yC0WMp8hMu60tZR0Cha
# V76Nhnj37DEYTX9ReNZ8hIOYe4jl7/r419CvEYVIrH6sN00yx49boUuumF9i2T8U
# uKGn9966fR5X6kgXj3o5WHhHVO+NBikDO0mlUh902wS/Eeh8F/UFaRp1z5SnROHw
# SJ+QQRZ1fisD8UTVDSupWJNstVkiqLq+ISTdEjJKGjVfIcsgA4l9cbk8Smlzddh4
# EfvFrpVNnes4c16Jidj5XiPVdsn5n10jxmGpxoMc6iPkoaDhi6JjHd5ibfdp5uzI
# Xp4P0wXkgNs+CO/CacBqU0R4k+8h6gYldp4FCMgrXdKWfM4N0u25OEAuEa3Jyidx
# W48jwBqIJqImd93NRxvd1aepSeNeREXAu2xUDEW8aqzFQDYmr9ZONuc2MhTMizch
# NULpUEoA6Vva7b1XCB+1rxvbKmLqfY/M/SdV6mwWTyeVy5Z/JkvMFpnQy5wR14GJ
# cv6dQ4aEKOX5AgMBAAGjggGLMIIBhzAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/
# BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAgBgNVHSAEGTAXMAgGBmeBDAEE
# AjALBglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZbU2FL3MpdpovdYxqII+eyG8w
# HQYDVR0OBBYEFJ9XLAN3DigVkGalY17uT5IfdqBbMFoGA1UdHwRTMFEwT6BNoEuG
# SWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJTQTQw
# OTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAGCCsGAQUFBwEBBIGDMIGAMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wWAYIKwYBBQUHMAKG
# TGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJT
# QTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQwDQYJKoZIhvcNAQELBQADggIB
# AD2tHh92mVvjOIQSR9lDkfYR25tOCB3RKE/P09x7gUsmXqt40ouRl3lj+8QioVYq
# 3igpwrPvBmZdrlWBb0HvqT00nFSXgmUrDKNSQqGTdpjHsPy+LaalTW0qVjvUBhcH
# zBMutB6HzeledbDCzFzUy34VarPnvIWrqVogK0qM8gJhh/+qDEAIdO/KkYesLyTV
# OoJ4eTq7gj9UFAL1UruJKlTnCVaM2UeUUW/8z3fvjxhN6hdT98Vr2FYlCS7Mbb4H
# v5swO+aAXxWUm3WpByXtgVQxiBlTVYzqfLDbe9PpBKDBfk+rabTFDZXoUke7zPgt
# d7/fvWTlCs30VAGEsshJmLbJ6ZbQ/xll/HjO9JbNVekBv2Tgem+mLptR7yIrpaid
# RJXrI+UzB6vAlk/8a1u7cIqV0yef4uaZFORNekUgQHTqddmsPCEIYQP7xGxZBIhd
# mm4bhYsVA6G2WgNFYagLDBzpmk9104WQzYuVNsxyoVLObhx3RugaEGru+SojW4dH
# PoWrUhftNpFC5H7QEY7MhKRyrBe7ucykW7eaCuWBsBb4HOKRFVDcrZgdwaSIqMDi
# CLg4D+TPVgKx2EgEdeoHNHT9l3ZDBD+XgbF+23/zBjeCtxz+dL/9NWR6P2eZRi7z
# cEO1xwcdcqJsyz/JceENc2Sg8h3KeFUCS7tpFk7CrDqkMIIHWzCCBUOgAwIBAgIQ
# CLGfzbPa87AxVVgIAS8A6TANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQGEwJVUzEX
# MBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0
# ZWQgRzQgQ29kZSBTaWduaW5nIFJTQTQwOTYgU0hBMzg0IDIwMjEgQ0ExMB4XDTIz
# MTExNTAwMDAwMFoXDTI2MTExNzIzNTk1OVowYzELMAkGA1UEBhMCR0IxFDASBgNV
# BAcTC1doaXRsZXkgQmF5MR4wHAYDVQQKExVBTkRSRVdTVEFZTE9SLkNPTSBMVEQx
# HjAcBgNVBAMTFUFORFJFV1NUQVlMT1IuQ09NIExURDCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBAMOkYkLpzNH4Y1gUXF799uF0CrwW/Lme676+C9aZOJYz
# pq3/DIa81oWv9b4b0WwLpJVu0fOkAmxI6ocu4uf613jDMW0GfV4dRodutryfuDui
# t4rndvJA6DIs0YG5xNlKTkY8AIvBP3IwEzUD1f57J5GiAprHGeoc4UttzEuGA3yS
# qlsGEg0gCehWJznUkh3yM8XbksC0LuBmnY/dZJ/8ktCwCd38gfZEO9UDDSkie4VT
# Y3T7VFbTiaH0bw+AvfcQVy2CSwkwfnkfYagSFkKar+MYwu7gqVXxrh3V/Gjval6P
# dM0A7EcTqmzrCRtvkWIR6bpz+3AIH6Fr6yTuG3XiLIL6sK/iF/9d4U2PiH1vJ/xf
# dhGj0rQ3/NBRsUBC3l1w41L5q9UX1Oh1lT1OuJ6hV/uank6JY3jpm+OfZ7YCTF2H
# kz5y6h9T7sY0LTi68Vmtxa/EgEtG6JVNVsqP7WwEkQRxu/30qtjyoX8nzSuF7Tms
# RgmZ1SB+ISclejuqTNdhcycDhi3/IISgVJNRS/F6Z+VQGf3fh6ObdQLVwoT0JnJj
# bD8PzJ12OoKgViTQhndaZbkfpiVifJ1uzWJrTW5wErH+qvutHVt4/sEZAVS4PNfO
# cJXR0s0/L5JHkjtM4aGl62fAHjHj9JsClusj47cT6jROIqQI4ejz1slOoclOetCN
# AgMBAAGjggIDMIIB/zAfBgNVHSMEGDAWgBRoN+Drtjv4XxGG+/5hewiIZfROQjAd
# BgNVHQ4EFgQU0HdOFfPxa9Yeb5O5J9UEiJkrK98wPgYDVR0gBDcwNTAzBgZngQwB
# BAEwKTAnBggrBgEFBQcCARYbaHR0cDovL3d3dy5kaWdpY2VydC5jb20vQ1BTMA4G
# A1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzCBtQYDVR0fBIGtMIGq
# MFOgUaBPhk1odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVk
# RzRDb2RlU2lnbmluZ1JTQTQwOTZTSEEzODQyMDIxQ0ExLmNybDBToFGgT4ZNaHR0
# cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29kZVNpZ25p
# bmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5jcmwwgZQGCCsGAQUFBwEBBIGHMIGEMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXAYIKwYBBQUHMAKG
# UGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNENv
# ZGVTaWduaW5nUlNBNDA5NlNIQTM4NDIwMjFDQTEuY3J0MAkGA1UdEwQCMAAwDQYJ
# KoZIhvcNAQELBQADggIBAEkRh2PwMiyravr66Zww6Pjl24KzDcGYMSxUKOEU4byk
# cOKgvS6V2zeZIs0D/oqct3hBKTGESSQWSA/Jkr1EMC04qJHO/Twr/sBDCDBMtJ9X
# AtO75J+oqDccM+g8Po+jjhqYJzKvbisVUvdsPqFll55vSzRvHGAA6hjyDyakGLRO
# cNaSFZGdgOK2AMhQ8EULrE8Riri3D1ROuqGmUWKqcO9aqPHBf5wUwia8g980sTXq
# uO5g4TWkZqSvwt1BHMmu69MR6loRAK17HvFcSicK6Pm0zid1KS2z4ntGB4Cfcg88
# aFLog3ciP2tfMi2xTnqN1K+YmU894Pl1lCp1xFvT6prm10Bs6BViKXfDfVFxXTB0
# mHoDNqGi/B8+rxf2z7u5foXPCzBYT+Q3cxtopvZtk29MpTY88GHDVJsFMBjX7zM6
# aCNKsTKC2jb92F+jlkc8clCQQnl3U4jqwbj4ur1JBP5QxQprWhwde0+MifDVp0vH
# ZsVZ0pnYMCKSG5bUr3wOU7EP321DwvvEsTjCy/XDgvy8ipU6w3GjcQQFmgp/BX/0
# JCHX+04QJ0JkR9TTFZR1B+zh3CcK1ZEtTtvuZfjQ3viXwlwtNLy43vbe1J5WNTs0
# HjJXsfdbhY5kE5RhyfaxFBr21KYx+b+evYyolIS0wR6New6FqLgcc4Ge94yaYVTq
# MYIGUzCCBk8CAQEwfTBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgQ29kZSBTaWduaW5n
# IFJTQTQwOTYgU0hBMzg0IDIwMjEgQ0ExAhAIsZ/Ns9rzsDFVWAgBLwDpMA0GCWCG
# SAFlAwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcN
# AQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
# LwYJKoZIhvcNAQkEMSIEIODH/kBg13QKc7F+WbdRj4q9ETdaq20TTP/TCK4JFQKL
# MA0GCSqGSIb3DQEBAQUABIICACNvAhsFd91Ngd85/pqNHqdDYnhx31CYaCVfjssX
# gzin9hLJNQLJOciJbtjt6cwGGYgKU8GmEbZ3FXBvg2Agf7iMC6ytLYHfwOiP10ii
# s0IW1nMFZGnro+iGmZ/lhlduZFrfqDrmHtu3vMuTnG6aifByB1+q24LTPtDqiUJs
# 2/fKGOAwOnhEPqYte6s7TpeeiWb0mNpNZOUle6eLbWUMdaj9JEGQcnrYC5lG9QxK
# unzpO2wcTHUxP7dYctiwhnt53RFUKQdH1N+8Y4xiO3znZgItM/WvtGuM9lk9naqy
# 4EHyieJvPZ8WAuKZW7nl6tTGHu5KfXCsIQDg5nVoWyPyq2rBGjj7lIyELbu5HnH+
# VagdfCn9gjCvXo+80iUang2qodqjoUYHgsK8PL8TFOUpuJk3Nw4xIx4DWpn205up
# SAHZAU6lRjCXdXYJ0eo0sjSJ5ijsJZk/isgsmPoyF9/dffRwD1ycjDzx9wEUxFVo
# YUzyXBqFF0epOzcPvFTkauNP1kyDY3CEhghPQ1y9HSSVf+4NfHWOQKhHnAvf68+t
# 7jyH4D1mVPjnmCuU1CILdJQxIKKfOZALXmqnP/3YlLc2ojnjvxns2WgD5Kjlo2Xw
# rt+qaxOWikR3s3Gvr6MwFBnv6nqYqUbaKeTmNsJgM6CjDEkmczfWRbbt4WsKXUjZ
# wfh/oYIDIDCCAxwGCSqGSIb3DQEJBjGCAw0wggMJAgEBMHcwYzELMAkGA1UEBhMC
# VVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBU
# cnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6
# f5WHxvnpBOMzBDANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG
# 9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1MDIyMjE4NDE0NlowLwYJKoZIhvcNAQkE
# MSIEIPWBP/j4z1mGhmIs8VaG5MphH5SWxQPXpCDhd0IUdTxYMA0GCSqGSIb3DQEB
# AQUABIICADD9eqUHvJES9ihu2GiyK+mXdl2uD2Pc2tJ7L+0mqkX4payHvlkhotgY
# ML3EsuLAvI/iCo/yuLuHNvoS+9IpLqypo2QxWwF4pR++S7g1ZLcFCHzV/vDoaTws
# 21AO+wSBU4jgGPbXY9ugpfVZ0B3apg+xOLwXGCsIPPecUBVQVscOKlcRVsRF5IML
# omaOcbYqXyL1cHiT5LbCRgWSfuFMdf9AL4/ZwAvxA44v2VVY9dQei/+iAAWe6ssd
# v6z3JuuPTJHWEc8dlLYvJ9v9LqWtMghyk3LoZwEgC9c1k6h/iECuBEst56SImxVh
# VAdCvT+hkoO89peGxyxair/MXx8HIsvKH2wdAAs/akMsgb58YJRUv2y0zc6dtoZt
# 4GQwM8X+R+JRWTsjJF4DGRZXt+tNYNK++nPsfbPc61SwPViUlwjtzInyKY1LBnOu
# s3Ccn2lgdBGQ+AlyW8iKw/Y3QkxZvfopeDhRTsJ/DQfQCR9YTTD136+O4GzHNoxj
# XOWVf+aFIjyEQy7f65wxCah8FZ1qMvoaiILHTUXG+aYQ/97W2LoICnRvghExlqQL
# PcaSeYg8lFo9dHv6rftLLQlABfyysgbmUyJDeLg27MR05IKrzdKII+piQ+gXdk38
# rEVYokebb4jRp0DBx+4jfYC+tLoy2/e0aEK6gR3wU8FQJWa92qdB
# SIG # End signature block
