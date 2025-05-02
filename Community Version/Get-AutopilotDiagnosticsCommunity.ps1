
<#PSScriptInfo

.VERSION 6.2
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
Version 6.2: Bug fixes.
Version 6.1: Bug fixes.
Verison 6.0: Added APv2 support and various other enhancements
Version 5.14: Added auto-install of Graph modules
Version 5.13: Fixed issue with bearer
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
 
#> 
<#
.SYNOPSIS
Displays Windows Autopilot diagnostics information from the current PC or a captured set of logs.
 
.PARAMETER Online
Look up the actual policy and app names via the Microsoft Graph API
 
.PARAMETER AllSessions
Show all ESP progress instead of just the final details.
 
.PARAMETER File
Processes the information in the specified file (captured either by MDMDiagnosticsTool.exe -area Autopilot -cab filename.cab or MDMDiagnosticsTool.exe -area Autopilot -zip filename.zip) instead of from the registry.
 
.PARAMETER ShowPolicies
Shows the policy details as recorded in the NodeCache registry keys, in the order that the policies were received by the client.

.PARAMETER Tenant
The GUID (text string), needed when specifying an app ID/secret.
 
.PARAMETER AppId
The app ID (GUID) for the Entra ID app being used to authenticate with Intune

.PARAMETER AppSecret
The app secret (effectively a password) for the specified app ID.
 
.PARAMETER Bearer
An existing bearer token that will be used to authentcate to Intune.

.EXAMPLE
.\Get-AutopilotDiagnostics.ps1
 
.EXAMPLE
.\Get-AutopilotDiagnostics.ps1 -Online
 
.EXAMPLE
.\Get-AutopilotDiagnostics.ps1 -AllSessions
 
.EXAMPLE
.\Get-AutopilotDiagnostics.ps1 -File C:\Autopilot.cab -Online -AllSessions
 
.EXAMPLE
.\Get-AutopilotDiagnostics.ps1 -File C:\Autopilot.zip
 
.EXAMPLE
.\Get-AutopilotDiagnostics.ps1 -ShowPolicies
 
#>

[CmdletBinding()]
param(
    [Alias("CABFile","ZIPFile","FullName")][Parameter(Mandatory = $False, ValueFromPipelineByPropertyName = $true)] [String] $File = $null,
    [Parameter(Mandatory = $False)] [Switch] $Online = $false,
    [Parameter(Mandatory = $False)] [Switch] $AllSessions = $false,
    [Parameter(Mandatory = $False)] [Switch] $ShowPolicies = $false,
    [Parameter(Mandatory = $false)] [string] $Tenant,
    [Parameter(Mandatory = $false)] [string] $AppId,
    [Parameter(Mandatory = $false)] [string] $AppSecret,
    [Parameter(Mandatory = $false)] [string] $Bearer
)

Begin {

    # Configure constants and global variables
    $script:officeStatus = @{"0" = "None"; "10" = "Initialized"; "20" = "Download In Progress"; "25" = "Pending Download Retry";
        "30" = "Download Failed"; "40" = "Download Completed"; "48" = "Pending User Session"; "50" = "Enforcement In Progress"; 
        "55" = "Pending Enforcement Retry"; "60" = "Enforcement Failed"; "70" = "Success / Enforcement Completed"
    }
    $script:espStatus = @{"1" = "Not Installed"; "2" = "Downloading / Installing"; "3" = "Success / Installed"; "4" = "Error / Failed" }
    $script:policyStatus = @{"0" = "Not Processed"; "1" = "Processed" }

    enum AutopilotScenarioEnum {
        Unknown
        AutopilotV1
        AutopilotJson
        EspOnly
        AutopilotV2
    }

    enum WorkloadState
    {
        NotStarted
        Completed
        Skipped
        Uninstalled
        Failed
        InProgress
        RebootRequired
        Cancelled
    }
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
            [Parameter(Mandatory = $false)] [string] $Tenant,
            [Parameter(Mandatory = $false)] [string] $AppId,
            [Parameter(Mandatory = $false)] [string] $AppSecret,
            [Parameter(Mandatory = $false)] [string] $scopes,
            [Parameter(Mandatory = $false)] [string] $Bearer
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
                    Write-Verbose "Version 2 module detected"
                    $accesstokenfinal = ConvertTo-SecureString -String $accessToken -AsPlainText -Force
                }
                else {
                    Write-Verbose "Version 1 Module Detected"
                    Select-MgProfile -Name Beta
                    $accesstokenfinal = $accessToken
                }
                Connect-MgGraph -AccessToken $accesstokenfinal -NoWelcome
                Write-Verbose "Connected to Intune tenant $Tenant using app-based authentication (Azure AD authentication not supported)"
            }
            elseif ($Bearer -ne "") {
                if ($version -eq 2) {
                    Write-Verbose "Version 2 module detected"
                    $accesstokenfinal = ConvertTo-SecureString -String $Bearer -AsPlainText -Force
                }
                else {
                    Write-Verbose "Version 1 Module Detected"
                    Select-MgProfile -Name Beta
                    $accesstokenfinal = $Bearer
                }
                Connect-MgGraph -AccessToken $accesstokenfinal -NoWelcome
                Write-Verbose "Connected to Intune tenant $Tenant using app-based authentication (Azure AD authentication not supported)"
            }
            else {
                if ($version -eq 2) {
                    Write-Verbose "Version 2 module detected"
                }
                else {
                    Write-Verbose "Version 1 Module Detected"
                    Select-MgProfile -Name Beta
                }
                Connect-MgGraph -scopes $scopes -NoWelcome
            }
            # Return the context
            $graph = Get-MgContext
            Write-Host "Connected to Intune tenant $($graph.TenantId)"
            $graph
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
            }
            else {
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
                    }
                    elseif ($currentUser -eq "S-0-0-00-0000000000-0000000000-000000000-000") {
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

    Function ProcessSidecarV2() {
        param
        (
            [Parameter(ValueFromPipelineByPropertyName = $True)] [bool] $display = $true
        )

        Begin {
            if ($display) { Write-Host "Sidecar apps:" }
            if ($null -eq $script:DOEvents -and (-not $script:useFile)) {
                $script:DOEvents = Get-DeliveryOptimizationLog | Where-Object { $_.Function -match "(DownloadStart)|(DownloadCompleted)" -and $_.Message -like "*.intunewin.bin,*" }
            }
        }

        Process {
            if (Test-Path "$sidecarWin32Apps\ProvisioningProgress") {
                $details = Get-ItemPropertyValue -Path "$sidecarWin32Apps\ProvisioningProgress" -Name "ProvisioningProgress"
                if ($details) {
                    $provisioningProgress = $details | ConvertFrom-Json
                    $provisioningProgress.Workloads | ForEach-Object {
                        # "WorkloadId":"41e931ef-9951-4646-aa00-6df474a5d66d","FriendlyName":"PowerToys 0.90.1","WorkloadState":1,"StartTime":"\/Date(1745871176634)\/","EndTime":"\/Date(1745871256672)\/","ErrorCode":null }
                        RecordStatus -detail $_.FriendlyName -status "Installation started" -color "Yellow" -date $_.StartTime
                        $status = [WorkloadState]$_.WorkloadState
                        if ($status -eq [WorkloadState]::Completed) {
                            if ($display) { Write-Host " $($_.FriendlyName) : $status" -ForegroundColor Green }
                            RecordStatus -detail $_.FriendlyName -status $status -color "Green" -date $_.EndTime
                        } elseif ($status -eq [WorkloadState]::Failed) {
                            $enforcementStatus = Get-ItemPropertyValue -Path "$sidecarWin32Apps\00000000-0000-0000-0000-000000000000\$($_.WorkloadId)*\EnforcementStateMessage" -Name EnforcementStateMessage | ConvertFrom-Json
                            if ($display) { Write-Host " $($_.FriendlyName) : $status, rc = $($enforcementStatus.ErrorCode)" -ForegroundColor Red }
                            RecordStatus -detail $_.FriendlyName -status $status -color "Red" -date $_.EndTime
                        } else {
                            if ($display) { Write-Host " $($_.FriendlyName) : $status" -ForegroundColor Yellow }
                            RecordStatus -detail $_.FriendlyName -status $status -color "Yellow" -date $_.EndTime
                        }

                        # Try to find the DO events.
                        if ($script:DOEvents) {
                            $appName = $_.FriendlyName
                            $appId = $_.WorkloadId
                            $script:DOEvents | Where-Object { $_.Message -ilike "*$appId*" } | ForEach-Object {
                                if ($_.Function.Contains("DownloadStart")) 
                                {
                                    $op = "DownloadStart"
                                } else {
                                    $op = "DownloadCompleted"
                                }
                                RecordStatus -detail $appName -status "DO $op" -color "Yellow" -date $_.TimeCreated
                            }    
                        }
                    }
                } else {
                    if ($display) { Write-Host " Provisioning progress details not found." }
                }            
            } else {
                if ($display) { Write-Host " Provisioning progress not found." }
            }
        }
    }

    Function ProcessSidecarV2Scripts() {
        param
        (
            [Parameter(ValueFromPipelineByPropertyName = $True)] [bool] $display = $true
        )

        Begin {
            if ($display) { Write-Host "Sidecar scripts:" }
        }

        Process {
            if (Test-Path "$sidecarPath\Policies\00000000-0000-0000-0000-000000000000") {
                Get-ChildItem -Path "$sidecarPath\Policies\00000000-0000-0000-0000-000000000000" | ForEach-Object {
                    $scriptId = $_.PSChildName
                    $scriptName = $scriptId
                    $properties = Get-ItemProperty -Path $_.PSPath
                    $result = $properties.Result
                    $when = [DateTime]$properties.LastUpdatedTimeUtc
                    if ($Online) {
                        $scripts | Where-Object { $scriptId -eq $_.Id } | ForEach-Object {
                            $scriptName = "$($_.DisplayName) (script)"                       
                        }
                    }
                    if ($result -eq "Success") {
                        if ($display) { Write-Host " $scriptName : $result" -ForegroundColor Green }
                        RecordStatus -detail $scriptName -status $result -color "Green" -date $when
                    } elseif ($result -eq "Failed") {
                        if ($display) { Write-Host " $scriptName : $result" -ForegroundColor Red }
                        RecordStatus -detail $scriptName -status $result -color "Red" -date $when
                    } else {
                        if ($display) { Write-Host " $scriptName : $result" -ForegroundColor Yellow }
                        RecordStatus -detail $scriptName -status $result -color "Red" -date $when
                    }
                }            
            } else {
                if ($display) { Write-Host " Provisioning script info not found." }
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
                $script:DOEvents = Get-DeliveryOptimizationLog | Where-Object { $_.Function -match "(DownloadStart)|(DownloadCompleted)" -and $_.Message -like "*.intunewin.bin,*" }
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
                $sidecarApp = "$sidecarWin32Apps\$currentUser\$appGuid"
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
                        if ($display) { Write-Host " Win32 $win32Key : $status ($($espStatus[$status.ToString()]), rc = $exitCode" -ForegroundColor Red }
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
                            if ($_.Function.Contains("DownloadStart")) 
                            {
                                $op = "DownloadStart"
                            } else {
                                $op = "DownloadCompleted"
                            }
                            RecordStatus -detail "Win32 $win32Key" -status "DO $op" -color "Yellow" -date $_.TimeCreated.ToLocalTime()
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
            [object] $e,
            [string] $sidecarProductCode
        )

        # Fix up the name
        if ($event.Id -eq 1924) {
            $r = $event.Properties[2].Value
        } else {
            $r = $event.Properties[0].Value
        }
        $productCode = $r.Replace("{","").Replace("}","")
        if ($productCode -eq $sidecarProductCode) {
            return "Intune Management Extension ($r)"
        }

        # See if we can find the real name
        if (Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{{$productCode}}") {
            $displayName = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$productCode").DisplayName
            return "$displayName ($r)"
        }
        else {
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
                        $productCode = (Get-ItemPropertyValue -Path $_.PSPath -Name ProductCode).Replace("{","").Replace("}","")
                    }
                }
            }

            # Process device management events
            if ($script:useFile) {
                $events = Get-WinEvent -Path "$($env:TEMP)\ESPStatus.tmp\microsoft-windows-devicemanagement-enterprise-diagnostics-provider-admin.evtx" -Oldest | ? { ($_.Id -in 1905, 1906, 1920, 1922, 1924) -or $_.Id -in (72, 100, 107, 109, 110, 111) }
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
                    1924 { $detail = (TrimMSI $event $productCode); $message = "Installation failed"; $color = "Red" }
                    { $_ -in (1922, 72) } { $color = "Green" }
                }
                RecordStatus -detail $detail -date $_.TimeCreated.ToUniversalTime() -status $message -color $color
            }

            # Process device registration events
            if ($script:useFile) {
                $events = Get-WinEvent -Path "$($env:TEMP)\ESPStatus.tmp\microsoft-windows-user device registration-admin.evtx" -Oldest | ? { $_.Id -in (306, 101) }
            }
            else {
                try {
                    $events = Get-WinEvent -LogName 'Microsoft-Windows-User Device Registration/Admin' -Oldest -ErrorAction Stop | ? { $_.Id -in (306, 101) }
                }
                catch [Exception] {
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
                    101 { $detail = "Device Registration"; $message = "SCP discovery successful" }
                    304 { $detail = "Device Registration"; $message = "Hybrid AADJ device registration failed" }
                    306 { $detail = "Device Registration"; $message = "Hybrid AADJ device registration succeeded"; $color = 'Green' }
                }
                RecordStatus -detail $detail -date $_.TimeCreated.ToUniversalTime() -status $message -color $color
            }

            # Add DO events for Office click-to-run downloads
            if (-not $script:useFile) {
                Get-DeliveryOptimizationLog | Where-Object { $_.Function -match "(DownloadStart)|(DownloadCompleted)" -and $_.Message -like "*Microsoft Office Click-to-Run*" } | ForEach-Object {
                    # Extract the file ID because we want to list each file downloaded
                    $fileId = ""
                    $fileIdStart = $_.Message.IndexOf("fileId: ")
                    if ($fileIdStart -eq -1) {
                        # Might be using "fileId = ", because this DO event information sucks
                        $fileIdStart = $_.Message.IndexOf("fileId = ")
                        $skip = 9
                    } else {
                        $skip = 8
                    }
                    if ($fileIdStart -gt 0) {
                        # Get from the start of the actual ID
                        $fileId = $_.Message.Substring($fileIdStart + $skip)
                        # Find the end and chop it off
                        $fileIdEnd = $fileId.IndexOf(",")
                        $fileId = $fileId.Substring(0, $fileIdEnd)
                        # Remove the extra GUID from the beginning
                        $fileId = $fileId.Substring(37)
                    }
                    if ($_.Function.Contains("DownloadStart")) 
                    {
                        $op = "DownloadStart"
                    } else {
                        $op = "DownloadCompleted"
                    }
                    RecordStatus -detail "Microsoft Office C2R ($fileId)" -status $op -color "Yellow" -date $_.TimeCreated
                }
            }

        }
    
    }
    
    #------------------------
    # Main code
    #------------------------

    $script:observedTimeline = @()

    # If online, make sure we are able to authenticate
    if ($Online) {

        ##Check if we need to install the module
        #Install MS Graph if not available
        if (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication) {
            Write-Verbose "Microsoft Graph already installed"
        } 
        else {
            try {
                Install-Module -Name Microsoft.Graph.Authentication -Repository PSGallery -Force 
            }
            catch [Exception] {
                $_.message 
            }
        }

        #Connect to Graph
        if ($AppId -and $AppSecret -and $tenant) {
            $graph = Connect-ToGraph -Tenant $tenant -AppId $clientid -AppSecret $clientsecret
        }
        elseif ($Bearer) {
            $graph = Connect-ToGraph -bearer $Bearer
        }
        else {
            $graph = Connect-ToGraph -Scopes "DeviceManagementApps.Read.All, DeviceManagementConfiguration.Read.All"
        }

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

        # Get a list of platform scripts
        Write-Host "Getting list of scripts"
        $scriptsuri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts"
        #$script:policies = Get-MgBetaDeviceManagementConfigurationPolicy -All
        $script:scripts = getallpagination -url $scriptsuri
    }

    # Process log files if needed
    $script:useFile = $false
    if ($File) {

        Write-Host "Using contents of file: $File"
        if (Test-Path "$($env:TEMP)\ESPStatus.tmp") {
            Remove-Item "$($env:TEMP)\ESPStatus.tmp" -Recurse -Force
        }
        New-Item -Path "$($env:TEMP)\ESPStatus.tmp" -ItemType "directory" | Out-Null
        $script:useFile = $true

        # If using a CAB file, extract the needed files from it
        if ($File.ToLower().EndsWith(".cab")) {
            $null = & expand.exe "$File" -F:* "$($env:TEMP)\ESPStatus.tmp\" 
        }
        else {
            # If using a ZIP file, just extract the entire contents (not as easy to do selected files)
            Expand-Archive -Path $File -DestinationPath "$($env:TEMP)\ESPStatus.tmp\"
            # If this is an Intune diagnostics zip, the "real" logs are buried deeper.  If we can find them, extract them
            $realFile = Get-ChildItem "$($env:TEMP)\ESPStatus.tmp" -Filter "*MDMDiagnostics*_cab" | Get-ChildItem
            if ($realFile) {
                # Expand them into the temp folder -- creates a bit of a mess, but it's a temporary mess...
                $null = & expand.exe "$($realFile.FullName)" -F:* "$($env:TEMP)\ESPStatus.tmp\"
            }
        }

        # Get the hardware hash information
        Get-ChildItem "$($env:TEMP)\ESPStatus.tmp" -Filter "*.csv" | ForEach-Object {
            $csv = Get-Content $_.FullName | ConvertFrom-Csv
            $hash = $csv.'Hardware Hash'
        }

        # Edit the path in the .reg file
        $content = Get-Content -Path "$($env:TEMP)\ESPStatus.tmp\MdmDiagReport_RegistryDump.reg"
        $content = $content -replace "\[HKEY_CURRENT_USER\\", "[HKEY_CURRENT_USER\ESPStatus.tmp\USER\"
        $content = $content -replace "\[HKEY_LOCAL_MACHINE\\", "[HKEY_CURRENT_USER\ESPStatus.tmp\MACHINE\"
        $content = $content -replace '^ "', '"'
        $content = $content -replace '^ @', '@'
        $content = $content -replace 'DWORD:', 'dword:'

        $stream = [System.IO.StreamWriter] "$($env:TEMP)\ESPStatus.tmp\MdmDiagReport_Edited.reg"
        $stream.WriteLine("Windows Registry Editor Version 5.00`n")
        $content | ForEach-Object {
            # Escape backslashes and quotes
            $line = $_.Trim()
            $textStart = $line.IndexOf("""=""") + 3
            if ($textStart -gt 3) {
                $textLen = $line.Length - $textStart
                $toEdit = $line.Substring($textStart, $textLen - 1)
                $toEdit = $toEdit.Replace('\', '\\')
                $toEdit = $toEdit.Replace('"', '\"')
                $line = "$($line.Substring(0, $textStart))$toEdit"""
            }
            # Append it to the file
            # Write-Host $line
            $stream.WriteLine($line)
        }
        $stream.Close()

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
        $script:sidecarPath = "HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\IntuneManagementExtension"
        $script:sidecarWin32Apps = "HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\IntuneManagementExtension\Win32Apps"
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
        $script:sidecarPath = "HKLM:\Software\Microsoft\IntuneManagementExtension"
        $script:sidecarWin32Apps = "HKLM:\Software\Microsoft\IntuneManagementExtension\Win32Apps"
        $script:enrollmentsPath = "HKLM:\Software\Microsoft\enrollments"

        $hash = (Get-WmiObject -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'").DeviceHardwareData
    }

    # Display Autopilot diag details
    Write-Host ""
    Write-Host "AUTOPILOT DIAGNOSTICS" -ForegroundColor Magenta
    Write-Host ""

    # Determine scenario
    $script:AutopilotScenario = [AutopilotScenarioEnum]::Unknown
    $correlations = Get-ItemProperty "$autopilotDiagPath\EstablishedCorrelations"
    $values = Get-ItemProperty "$provisioningPath\AutopilotSettings"
    if ($values.AutopilotDevicePrepHint -eq 0) {
        $script:AutopilotScenario = [AutopilotScenarioEnum]::AutopilotV2
        Write-Host "Scenario: Autopilot device preparation (v2)"
        $settings = Get-ItemProperty -Path "$provisioningPath\AutopilotSettings\DevicePreparation"
        $pageSettings = $settings.PageSettings | ConvertFrom-Json
        # {"AgentDownloadTimeoutSeconds":1800,"PageTimeoutSeconds":3600,"ErrorMessage":"Contact your oganization's support person for help.","AllowSkipOnFailure":true,"AllowDiagnostics":true}
        Write-Host "AgentDownloadTimeoutSeconds: $($pageSettings.AgentDownloadTimeoutSeconds)"
        Write-Host "PageTimeoutSeconds: $($pageSettings.PageTimeoutSeconds)"
        Write-Host "AllowSkipOnFailure: $($pageSettings.AllowSkipOnFailure)"
        Write-Host "AllowDiagnostics: $($pageSettings.AllowDiagnostics)"

        Get-ChildItem $enrollmentsPath | ForEach-Object {
            $properties = Get-ItemProperty -Path $_.PSPath
            if ($properties.ProviderId -eq "MS DM Server") {
                Write-Host "TenantID: $($properties.AADTenantID)"
                Write-Host "UPN: $($properties.UPN)"
            }
        }

    } else {
        $values = Get-ItemProperty "$autopilotDiagPath"
        if ($values.CloudAssignedTenantId) {
            if ($values.DeploymentProfileName -and $values.DeploymentProfileName -ne "") {
                $script:AutopilotScenario = [AutopilotScenarioEnum]::AutopilotV1
                Write-Host "Scenario: Autopilot (v1)"
                Write-Host "Profile: $($values.DeploymentProfileName)"
            } else {
                $script:AutopilotScenario = [AutopilotScenarioEnum]::AutopilotJson
                Write-Host "Scenario: Autopilot for existing devices (v1)"
                Write-Host "Correlation ID: $($values.ZtdCorrelationId)"                
            }
            Write-Host "TenantDomain: $($values.CloudAssignedTenantDomain)"
            Write-Host "TenantID: $($values.CloudAssignedTenantId)"
            Write-Host "OobeConfig: $($values.CloudAssignedOobeConfig)"

            if (($values.CloudAssignedOobeConfig -band 1024) -gt 0) {
                Write-Host " Skip keyboard: Yes             1 - - - - - - - - - -"
            }
            else {
                Write-Host " Skip keyboard: No              0 - - - - - - - - - -"
            }
            if (($values.CloudAssignedOobeConfig -band 512) -gt 0) {
                Write-Host " Enable patch download: Yes     - 1 - - - - - - - - -"
            }
            else {
                Write-Host " Enable patch download: No      - 0 - - - - - - - - -"
            }
            if (($values.CloudAssignedOobeConfig -band 256) -gt 0) {
                Write-Host " Skip Windows upgrade UX: Yes   - - 1 - - - - - - - -"
            }
            else {
                Write-Host " Skip Windows upgrade UX: No    - - 0 - - - - - - - -"
            }
            if (($values.CloudAssignedOobeConfig -band 128) -gt 0) {
                Write-Host " AAD TPM Required: Yes          - - - 1 - - - - - - -"
            }
            else {
                Write-Host " AAD TPM Required: No           - - - 0 - - - - - - -"
            }
            if (($values.CloudAssignedOobeConfig -band 64) -gt 0) {
                Write-Host " AAD device auth: Yes           - - - - 1 - - - - - -"
            }
            else {
                Write-Host " AAD device auth: No            - - - - 0 - - - - - -"
            }
            if (($values.CloudAssignedOobeConfig -band 32) -gt 0) {
                Write-Host " TPM attestation: Yes           - - - - - 1 - - - - -"
            }
            else {
                Write-Host " TPM attestation: No            - - - - - 0 - - - - -"
            }
            if (($values.CloudAssignedOobeConfig -band 16) -gt 0) {
                Write-Host " Skip EULA: Yes                 - - - - - - 1 - - - -"
            }
            else {
                Write-Host " Skip EULA: No                  - - - - - - 0 - - - -"
            }
            if (($values.CloudAssignedOobeConfig -band 8) -gt 0) {
                Write-Host " Skip OEM registration: Yes     - - - - - - - 1 - - -"
            }
            else {
                Write-Host " Skip OEM registration: No      - - - - - - - 0 - - -"
            }
            if (($values.CloudAssignedOobeConfig -band 4) -gt 0) {
                Write-Host " Skip express settings: Yes     - - - - - - - - 1 - -"
            }
            else {
                Write-Host " Skip express settings: No      - - - - - - - - 0 - -"
            }
            if (($values.CloudAssignedOobeConfig -band 2) -gt 0) {
                Write-Host " Disallow admin: Yes            - - - - - - - - - 1 -"
            }
            else {
                Write-Host " Disallow admin: No             - - - - - - - - - 0 -"
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
                    Write-Host "Subscenarios: Hybrid Azure AD Join"
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
                    Write-Host "Subscenario: Azure AD Join"
                }
            }
            else {
                Write-Host "Subscenario: Not available (JSON not found)"
            }

        }
    }

    if (-not $script:useFile) {
        $osVersion = (Get-WmiObject win32_operatingsystem).Version
        Write-Host "OS version: $osVersion"
    }
    Write-Host "EntDMID: $($correlations.EntDMID)"

    # Get ESP properties
    Get-ChildItem $enrollmentsPath | Where-Object { Test-Path "$($_.PSPath)\FirstSync" } | % {
        if ($script:AutopilotScenario -eq [AutopilotScenarioEnum]::Unknown) {
            $script:AutopilotScenario = [AutopilotScenarioEnum]::EspOnly
        }
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
        Write-Host " From Connected Cache: $($ccPct)% ($($stats.DownloadCacheHostBytes))"
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
    
    if ($script:AutopilotScenario -eq [AutopilotScenarioEnum]::AutopilotV2) {

        # Process scripts
        Write-Host " "
        Write-Host "SCRIPTS:" -ForegroundColor Magenta
        Write-Host " "
        ProcessSidecarV2Scripts

        # Process Win32 apps
        Write-Host " "
        Write-Host "APPS:" -ForegroundColor Magenta
        Write-Host " "
        ProcessSidecarV2

    } else {
        # Make sure the tracking path exists
        if (Test-Path $path) {

            # Process device ESP sessions
            Write-Host " "
            Write-Host "DEVICE ESP:" -ForegroundColor Magenta
            Write-Host " "

            if (Test-Path "$path\ExpectedPolicies") {
                [array]$items = Get-ChildItem "$path\ExpectedPolicies"
                if ($items) {
                    AddDisplay ([ref]$items)
                    $items | ProcessPolicies
                }
            }
            if (Test-Path "$path\ExpectedMSIAppPackages") {
                [array]$items = Get-ChildItem "$path\ExpectedMSIAppPackages"
                if ($items) {
                    AddDisplay ([ref]$items)
                    $items | ProcessApps -currentUser "S-0-0-00-0000000000-0000000000-000000000-000" 
                }
            }
            if (Test-Path "$path\ExpectedModernAppPackages") {
                [array]$items = Get-ChildItem "$path\ExpectedModernAppPackages"
                if ($items) {
                    AddDisplay ([ref]$items)
                    $items | ProcessModernApps -currentUser "S-0-0-00-0000000000-0000000000-000000000-000"
                }
            }
            if (Test-Path "$path\Sidecar") {
                [array]$items = Get-ChildItem "$path\Sidecar" | ? { $_.Property -match "./Device" -and $_.Name -notmatch "LastLoggedState" }
                if ($items) {
                    AddDisplay ([ref]$items)
                    $items | ProcessSidecar -currentUser "00000000-0000-0000-0000-000000000000"
                }
            }
            if (Test-Path "$path\ExpectedSCEPCerts") {
                [array]$items = Get-ChildItem "$path\ExpectedSCEPCerts"
                if ($items) {
                    AddDisplay ([ref]$items)
                    $items | ProcessCerts
                }
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
                    if ($items) {
                        AddDisplay ([ref]$items)
                        $items | ProcessPolicies
                    }
                }
                if (Test-Path "$userPath\ExpectedMSIAppPackages") {
                    [array]$items = Get-ChildItem "$userPath\ExpectedMSIAppPackages" 
                    if ($items) {
                        AddDisplay ([ref]$items)
                        $items | ProcessApps -currentUser $userSid
                    }
                }
                if (Test-Path "$userPath\ExpectedModernAppPackages") {
                    [array]$items = Get-ChildItem "$userPath\ExpectedModernAppPackages"
                    if ($items) {
                        AddDisplay ([ref]$items)
                        $items | ProcessModernApps -currentUser $userSid
                    }
                }
                if (Test-Path "$userPath\Sidecar") {
                    [array]$items = Get-ChildItem "$path\Sidecar" | ? { $_.Property -match "./User" }
                    if ($items) {
                        AddDisplay ([ref]$items)
                        $items | ProcessSidecar -currentUser $userSid
                    }
                }
                if (Test-Path "$userPath\ExpectedSCEPCerts") {
                    [array]$items = Get-ChildItem "$userPath\ExpectedSCEPCerts"
                    if ($items) {
                        AddDisplay ([ref]$items)
                        $items | ProcessCerts
                    }
                }
            }
        }
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

    # Remove the temp folder info if it exists
    if (Test-Path "$($env:TEMP)\ESPStatus.tmp") {
        Remove-Item "$($env:TEMP)\ESPStatus.tmp" -Recurse -Force
    }
}



# SIG # Begin signature block
# MIIoEwYJKoZIhvcNAQcCoIIoBDCCKAACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDXzoyja7Z1p+aC
# LIICgJrKiQXkZE0Iecv8q+K0kF9WpqCCIRYwggWNMIIEdaADAgECAhAOmxiO+dAt
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
# LwYJKoZIhvcNAQkEMSIEIHbxDJreTUq2BGWW5xQFD1tNLqHQ8DGMB76eXe9Y3fRo
# MA0GCSqGSIb3DQEBAQUABIICAHgHTCa2BCSHzSlIflyZIAl+vVilA3R/Q5OKi8av
# S50mHExQpC/kPYjsoL9O0eWvizQepvG+7gS7zCH4kkrIIJod21Rk4Qc3cHNV+9x4
# Di/lF5i8QF+2zo+msNuxHCWqseIff/p9gRxMBJXupx4jt6rShDmDDEwLqQIJadA3
# C9ECCcaLht3kpsFrlmYN6tERnCvXuRZPyrnwHJb5uI26NKxNmRe2UmDwwtMSznlC
# hmIxMP7pVtMOFkSBtZsOFpreVlDzehVF8x5deiz33Ify29EzQFrrDThIYZv4O0HH
# uEoSmRz1lsKygpo28p5F46GFrZVsAusrBaP2XXf2zYq777BHc1hK+QRlrp5GpXMv
# ruVCqyQZaoukFokrYtxjVL6A+kBH6jk+hrNYyHBvJFUO7RrIukiNiMbbvS7yXYpL
# rx9H88lje7mXhfKjZTsH4yGIxBJX9fqre9IooZ2mT9TYiTrUWGkxduUVHtinV72Q
# EsL2dlkznT0fiWnNQ9G7zefhNHnJoFwUVnBF+0fJoUS6iE05pIDeyKqf8pJ9qE7t
# 21jQlRLne9XVhKM2PaSafLPFdhK1FmWhZKmgo77BYnC27V5U1LObbSIgh0C/AP4f
# QAAF0jeRVbLUHD/5hMdZuBjzW9/aOss4bbQnnTZJ7Y0pBJDs7IhPVJ0CgAivTcgk
# xnJ+oYIDIDCCAxwGCSqGSIb3DQEJBjGCAw0wggMJAgEBMHcwYzELMAkGA1UEBhMC
# VVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBU
# cnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6
# f5WHxvnpBOMzBDANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG
# 9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1MDUwMTA5MDk0NlowLwYJKoZIhvcNAQkE
# MSIEIIeZ3UETWp98EnoeKe/d4wO/yzVD+sztGyMbXWlyLVdDMA0GCSqGSIb3DQEB
# AQUABIICAEnTSAE006RaV4zsq7uNPL81vdnDSDF/KF6j7Q9lWb7gOryROU8BNEvR
# OtHpTUJz3opK3O4UyAeqjZepLyBfEGXMvcOwZkWJGKK4D/8eNxq4II5FboH0lC1l
# y7AAZCYrvs/My+sPelUiLIA0znUWSU3N/IM4DJ6XjREo8G0s//3AcfAPme8MKz23
# hyXf3zojCuJ/Jl/k0ajlVQ3/x/O7N2kc/ojdAOVtL8luRZ/nLKIOzOksv/Exb4uA
# NTGBIrhpGqTmWzgffYR22XMeinFNuq2ndFhl8NuM1YNr8ooiR3MqlBR8Z7mIQD9o
# USkI5acHYBTOKE/GYtl/oyIH3ylgKqqmR/UToztklzN305aSTlLekidEyRaOIqc0
# UvA++mstsdBfo7XISSqmr6opV9EYcsqPCnrm/zB0A+c4QF/VmzliqSCbW5kVjpy0
# v3yAuVMDofPSLBtDeEvhkKya1gyEZU3I+2FIkhorYn+VxBWA9gmZUZ1IIA92g6VC
# lbQCDdhDtVsxInYCluM1EjfmgmhNa+vUQyj+gz38KWSCwAycq5c39REsla0jGmIJ
# N16VnUtPWBKvZenXs7vm0iTik32mgMZcMmpFXNaomZ6qUJc6k7PHbZ/vsGo2Tu5k
# 58XwYLuH3Kd6uy9Ehh3RoWAqkVFXBZCxQceMo52UbNcA/uWtBkvB
# SIG # End signature block
