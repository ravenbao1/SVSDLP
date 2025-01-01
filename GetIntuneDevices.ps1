# Install necessary PowerShell modules
# Install-Module -Name Microsoft.Graph.Intune -Force -AllowClobber
# Install-Module -Name ImportExcel -Force -AllowClobber

# Azure AD App Registration details
$clientId = "2ca3083d-78b8-4fe5-a509-f0ea09d2f7da"
$clientSecret = "F0-8Q~vY.lPo.4cXpAIQ-xzlV4wDRBbQsIvmbctE"
$tenantId = "1acfda79-c7e2-4b4b-8bcc-06449fbe9213"

# Authenticate to Microsoft Graph to retrieve Intune data
$tokenBody = @{  
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}

try {
    Write-Output "Attempting to retrieve access token..."
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -ContentType "application/x-www-form-urlencoded" -Body $tokenBody
    if ($tokenResponse -and $tokenResponse.access_token) {
        $token = $tokenResponse.access_token
        Write-Output "Access token retrieved successfully."
    } else {
        Write-Error "Failed to retrieve access token. Response: $tokenResponse"
        exit
    }
} catch {
    Write-Output "Failed to retrieve access token. Error details: $_"
    exit
}

# Set headers for Microsoft Graph API requests
$headers = @{
    Authorization = "Bearer $token"
    ContentType   = "application/json"
}

# Define Intune Graph API endpoints 
$intuneDevicesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
$entraDevicesUrl = "https://graph.microsoft.com/v1.0/devices`?$`select=displayName,deviceId,trustType"
$userUrl = "https://graph.microsoft.com/v1.0/users`?$`select=userPrincipalName,displayName,jobTitle,department,city,country,usageLocation"


# Function to retrieve all paginated results
function Get-AllObjects {
    param (
        [string]$Url,
        [hashtable]$Headers
    )

    $allObjects = @()

    do {
        # Get the data from the current page
        $response = Invoke-RestMethod -Uri $Url -Headers $headers -Method Get
        
        if ($response.value) {
            $allObjects += $response.value
        }

        # Check for the next link
        $Url = $response.'@odata.nextLink'
    } while ($Url) # Continue if there's a next page

    return $allObjects
}


# Retrieve all devices
try {
    Write-Output "Retrieving all Intune device information..."
    $allIntuneDevices = Get-AllObjects -Url $intuneDevicesUrl -Headers $headers -Method Get
    Write-Output "Successfully retrieved all device information from Intune."
    Write-Output "Retrieving all User information..."
    $allUsers = Get-AllObjects -Url $userUrl -Headers $headers -Method Get
    Write-Output "Successfully retrieved all user information from Entra."
    Write-Output "Retrieving all Entra device information..."
    $allEntraDevices = Get-AllObjects -Url $entraDevicesUrl -Headers $headers -Method Get
    Write-Output "Successfully retrieved all device information from Entra."
} catch {
    Write-Output "Failed to retrieve information. Error: $_"
    exit
}

# Prepare and export all devices to Excel
$excelIntuneData = foreach ($device in $allIntuneDevices) {
    [PSCustomObject]@{
        DeviceName       = "$($device.deviceName)"
        UserPrincipalName= "$($device.UserPrincipalName)"
        EntraDeviceID    = "$($device.AzureAdDeviceId)"
        IntuneDeviceID   = "$($device.id)"
        OperatingSystem  = "$($device.operatingSystem)"
        OSVersion        = "$($device.OSVersion)"
        ComplianceState  = "$($device.complianceState)"
        Model            = "$($device.Model)"
        Manufacturer     = "$($device.Manufacturer)"
        SerialNumber     = "$($device.SerialNumber)" # Retains leading zeros
        MAC              = "$($device.WiFiMacAddress)"
        IntuneLastSync   = "$($device.lastSyncDateTime)"
        Encryption       = "$($device.IsEncrypted)"
        TotalStorage     = "$($device.totalStorageSpaceInBytes)"
        FreeStorage      = "$($device.freeStorageSpaceInBytes)"
        PhysicalMemory   = ""
        Source           = "Cloud"
        Remarks          = ""
    }
}

$excelEntraDeviceData = foreach ($entraDevice in $allEntraDevices) {
    [PSCustomObject]@{
        DeviceName       = "$($entraDevice.displayName)"
        EntraDeviceID    = "$($entraDevice.deviceId)"
        TrustType        = "$($entraDevice.trustType)"
    }
}

$excelEntraUserData = foreach ($user in $allUsers) {
    [PSCustomObject]@{
        UPN              = "$($user.UserPrincipalName)"
        UserDisplayName  = "$($user.displayName)"
        JobTitle         = "$($user.jobTitle)"
        Department       = "$($user.department)"
        City             = "$($user.city)"
        Country          = "$($user.country)"
    }
}

# Paths for CSV files
$csvIntunePath = Join-Path -Path $PSScriptRoot -ChildPath "IntuneDeviceReport.csv"
$csvEntraDevicePath = Join-Path -Path $PSScriptRoot -ChildPath "EntraDeviceReport.csv"
$csvEntraUserPath = Join-Path -Path $PSScriptRoot -ChildPath "UserReport.csv"

try {
    Write-Output "Exporting data to CSV..."

    # Export Intune data to CSV
    $excelIntuneData | Export-Csv -Path $csvIntunePath -NoTypeInformation -Encoding UTF8
    Write-Output "Intune Device information exported successfully to $csvIntunePath"

    # Export Entra device data to CSV
    $excelEntraDeviceData | Export-Csv -Path $csvEntraDevicePath -NoTypeInformation -Encoding UTF8
    Write-Output "Entra Device information exported successfully to $csvEntraDevicePath"

    # Export Entra user data to CSV
    $excelEntraUserData | Export-Csv -Path $csvEntraUserPath -NoTypeInformation -Encoding UTF8
    Write-Output "User information exported successfully to $csvEntraUserPath"
} catch {
    Write-Output "Failed to export device information to CSV. Error: $_"
    exit
}

