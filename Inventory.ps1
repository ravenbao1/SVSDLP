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

# Retrieve Intune devices information
try {
    $intuneDevices = Invoke-RestMethod -Uri $intuneDevicesUrl -Headers $headers -Method Get
    Write-Output "Device information retrieved successfully from Intune."
} catch {
    Write-Output "Failed to retrieve device information from Intune. Error: $_"
    exit
}

# Install ImportExcel if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -AllowClobber
}

foreach ($device in $intuneDevices.value){
    $device

}

# Prepare and export device information to Excel
$excelData = foreach ($device in $intuneDevices.value) {
    [PSCustomObject]@{
        DeviceName       = $device.deviceName
        User             = $device.UserPrincipalName
        OperatingSystem  = $device.operatingSystem
        OSVersion        = $device.OSVersion
        ComplianceState  = $device.complianceState
        Model            = $device.Model
        Manufacturer     = $device.Manufacturer
        SerialNumber     = $device.SerialNumber
        MAC              = $device.WiFiMacAddress
        LastSyncDateTime = $device.lastSyncDateTime
        Encryption       = $device.IsEncrypted
        ReportTime       = Get-Date
        Source           = "Intune"
        Remarks          = ""
    }
    $device.trusttype
}

$excelFilePath = "C:\Temp\DeviceInventory.xlsx"

# Check if the file path exists, if not create it
if (-not (Test-Path -Path (Split-Path -Path $excelFilePath -Parent))) {
    New-Item -Path (Split-Path -Path $excelFilePath -Parent) -ItemType Directory -Force 
}

try {
    $excelData | Export-Excel -Path $excelFilePath -AutoSize
    Write-Output "Device information exported successfully to $excelFilePath"
} catch {
    Write-Output "Failed to export device information to Excel. Error: $_"
    exit
}
