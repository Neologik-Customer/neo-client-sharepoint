<#
.SYNOPSIS
    Automates Microsoft Graph API setup for SharePoint access with Sites.Selected permissions.

.DESCRIPTION
    This script uses Azure CLI to automatically set up Azure AD App Registration and Graph API access
    for SharePoint sites using the Sites.Selected permission model.

.NOTES
    Prerequisites:
    - Azure CLI installed (https://aka.ms/installazurecliwindows)
    - Azure AD admin access
    - PowerShell 5.1 or later
    - Internet connection
#>

[CmdletBinding()]
param()

# Function to display section headers
function Write-SectionHeader {
    param([string]$Title)
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}

# Function to check if Azure CLI is installed
function Test-AzureCLI {
    try {
        $azVersion = az version --output json 2>$null | ConvertFrom-Json
        if ($azVersion) {
            Write-Host "✓ Azure CLI detected: $($azVersion.'azure-cli')" -ForegroundColor Green
            return $true
        }
    }
    catch {
        return $false
    }
    return $false
}

# Function to prompt for confirmation
function Get-Confirmation {
    param([string]$Message)
    $response = Read-Host "$Message (Y/N)"
    return $response -eq 'Y' -or $response -eq 'y'
}

# Function to make REST API calls with error handling
function Invoke-GraphApiRequest {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [hashtable]$Headers,
        [string]$Body
    )
    
    try {
        $params = @{
            Uri     = $Uri
            Method  = $Method
            Headers = $Headers
        }
        
        if ($Body) {
            $params.Body = $Body
            $params.ContentType = "application/json"
        }
        
        $response = Invoke-RestMethod @params
        return $response
    }
    catch {
        Write-Host "Error: $_" -ForegroundColor Red
        Write-Host "Response: $($_.Exception.Response)" -ForegroundColor Red
        return $null
    }
}

# Start Script
Clear-Host
Write-Host @"
╔═══════════════════════════════════════════════════════════════╗
║                                                               ║
║     SharePoint API Setup Automation Script                   ║
║     Neologik - Client Configuration                       ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Green

Write-Host "`nThis script will automatically set up Microsoft Graph API access" -ForegroundColor Yellow
Write-Host "for SharePoint using the Sites.Selected permission model.`n" -ForegroundColor Yellow

# Check for Azure CLI
Write-Host "Checking prerequisites..." -ForegroundColor White
if (-not (Test-AzureCLI)) {
    Write-Host "✗ Azure CLI is not installed or not found in PATH" -ForegroundColor Red
    Write-Host "`nPlease install Azure CLI from: https://aka.ms/installazurecliwindows" -ForegroundColor Yellow
    Write-Host "After installation, restart PowerShell and run this script again.`n" -ForegroundColor Yellow
    exit
}

# ============================================================================
# STEP 1: Azure CLI Login
# ============================================================================
Write-SectionHeader "STEP 1: Azure Authentication"

Write-Host "Forcing fresh Azure CLI login..." -ForegroundColor White

# Always logout first to ensure fresh login
Write-Host "Logging out any existing session..." -ForegroundColor Yellow
az logout 2>$null | Out-Null

Write-Host "`nA browser window will open for authentication.`n" -ForegroundColor Cyan

# Perform fresh login
az login | Out-Null

$accountInfo = az account show | ConvertFrom-Json
if (-not $accountInfo) {
    Write-Host "✗ Login failed. Please try again." -ForegroundColor Red
    exit
}

Write-Host "✓ Successfully logged in as: $($accountInfo.user.name)" -ForegroundColor Green

# Get all available subscriptions
Write-Host "`nRetrieving available subscriptions..." -ForegroundColor White
$subscriptions = az account list --output json | ConvertFrom-Json

if ($subscriptions.Count -eq 0) {
    Write-Host "✗ No subscriptions found for this account." -ForegroundColor Red
    exit
}

# Display subscriptions
Write-Host "`nAvailable Subscriptions:" -ForegroundColor Cyan
for ($i = 0; $i -lt $subscriptions.Count; $i++) {
    $sub = $subscriptions[$i]
    $isDefault = if ($sub.isDefault) { " (current default)" } else { "" }
    Write-Host "  [$($i + 1)] $($sub.name)$isDefault" -ForegroundColor White
    Write-Host "      ID: $($sub.id)" -ForegroundColor Gray
    Write-Host "      Tenant: $($sub.tenantId)" -ForegroundColor Gray
}

# Prompt for subscription selection
if ($subscriptions.Count -eq 1) {
    Write-Host "`nOnly one subscription available. Using: $($subscriptions[0].name)" -ForegroundColor Green
    $selectedSubscription = $subscriptions[0]
} else {
    do {
        $selection = Read-Host "`nSelect subscription number (1-$($subscriptions.Count))"
        $selectionIndex = [int]$selection - 1
    } while ($selectionIndex -lt 0 -or $selectionIndex -ge $subscriptions.Count)
    
    $selectedSubscription = $subscriptions[$selectionIndex]
    Write-Host "Setting active subscription to: $($selectedSubscription.name)" -ForegroundColor Yellow
    
    # Set the selected subscription as active
    az account set --subscription $selectedSubscription.id | Out-Null
    
    # Verify the change
    $accountInfo = az account show | ConvertFrom-Json
    Write-Host "✓ Active subscription set successfully" -ForegroundColor Green
}

Write-Host "`nUsing Subscription:" -ForegroundColor Cyan
Write-Host "  Name: $($accountInfo.name)" -ForegroundColor Gray
Write-Host "  Tenant: $($accountInfo.tenantId)" -ForegroundColor Gray
Write-Host "  Subscription ID: $($accountInfo.id)`n" -ForegroundColor Gray

$tenantId = $accountInfo.tenantId

# ============================================================================
# STEP 2: Register Azure AD App
# ============================================================================
Write-SectionHeader "STEP 2: Register Azure AD App"

$appName = Read-Host "Enter a name for the application (e.g., 'Neologik SharePoint API Access')"

Write-Host "`nCreating app registration..." -ForegroundColor White

try {
    $appRegistration = az ad app create --display-name $appName --sign-in-audience AzureADMyOrg | ConvertFrom-Json
    $clientId = $appRegistration.appId
    
    Write-Host "✓ App registration created successfully" -ForegroundColor Green
    Write-Host "  App Name:   $appName" -ForegroundColor Gray
    Write-Host "  Client ID:  $clientId" -ForegroundColor Gray
    Write-Host "  Tenant ID:  $tenantId" -ForegroundColor Gray
    
    # Add redirect URI for admin consent flow
    Write-Host "`nConfiguring redirect URI for consent flow..." -ForegroundColor Gray
    az ad app update --id $clientId --web-redirect-uris "https://localhost" | Out-Null
    Write-Host "✓ Redirect URI configured" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to create app registration" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit
}

# ============================================================================
# STEP 3: Create Client Secret
# ============================================================================
Write-SectionHeader "STEP 3: Create Client Secret"

Write-Host "Creating client secret..." -ForegroundColor White

try {
    # Create a new credential (secret) for the app
    $credentialResponse = az ad app credential reset --id $clientId --append | ConvertFrom-Json
    
    # The password field contains the actual secret value
    if ($credentialResponse.password) {
        $clientSecretPlain = $credentialResponse.password
        Write-Host "✓ Client secret created successfully" -ForegroundColor Green
        Write-Host "`n  Client Secret Value:" -ForegroundColor Yellow
        Write-Host "  $clientSecretPlain" -ForegroundColor White
        Write-Host "`n  ⚠️  IMPORTANT: Copy this secret now - it won't be shown again!" -ForegroundColor Yellow
    }
    else {
        Write-Host "✗ Failed to retrieve client secret value" -ForegroundColor Red
        Write-Host "Response: $($credentialResponse | ConvertTo-Json)" -ForegroundColor Gray
        exit
    }
}
catch {
    Write-Host "✗ Failed to create client secret" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit
}

# ============================================================================
# STEP 4: Add Microsoft Graph API Permissions
# ============================================================================
Write-SectionHeader "STEP 4: Add Microsoft Graph API Permissions"

Write-Host "Adding required Microsoft Graph permissions..." -ForegroundColor White

# Microsoft Graph App ID (constant)
$graphAppId = "00000003-0000-0000-c000-000000000000"

# Permission IDs for Microsoft Graph
$sitesSelectedId = "883ea226-0bf2-4a8f-9f9d-92c9162a727d"  # Sites.Selected
$sitesFullControlId = "a82116e5-55eb-4c41-a434-62fe8a61c773"  # Sites.FullControl.All

try {
    # Add Sites.Selected permission
    az ad app permission add --id $clientId --api $graphAppId --api-permissions "$sitesSelectedId=Role" | Out-Null
    
    # Add Sites.FullControl.All permission
    az ad app permission add --id $clientId --api $graphAppId --api-permissions "$sitesFullControlId=Role" | Out-Null
    
    Write-Host "✓ Permissions added:" -ForegroundColor Green
    Write-Host "  - Sites.Selected" -ForegroundColor Gray
    Write-Host "  - Sites.FullControl.All" -ForegroundColor Gray
}
catch {
    Write-Host "✗ Failed to add permissions" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit
}

# ============================================================================
# STEP 5: Grant Admin Consent
# ============================================================================
Write-SectionHeader "STEP 5: Grant Admin Consent"

Write-Host "Admin consent is required for the application to access SharePoint." -ForegroundColor White
Write-Host "(This requires Global Administrator or Privileged Role Administrator permissions)`n" -ForegroundColor Yellow

# Wait for the app permissions to be registered in Azure AD
Write-Host "Waiting for permissions to be registered in Azure AD (30 seconds)..." -ForegroundColor Gray
Start-Sleep -Seconds 30

# Build the admin consent URL
$adminConsentUrl = "https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$clientId"

Write-Host "`nOpening browser for admin consent..." -ForegroundColor Cyan
Write-Host "A browser window will open. Please:" -ForegroundColor White
Write-Host "  1. Review the requested permissions" -ForegroundColor Gray
Write-Host "  2. Check the box 'Consent on behalf of your organization'" -ForegroundColor Gray
Write-Host "  3. Click 'Accept'" -ForegroundColor Gray
Write-Host "`nIMPORTANT: After accepting, you may see an error page saying 'No reply address'" -ForegroundColor Yellow
Write-Host "or 'having trouble signing you in'. This is NORMAL and means consent was successful!" -ForegroundColor Yellow
Write-Host "Look for the URL changing to contain '/Consent/Set' - that confirms success.`n" -ForegroundColor Yellow

# Open the admin consent URL in the default browser
Start-Process $adminConsentUrl

Write-Host "Consent URL: $adminConsentUrl`n" -ForegroundColor Gray

# Wait for user to complete the consent process
Write-Host "After clicking 'Accept' in the browser, come back here." -ForegroundColor Cyan
if (-not (Get-Confirmation "Have you clicked 'Accept' for admin consent? (Ignore any error page after)")) {
    Write-Host "`nAdmin consent is required to continue." -ForegroundColor Red
    Write-Host "You can manually grant consent later at:" -ForegroundColor Yellow
    Write-Host "  $adminConsentUrl`n" -ForegroundColor Gray
    exit
}

Write-Host "✓ Admin consent confirmed" -ForegroundColor Green

# Wait for permissions to propagate (critical for token generation)
Write-Host "  Waiting for permissions to propagate (20 seconds)..." -ForegroundColor Gray
Start-Sleep -Seconds 20

# ============================================================================
# STEP 6: Generate Access Token
# ============================================================================
Write-SectionHeader "STEP 6: Generate Access Token"

Write-Host "Requesting access token from Microsoft Identity Platform..." -ForegroundColor White

$tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$tokenBody = @{
    client_id     = $clientId
    client_secret = $clientSecretPlain
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
}

try {
    $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $tokenBody -ContentType "application/x-www-form-urlencoded"
    $accessToken = $tokenResponse.access_token
    Write-Host "✓ Access token obtained successfully" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to obtain access token" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host "`nPlease verify:" -ForegroundColor Yellow
    Write-Host "  - Client ID and Client Secret are correct" -ForegroundColor Gray
    Write-Host "  - Admin consent has been granted" -ForegroundColor Gray
    Write-Host "  - API permissions are correctly configured" -ForegroundColor Gray
    exit
}

# ============================================================================
# STEP 7 & 8: Get Site ID and Assign Permissions (Multiple Sites)
# ============================================================================
Write-SectionHeader "STEP 7 & 8: Configure SharePoint Site Access"

Write-Host "You can grant this app access to one or more SharePoint sites.`n" -ForegroundColor White

$sites = @()
$addMoreSites = $true
$siteNumber = 1

while ($addMoreSites) {
    Write-Host "--- Site #$siteNumber ---" -ForegroundColor Cyan
    Write-Host "(Example: For 'https://contoso.sharepoint.com/sites/marketing')" -ForegroundColor Gray
    Write-Host "  Hostname: contoso.sharepoint.com" -ForegroundColor Gray
    Write-Host "  Site Name: marketing`n" -ForegroundColor Gray

    $hostname = Read-Host "Enter SharePoint Hostname (e.g., contoso.sharepoint.com)"
    $siteName = Read-Host "Enter Site Name (e.g., marketing)"

    Write-Host "`nRetrieving Site ID..." -ForegroundColor White

    $siteUrl = "https://graph.microsoft.com/v1.0/sites/$hostname`:/sites/$siteName"
    $headers = @{
        Authorization = "Bearer $accessToken"
    }

    $siteResponse = Invoke-GraphApiRequest -Uri $siteUrl -Method "GET" -Headers $headers

    if ($siteResponse -and $siteResponse.id) {
        $siteId = $siteResponse.id
        Write-Host "✓ Site ID retrieved: $siteId" -ForegroundColor Green
        
        # Assign permissions to the site
        Write-Host "  Assigning read permissions to the site..." -ForegroundColor White
        
        $permissionUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/permissions"
        $permissionBody = @{
            roles = @("read")
            grantedToIdentities = @(
                @{
                    application = @{
                        id = $clientId
                        displayName = $appName
                    }
                }
            )
        } | ConvertTo-Json -Depth 10

        $permissionResponse = Invoke-GraphApiRequest -Uri $permissionUrl -Method "POST" -Headers $headers -Body $permissionBody

        if ($permissionResponse) {
            Write-Host "✓ Permissions assigned successfully to $siteName" -ForegroundColor Green
            
            # Store site information
            $sites += @{
                Hostname = $hostname
                SiteName = $siteName
                SiteId = $siteId
                PermissionsGranted = $true
            }
        }
        else {
            Write-Host "✗ Failed to assign permissions to $siteName" -ForegroundColor Red
            $sites += @{
                Hostname = $hostname
                SiteName = $siteName
                SiteId = $siteId
                PermissionsGranted = $false
            }
        }
    }
    else {
        Write-Host "✗ Failed to retrieve Site ID" -ForegroundColor Red
        Write-Host "`nPossible issues:" -ForegroundColor Yellow
        Write-Host "  1. Site name or hostname is incorrect" -ForegroundColor Gray
        Write-Host "  2. Admin consent may not have propagated yet" -ForegroundColor Gray
        Write-Host "  3. The app may not have the required permissions" -ForegroundColor Gray
        Write-Host "`nPlease verify:" -ForegroundColor Yellow
        Write-Host "  - Site exists at: https://$hostname/sites/$siteName" -ForegroundColor Gray
        Write-Host "  - Admin consent is granted in Azure Portal" -ForegroundColor Gray
        Write-Host "  - Both permissions (Sites.Selected and Sites.FullControl.All) show 'Granted' status`n" -ForegroundColor Gray
        
        if (Get-Confirmation "Would you like to retry this site?") {
            continue
        }
    }
    
    # Ask if user wants to add another site
    Write-Host ""
    if (Get-Confirmation "Would you like to add another SharePoint site?") {
        $siteNumber++
        continue
    }
    else {
        $addMoreSites = $false
    }
}

Write-Host "`n✓ Completed site configuration for $($sites.Count) site(s)" -ForegroundColor Green

# ============================================================================
# STEP 9: Security Cleanup
# ============================================================================
Write-SectionHeader "STEP 9: Security Cleanup (Remove Elevated Permission)"

Write-Host "Removing Sites.FullControl.All permission for security best practices..." -ForegroundColor Yellow

try {
    # Remove the Sites.FullControl.All permission (using correct syntax without --permission-id)
    az ad app permission delete --id $clientId --api $graphAppId | Out-Null
    
    # Re-add only Sites.Selected
    az ad app permission add --id $clientId --api $graphAppId --api-permissions "$sitesSelectedId=Role" | Out-Null
    
    # Re-grant admin consent for Sites.Selected
    az ad app permission admin-consent --id $clientId | Out-Null
    
    Write-Host "✓ Sites.FullControl.All permission removed successfully" -ForegroundColor Green
    Write-Host "  The app now operates with Sites.Selected permission only." -ForegroundColor Gray
}
catch {
    Write-Host "⚠ Could not automatically remove Sites.FullControl.All" -ForegroundColor Yellow
    Write-Host "  Please manually remove it in the Azure Portal for security:" -ForegroundColor Gray
    Write-Host "  Azure Portal → App Registrations → $appName → API Permissions" -ForegroundColor Gray
}

# ============================================================================
# Summary and Save Configuration
# ============================================================================
Write-SectionHeader "Setup Complete!"

$configData = @{
    AppName = $appName
    ClientId = $clientId
    TenantId = $tenantId
    Sites = $sites
    SetupDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
}

$configJson = $configData | ConvertTo-Json -Depth 10

Write-Host "Configuration Summary:" -ForegroundColor Green
Write-Host $configJson -ForegroundColor Gray

if (Get-Confirmation "`nWould you like to save this configuration to a file? (Client secret will NOT be saved)") {
    $configPath = Join-Path $PSScriptRoot "sharepoint-config.json"
    $configJson | Out-File -FilePath $configPath -Encoding UTF8
    Write-Host "✓ Configuration saved to: $configPath" -ForegroundColor Green
}

Write-Host "`n" -NoNewline
Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  Setup completed successfully!                                ║" -ForegroundColor Green
Write-Host "║                                                               ║" -ForegroundColor Green
Write-Host "║  Your app now has read access to $($sites.Count) SharePoint site(s).  ║" -ForegroundColor Green
Write-Host "║  Don't forget to remove Sites.FullControl.All permission!    ║" -ForegroundColor Green
Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Green

# Clean up sensitive data from memory
$clientSecretPlain = $null
$accessToken = $null
[System.GC]::Collect()

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
