<#
.SYNOPSIS
    Manages SharePoint permissions for an existing Azure Service Principal.

.DESCRIPTION
    This script takes an existing Service Principal, ensures it has the required permissions,
    creates a temporary client secret, lists all SharePoint sites, allows selection of sites
    to grant permissions to, and then cleans up by removing elevated permissions and the temporary secret.

.NOTES
    Prerequisites:
    - Azure CLI installed (https://aka.ms/installazurecliwindows)
    - Azure AD admin access
    - PowerShell 5.1 or later
    - Internet connection
    - Existing Service Principal in Azure AD
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
        if ($_.Exception.Response) {
            Write-Host "Response: $($_.Exception.Response)" -ForegroundColor Red
        }
        return $null
    }
}

# Start Script
Clear-Host
Write-Host @"
╔═══════════════════════════════════════════════════════════════╗
║                                                               ║
║     SharePoint Permission Manager                            ║
║     Service Principal Configuration                          ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Green

Write-Host "`nThis script will manage SharePoint permissions for an existing" -ForegroundColor Yellow
Write-Host "Service Principal using Microsoft Graph API.`n" -ForegroundColor Yellow

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

Write-Host "Checking Azure CLI login status..." -ForegroundColor White

try {
    $accountInfo = az account show 2>$null | ConvertFrom-Json
    if ($accountInfo) {
        Write-Host "✓ Already logged in as: $($accountInfo.user.name)" -ForegroundColor Green
        Write-Host "  Tenant: $($accountInfo.tenantId)" -ForegroundColor Gray
        Write-Host "  Subscription: $($accountInfo.name)`n" -ForegroundColor Gray
        
        if (-not (Get-Confirmation "Continue with this account?")) {
            Write-Host "Logging out..." -ForegroundColor Yellow
            az logout | Out-Null
            Write-Host "Please run the script again to login with a different account." -ForegroundColor Yellow
            exit
        }
    }
}
catch {
    Write-Host "Not currently logged in. Starting authentication..." -ForegroundColor Yellow
    Write-Host "`nA browser window will open for authentication.`n" -ForegroundColor Cyan
    
    az login | Out-Null
    
    $accountInfo = az account show | ConvertFrom-Json
    if (-not $accountInfo) {
        Write-Host "✗ Login failed. Please try again." -ForegroundColor Red
        exit
    }
    
    Write-Host "✓ Successfully logged in as: $($accountInfo.user.name)" -ForegroundColor Green
}

$tenantId = $accountInfo.tenantId

# ============================================================================
# STEP 2: Get or Create Service Principal
# ============================================================================
Write-SectionHeader "STEP 2: Service Principal Setup"

Write-Host "Would you like to:" -ForegroundColor White
Write-Host "  1. Use an existing Service Principal" -ForegroundColor Cyan
Write-Host "  2. Create a new Service Principal" -ForegroundColor Cyan
Write-Host ""

do {
    $spChoice = Read-Host "Enter your choice (1 or 2)"
} while ($spChoice -ne '1' -and $spChoice -ne '2')

if ($spChoice -eq '2') {
    Write-Host "`nHow would you like to provide the Service Principal name?" -ForegroundColor White
    Write-Host "  1. Derive from Resource Group name (e.g. rg-neo-abc-dev-uks-01)" -ForegroundColor Cyan
    Write-Host "  2. Enter a short name (e.g. neosecdev01)  — will be prefixed automatically" -ForegroundColor Cyan
    Write-Host "  3. Enter the full name as plain text" -ForegroundColor Cyan
    Write-Host ""

    do {
        $nameChoice = Read-Host "Enter your choice (1, 2 or 3)"
    } while ($nameChoice -ne '1' -and $nameChoice -ne '2' -and $nameChoice -ne '3')

    if ($nameChoice -eq '1') {
        # Derive SP name from Resource Group name
        $rgName = Read-Host "`nEnter the Resource Group name (e.g. rg-neo-abc-dev-uks-01)"

        # Strip leading 'rg-' prefix
        $rgStripped = $rgName -replace '^rg-', ''

        # Split into parts and remove the 2nd-last segment (location value)
        $parts = $rgStripped -split '-'
        if ($parts.Count -ge 2) {
            $partsList = [System.Collections.Generic.List[string]]$parts
            $partsList.RemoveAt($partsList.Count - 2)
            $condensed = ($partsList | ForEach-Object { $_.ToLower() }) -join ''
        }
        else {
            $condensed = $rgStripped.ToLower() -replace '-', ''
        }

        $appName = "Neologik SharePoint Access - $condensed"
    }
    elseif ($nameChoice -eq '2') {
        $shortName = Read-Host "`nEnter the short name (e.g. neosecdev01)"
        $appName = "Neologik SharePoint Access - $($shortName.Trim())"
    }
    else {
        $appName = Read-Host "`nEnter the full Service Principal name"
        $appName = $appName.Trim()
    }

    Write-Host "`nService Principal name: " -ForegroundColor White -NoNewline
    Write-Host $appName -ForegroundColor Cyan

    if (-not (Get-Confirmation "Create Service Principal with this name?")) {
        Write-Host "Exiting. Please re-run and enter a different name." -ForegroundColor Yellow
        exit
    }

    Write-Host "`nCreating App Registration..." -ForegroundColor White
    try {
        $newApp = az ad app create --display-name $appName --output json 2>$null | ConvertFrom-Json
        if (-not $newApp) {
            Write-Host "✗ Failed to create App Registration" -ForegroundColor Red
            exit
        }
        $clientId = $newApp.appId
        Write-Host "✓ App Registration created: $appName" -ForegroundColor Green
        Write-Host "  Client ID: $clientId" -ForegroundColor Gray

        Write-Host "Creating Service Principal..." -ForegroundColor White
        $newSp = az ad sp create --id $clientId --output json 2>$null | ConvertFrom-Json
        if (-not $newSp) {
            Write-Host "✗ Failed to create Service Principal" -ForegroundColor Red
            exit
        }
        Write-Host "✓ Service Principal created" -ForegroundColor Green

        Write-Host "  Waiting for Service Principal to propagate (15 seconds)..." -ForegroundColor Gray
        Start-Sleep -Seconds 15
    }
    catch {
        Write-Host "✗ Failed to create Service Principal" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        exit
    }
}
else {
    $servicePrincipalId = Read-Host "`nEnter the Service Principal Client ID (Application ID)"

    Write-Host "`nValidating Service Principal..." -ForegroundColor White

    try {
        $spInfo = az ad sp show --id $servicePrincipalId 2>$null | ConvertFrom-Json
        if (-not $spInfo) {
            # Try to get app registration instead
            $appInfo = az ad app show --id $servicePrincipalId 2>$null | ConvertFrom-Json
            if ($appInfo) {
                $clientId = $appInfo.appId
                $appName = $appInfo.displayName
                Write-Host "✓ App Registration found: $appName" -ForegroundColor Green
                Write-Host "  Client ID: $clientId" -ForegroundColor Gray
            }
            else {
                Write-Host "✗ Service Principal or App Registration not found" -ForegroundColor Red
                exit
            }
        }
        else {
            $clientId = $spInfo.appId
            $appName = $spInfo.displayName
            Write-Host "✓ Service Principal found: $appName" -ForegroundColor Green
            Write-Host "  Client ID: $clientId" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "✗ Failed to retrieve Service Principal" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        exit
    }
}

# Use Client ID as the application identifier for permissions
$appIdentifier = $clientId

# ============================================================================
# STEP 3: Check and Add Required Permissions
# ============================================================================
Write-SectionHeader "STEP 3: Check and Add Required Permissions"

Write-Host "Checking current permissions..." -ForegroundColor White

# Microsoft Graph App ID (constant)
$graphAppId = "00000003-0000-0000-c000-000000000000"

# Permission IDs for Microsoft Graph
$sitesSelectedId = "883ea226-0bf2-4a8f-9f9d-92c9162a727d"  # Sites.Selected
$sitesFullControlId = "a82116e5-55eb-4c41-a434-62fe8a61c773"  # Sites.FullControl.All

try {
    # Get current permissions
    $currentPermissions = az ad app permission list --id $clientId | ConvertFrom-Json
    
    $hasSitesSelected = $false
    $hasSitesFullControl = $false
    
    foreach ($resource in $currentPermissions) {
        if ($resource.resourceAppId -eq $graphAppId) {
            foreach ($permission in $resource.resourceAccess) {
                if ($permission.id -eq $sitesSelectedId) {
                    $hasSitesSelected = $true
                    Write-Host "✓ Sites.Selected permission found" -ForegroundColor Green
                }
                if ($permission.id -eq $sitesFullControlId) {
                    $hasSitesFullControl = $true
                    Write-Host "✓ Sites.FullControl.All permission found" -ForegroundColor Green
                }
            }
        }
    }
    
    # Add missing permissions
    $permissionsAdded = $false

    if (-not $hasSitesSelected) {
        Write-Host "Adding Sites.Selected permission..." -ForegroundColor Yellow
        az ad app permission add --id $clientId --api $graphAppId --api-permissions "$sitesSelectedId=Role" 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-Host "✗ Failed to add Sites.Selected permission" -ForegroundColor Red
            exit
        }
        Write-Host "✓ Sites.Selected permission added" -ForegroundColor Green
        $permissionsAdded = $true
    }

    if (-not $hasSitesFullControl) {
        Write-Host "Adding Sites.FullControl.All permission..." -ForegroundColor Yellow
        az ad app permission add --id $clientId --api $graphAppId --api-permissions "$sitesFullControlId=Role" 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-Host "✗ Failed to add Sites.FullControl.All permission" -ForegroundColor Red
            exit
        }
        Write-Host "✓ Sites.FullControl.All permission added" -ForegroundColor Green
        $permissionsAdded = $true
    }

    if ($permissionsAdded) {
        Write-Host "`nWaiting for permissions to register (20 seconds)..." -ForegroundColor Gray
        Start-Sleep -Seconds 20
    }
}
catch {
    Write-Host "✗ Failed to check/add permissions" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit
}

# ============================================================================
# STEP 4: Grant Admin Consent
# ============================================================================
Write-SectionHeader "STEP 4: Grant Admin Consent"

Write-Host "Granting admin consent for permissions..." -ForegroundColor White

$consentOutput = az ad app permission admin-consent --id $clientId 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Admin consent granted successfully" -ForegroundColor Green
    Write-Host "  Waiting for consent to propagate (15 seconds)..." -ForegroundColor Gray
    Start-Sleep -Seconds 15
}
else {
    Write-Host "⚠ Automated admin consent failed:" -ForegroundColor Yellow
    Write-Host "  $consentOutput" -ForegroundColor Red
    Write-Host "`n  Opening browser for manual consent..." -ForegroundColor Yellow

    $adminConsentUrl = "https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$clientId"
    Start-Process $adminConsentUrl

    Write-Host "`nPlease complete admin consent in the browser window." -ForegroundColor Cyan
    Write-Host "After clicking 'Accept', come back here." -ForegroundColor Cyan

    if (-not (Get-Confirmation "Have you completed admin consent?")) {
        Write-Host "✗ Admin consent is required to continue." -ForegroundColor Red
        exit
    }

    Write-Host "✓ Admin consent confirmed" -ForegroundColor Green
    Start-Sleep -Seconds 10
}

# ============================================================================
# STEP 5: Create Client Secret
# ============================================================================
Write-SectionHeader "STEP 5: Create Temporary Client Secret"

Write-Host "Creating client secret named 'test'..." -ForegroundColor White

try {
    $secretResponse = az ad app credential reset --id $clientId --append --display-name "test" --output json 2>$null | ConvertFrom-Json

    if ($secretResponse -and $secretResponse.password) {
        $clientSecretPlain = $secretResponse.password
        Write-Host "✓ Client secret 'test' created successfully" -ForegroundColor Green

        # az ad app credential reset does not return keyId — list credentials to retrieve it
        $allCreds = az ad app credential list --id $clientId --output json 2>$null | ConvertFrom-Json
        $testCred = $allCreds | Where-Object { $_.displayName -eq "test" } | Sort-Object endDateTime -Descending | Select-Object -First 1
        $secretKeyId = $testCred.keyId
        Write-Host "  Key ID: $secretKeyId" -ForegroundColor Gray

        # Wait for the secret to propagate in Azure AD before using it
        Write-Host "  Waiting for secret to propagate (20 seconds)..." -ForegroundColor Gray
        Start-Sleep -Seconds 20
    }
    else {
        Write-Host "✗ Failed to create client secret" -ForegroundColor Red
        exit
    }
}
catch {
    Write-Host "✗ Failed to create client secret" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit
}

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
    Write-Host "`nPlease verify admin consent was granted properly." -ForegroundColor Yellow
    exit
}

# ============================================================================
# STEP 7: List All Available SharePoint Sites
# ============================================================================
Write-SectionHeader "STEP 7: List All Available SharePoint Sites"

Write-Host "Retrieving all SharePoint sites..." -ForegroundColor White

$headers = @{
    Authorization = "Bearer $accessToken"
}

$sitesListUrl = "https://graph.microsoft.com/v1.0/sites?search=*"

try {
    $sitesResponse = Invoke-GraphApiRequest -Uri $sitesListUrl -Method "GET" -Headers $headers
    
    if ($sitesResponse -and $sitesResponse.value) {
        $availableSites = @()
        
        foreach ($site in $sitesResponse.value) {
            $availableSites += [PSCustomObject]@{
                Number      = $availableSites.Count + 1
                Name        = $site.displayName
                WebUrl      = $site.webUrl
                SiteId      = $site.id
            }
        }
        
        Write-Host "`n✓ Found $($availableSites.Count) SharePoint sites:" -ForegroundColor Green
        Write-Host ""
        
        foreach ($site in $availableSites) {
            Write-Host "$($site.Number). " -ForegroundColor Cyan -NoNewline
            Write-Host "$($site.Name)" -ForegroundColor White
            Write-Host "   URL: $($site.WebUrl)" -ForegroundColor Gray
            Write-Host "   ID:  $($site.SiteId)" -ForegroundColor DarkGray
        }
        
        Write-Host ""
    }
    else {
        Write-Host "✗ No sites found or failed to retrieve sites" -ForegroundColor Red
        Write-Host "This could mean:" -ForegroundColor Yellow
        Write-Host "  - No SharePoint sites exist in the tenant" -ForegroundColor Gray
        Write-Host "  - Permissions haven't fully propagated yet" -ForegroundColor Gray
        exit
    }
}
catch {
    Write-Host "✗ Failed to retrieve SharePoint sites" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit
}

# ============================================================================
# STEP 8: Select Sites and Assign Permissions
# ============================================================================
Write-SectionHeader "STEP 8: Assign Permissions to Sites"

$selectedSites = @()
$continueSelection = $true

while ($continueSelection) {
    $siteSelection = Read-Host "`nWhat site do you want to add permission to? (Enter number)"
    
    try {
        $siteNumber = [int]$siteSelection
        
        if ($siteNumber -ge 1 -and $siteNumber -le $availableSites.Count) {
            $selectedSite = $availableSites[$siteNumber - 1]
            
            Write-Host "`nAdding permissions to: $($selectedSite.Name)" -ForegroundColor White
            
            $permissionUrl = "https://graph.microsoft.com/v1.0/sites/$($selectedSite.SiteId)/permissions"
            $permissionBody = @{
                roles = @("write")
                grantedToIdentities = @(
                    @{
                        application = @{
                            id = $clientId
                            displayName = $clientId
                        }
                    }
                )
            } | ConvertTo-Json -Depth 10
            
            $permissionResponse = Invoke-GraphApiRequest -Uri $permissionUrl -Method "POST" -Headers $headers -Body $permissionBody
            
            if ($permissionResponse) {
                Write-Host "✓ Permissions granted successfully to: $($selectedSite.Name)" -ForegroundColor Green
                $selectedSites += $selectedSite
            }
            else {
                Write-Host "✗ Failed to grant permissions to: $($selectedSite.Name)" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Invalid selection. Please enter a number between 1 and $($availableSites.Count)" -ForegroundColor Red
            continue
        }
    }
    catch {
        Write-Host "Invalid input. Please enter a valid number." -ForegroundColor Red
        continue
    }
    
    if (Get-Confirmation "`nDo you want to add another site?") {
        $continueSelection = $true
    }
    else {
        $continueSelection = $false
    }
}

Write-Host "`n✓ Permissions granted to $($selectedSites.Count) site(s)" -ForegroundColor Green

# ============================================================================
# STEP 9: Security Cleanup - Remove FullControl Permission
# ============================================================================
Write-SectionHeader "STEP 9: Security Cleanup - Remove FullControl Permission"

Write-Host "Removing Sites.FullControl.All permission..." -ForegroundColor Yellow

try {
    # Revoke the admin-consented role assignment first (the actual active grant)
    $spObjectId = (az ad sp show --id $clientId --output json 2>$null | ConvertFrom-Json).id
    $graphSpId  = (az ad sp show --id $graphAppId --output json 2>$null | ConvertFrom-Json).id

    if ($spObjectId -and $graphSpId) {
        # Use az rest (logged-in admin user token) — the SP's own token lacks AppRoleAssignment.ReadWrite.All
        $roleAssignmentsJson = az rest --method GET `
            --url "https://graph.microsoft.com/v1.0/servicePrincipals/$spObjectId/appRoleAssignments" `
            --output json 2>$null | ConvertFrom-Json

        $fullControlAssignment = $roleAssignmentsJson.value | Where-Object {
            $_.appRoleId -eq $sitesFullControlId -and $_.resourceId -eq $graphSpId
        }

        if ($fullControlAssignment) {
            az rest --method DELETE `
                --url "https://graph.microsoft.com/v1.0/servicePrincipals/$spObjectId/appRoleAssignments/$($fullControlAssignment.id)" `
                2>&1 | Out-Null
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Sites.FullControl.All admin consent revoked" -ForegroundColor Green
            }
            else {
                Write-Host "⚠ Could not revoke role assignment - please remove manually in Azure Portal" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "  No active role assignment found for Sites.FullControl.All" -ForegroundColor Gray
        }
    }

    # Delete all Graph permissions
    az ad app permission delete --id $clientId --api $graphAppId | Out-Null

    # Re-add Sites.Selected only
    az ad app permission add --id $clientId --api $graphAppId --api-permissions "$sitesSelectedId=Role" | Out-Null

    # Re-grant admin consent
    az ad app permission admin-consent --id $clientId | Out-Null

    Write-Host "✓ Sites.FullControl.All permission removed" -ForegroundColor Green
    Write-Host "  App now has only Sites.Selected permission" -ForegroundColor Gray
}
catch {
    Write-Host "⚠ Could not automatically remove Sites.FullControl.All" -ForegroundColor Yellow
    Write-Host "  Please manually remove it in Azure Portal" -ForegroundColor Gray
}

# ============================================================================
# STEP 10: Delete Client Secret
# ============================================================================
Write-SectionHeader "STEP 10: Delete Temporary Client Secret"

Write-Host "Deleting client secret 'test'..." -ForegroundColor Yellow

try {
    az ad app credential delete --id $clientId --key-id $secretKeyId | Out-Null
    Write-Host "✓ Client secret 'test' deleted successfully" -ForegroundColor Green
}
catch {
    Write-Host "⚠ Could not automatically delete client secret" -ForegroundColor Yellow
    Write-Host "  Please manually delete it in Azure Portal" -ForegroundColor Gray
    Write-Host "  Key ID: $secretKeyId" -ForegroundColor Gray
}

# ============================================================================
# Summary
# ============================================================================
Write-SectionHeader "Setup Complete!"

Write-Host "Summary:" -ForegroundColor Green
Write-Host "  Service Principal: $appName" -ForegroundColor Gray
Write-Host "  Client ID: $clientId" -ForegroundColor Gray
Write-Host "  Sites with permissions: $($selectedSites.Count)" -ForegroundColor Gray
Write-Host ""

foreach ($site in $selectedSites) {
    Write-Host "  - $($site.Name)" -ForegroundColor Gray
}

Write-Host "`n" -NoNewline
Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  Configuration completed successfully!                       ║" -ForegroundColor Green
Write-Host "║                                                               ║" -ForegroundColor Green
Write-Host "║  - Permissions verified and added                            ║" -ForegroundColor Green
Write-Host "║  - Sites configured: $($selectedSites.Count)                                        ║" -ForegroundColor Green
Write-Host "║  - FullControl permission removed                            ║" -ForegroundColor Green
Write-Host "║  - Temporary client secret deleted                           ║" -ForegroundColor Green
Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Green

# Clean up sensitive data from memory
$clientSecretPlain = $null
$accessToken = $null
[System.GC]::Collect()

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")