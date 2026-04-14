<#
.SYNOPSIS
    Grants SharePoint site permissions to a User-Assigned Managed Identity (UAMI).

.DESCRIPTION
    UAMIs surface in Azure AD as Enterprise Applications (service principals of type
    ManagedIdentity). They have no App Registration, no client secrets, and do not
    support the client-credentials token flow. This script therefore uses the
    logged-in Azure CLI admin user's token to call Microsoft Graph, then grants
    direct site-level permissions to the UAMI's service principal.

.NOTES
    Prerequisites:
    - Azure CLI installed (https://aka.ms/installazurecliwindows)
    - Logged-in account must be a SharePoint Administrator or Global Administrator
    - PowerShell 5.1 or later
    - The target UAMI must already exist in Azure
#>

[CmdletBinding()]
param()

# ============================================================================
# Helper Functions
# ============================================================================

function Write-SectionHeader {
    param([string]$Title)
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}

function Test-AzureCLI {
    try {
        $azVersion = az version --output json 2>$null | ConvertFrom-Json
        if ($azVersion) {
            Write-Host "✓ Azure CLI detected: $($azVersion.'azure-cli')" -ForegroundColor Green
            return $true
        }
    }
    catch { }
    return $false
}

function Get-Confirmation {
    param([string]$Message)
    $response = Read-Host "$Message (Y/N)"
    return $response -eq 'Y' -or $response -eq 'y'
}

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
            $params.Body        = $Body
            $params.ContentType = "application/json"
        }

        return Invoke-RestMethod @params
    }
    catch {
        Write-Host "API Error: $_" -ForegroundColor Red
        if ($_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                Write-Host "Response body: $($reader.ReadToEnd())" -ForegroundColor Red
            }
            catch { }
        }
        return $null
    }
}

# ============================================================================
# Banner
# ============================================================================
Clear-Host
Write-Host @"
╔═══════════════════════════════════════════════════════════════╗
║                                                               ║
║     SharePoint Permission Manager                            ║
║     User-Assigned Managed Identity (UAMI) Configuration      ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Green

Write-Host "`nThis script grants SharePoint site permissions to a UAMI." -ForegroundColor Yellow
Write-Host "It uses your Azure CLI admin token — no client secret is required.`n" -ForegroundColor Yellow

# ============================================================================
# STEP 1: Prerequisites
# ============================================================================
Write-SectionHeader "STEP 1: Check Prerequisites"

if (-not (Test-AzureCLI)) {
    Write-Host "✗ Azure CLI is not installed or not found in PATH" -ForegroundColor Red
    Write-Host "`nPlease install Azure CLI from: https://aka.ms/installazurecliwindows" -ForegroundColor Yellow
    exit
}

# ============================================================================
# STEP 2: Azure CLI Login
# ============================================================================
Write-SectionHeader "STEP 2: Azure Authentication"

Write-Host "Checking Azure CLI login status..." -ForegroundColor White

try {
    $accountInfo = az account show 2>$null | ConvertFrom-Json
}
catch {
    $accountInfo = $null
}

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
else {
    Write-Host "Not currently logged in. Starting authentication..." -ForegroundColor Yellow
    Write-Host "`nA browser window will open for authentication.`n" -ForegroundColor Cyan

    az login | Out-Null

    $accountInfo = az account show 2>$null | ConvertFrom-Json
    if (-not $accountInfo) {
        Write-Host "✗ Login failed. Please try again." -ForegroundColor Red
        exit
    }

    Write-Host "✓ Successfully logged in as: $($accountInfo.user.name)" -ForegroundColor Green
}

$tenantId = $accountInfo.tenantId

# ============================================================================
# STEP 3: Look Up the UAMI
# ============================================================================
Write-SectionHeader "STEP 3: Locate the User-Assigned Managed Identity"

Write-Host "How would you like to identify the UAMI?" -ForegroundColor White
Write-Host "  1. By Client ID (appId / Application ID)" -ForegroundColor Cyan
Write-Host "  2. By display name" -ForegroundColor Cyan
Write-Host "  3. By Azure resource ID" -ForegroundColor Cyan
Write-Host ""

do {
    $lookupChoice = Read-Host "Enter your choice (1, 2 or 3)"
} while ($lookupChoice -ne '1' -and $lookupChoice -ne '2' -and $lookupChoice -ne '3')

$uamiSp = $null

switch ($lookupChoice) {
    '1' {
        $uamiClientId = Read-Host "`nEnter the UAMI Client ID (Application ID)"
        Write-Host "`nLooking up service principal by Client ID..." -ForegroundColor White

        $uamiSp = az ad sp show --id $uamiClientId --output json 2>$null | ConvertFrom-Json

        if (-not $uamiSp) {
            Write-Host "✗ No service principal found with Client ID: $uamiClientId" -ForegroundColor Red
            exit
        }
    }
    '2' {
        $uamiName = Read-Host "`nEnter the UAMI display name"
        Write-Host "`nSearching for managed identity service principals named '$uamiName'..." -ForegroundColor White

        # Filter to ManagedIdentity type and match on display name
        $spList = az ad sp list --filter "servicePrincipalType eq 'ManagedIdentity' and displayName eq '$uamiName'" --output json 2>$null | ConvertFrom-Json

        if (-not $spList -or $spList.Count -eq 0) {
            # Fallback: search all SPs by display name (slower but catches edge cases)
            Write-Host "  Exact match not found — doing a broader display-name search..." -ForegroundColor Gray
            $spList = az ad sp list --display-name $uamiName --output json 2>$null | ConvertFrom-Json
            $spList = $spList | Where-Object { $_.servicePrincipalType -eq 'ManagedIdentity' }
        }

        if (-not $spList -or $spList.Count -eq 0) {
            Write-Host "✗ No managed identity found with display name: $uamiName" -ForegroundColor Red
            Write-Host "  Tip: try option 1 (Client ID) or option 3 (resource ID) instead." -ForegroundColor Yellow
            exit
        }

        if ($spList.Count -gt 1) {
            Write-Host "`nMultiple matches found — please select one:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $spList.Count; $i++) {
                Write-Host "  $($i + 1). $($spList[$i].displayName)  [appId: $($spList[$i].appId)]" -ForegroundColor Cyan
            }
            do {
                $pick = Read-Host "`nEnter number"
                $pickInt = [int]$pick
            } while ($pickInt -lt 1 -or $pickInt -gt $spList.Count)

            $uamiSp = $spList[$pickInt - 1]
        }
        else {
            $uamiSp = $spList[0]
        }
    }
    '3' {
        $uamiResourceId = Read-Host "`nEnter the UAMI Azure resource ID"
        Write-Host "`nRetrieving UAMI details via az identity show..." -ForegroundColor White

        $identityInfo = az identity show --ids $uamiResourceId --output json 2>$null | ConvertFrom-Json

        if (-not $identityInfo) {
            Write-Host "✗ Could not retrieve UAMI with that resource ID" -ForegroundColor Red
            exit
        }

        Write-Host "  Found: $($identityInfo.name)  [clientId: $($identityInfo.clientId)]" -ForegroundColor Gray

        # Look up the corresponding service principal using the clientId
        $uamiSp = az ad sp show --id $identityInfo.clientId --output json 2>$null | ConvertFrom-Json

        if (-not $uamiSp) {
            Write-Host "✗ Could not find the service principal for this UAMI (clientId: $($identityInfo.clientId))" -ForegroundColor Red
            exit
        }
    }
}

# Validate it is actually a Managed Identity SP
if ($uamiSp.servicePrincipalType -ne 'ManagedIdentity') {
    Write-Host "`n⚠  Warning: the resolved service principal has type '$($uamiSp.servicePrincipalType)', not 'ManagedIdentity'." -ForegroundColor Yellow
    if (-not (Get-Confirmation "Continue anyway?")) {
        exit
    }
}

$uamiAppId   = $uamiSp.appId          # This is the 'application id' used in the permission grant
$uamiObjId   = $uamiSp.id             # Object ID of the SP
$uamiDisplay = $uamiSp.displayName

Write-Host "`n✓ UAMI service principal located:" -ForegroundColor Green
Write-Host "  Display Name : $uamiDisplay" -ForegroundColor White
Write-Host "  Client ID    : $uamiAppId" -ForegroundColor White
Write-Host "  Object ID    : $uamiObjId" -ForegroundColor White

# ============================================================================
# STEP 4: Obtain Admin Access Token (delegated, via Azure CLI)
# ============================================================================
Write-SectionHeader "STEP 4: Obtain Admin Access Token"

Write-Host "Requesting Microsoft Graph token for the logged-in admin account..." -ForegroundColor White
Write-Host "(Your account must be a SharePoint Administrator or Global Administrator)" -ForegroundColor Gray

try {
    $tokenJson    = az account get-access-token --resource https://graph.microsoft.com --output json 2>$null | ConvertFrom-Json
    $accessToken  = $tokenJson.accessToken

    if (-not $accessToken) {
        throw "Empty token returned"
    }

    Write-Host "✓ Access token obtained" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to obtain access token" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host "`nMake sure the logged-in account has SharePoint Admin or Global Admin rights." -ForegroundColor Yellow
    exit
}

$headers = @{ Authorization = "Bearer $accessToken" }

# ============================================================================
# STEP 5: Grant Sites.Selected Graph API Permission to the UAMI
# ============================================================================
Write-SectionHeader "STEP 5: Assign Sites.Selected Permission to UAMI"

Write-Host "Looking up the Microsoft Graph service principal..." -ForegroundColor White

$graphSp = Invoke-GraphApiRequest `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'" `
    -Method "GET" `
    -Headers $headers

if (-not $graphSp -or $graphSp.value.Count -eq 0) {
    Write-Host "✗ Could not find the Microsoft Graph service principal in this tenant." -ForegroundColor Red
    exit
}

$graphSpId = $graphSp.value[0].id
Write-Host "  Graph SP Object ID: $graphSpId" -ForegroundColor Gray

# Sites.Selected app role ID on Microsoft Graph (constant across all tenants)
$sitesSelectedRoleId = "883ea226-0bf2-4a8f-9f9d-92c7314d055b"

# Check if already assigned
Write-Host "Checking existing app role assignments on the UAMI..." -ForegroundColor White

$existingAssignments = Invoke-GraphApiRequest `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$uamiObjId/appRoleAssignments" `
    -Method "GET" `
    -Headers $headers

$alreadyAssigned = $existingAssignments.value | Where-Object {
    $_.appRoleId -eq $sitesSelectedRoleId -and $_.resourceId -eq $graphSpId
}

if ($alreadyAssigned) {
    Write-Host "✓ Sites.Selected is already assigned to this UAMI — skipping." -ForegroundColor Green
}
else {
    Write-Host "Assigning Sites.Selected to UAMI '$uamiDisplay'..." -ForegroundColor White

    $assignBody = @{
        principalId = $uamiObjId
        resourceId  = $graphSpId
        appRoleId   = $sitesSelectedRoleId
    } | ConvertTo-Json

    $assignResponse = Invoke-GraphApiRequest `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$uamiObjId/appRoleAssignments" `
        -Method "POST" `
        -Headers $headers `
        -Body $assignBody

    if ($assignResponse) {
        Write-Host "✓ Sites.Selected granted to '$uamiDisplay'" -ForegroundColor Green
    }
    else {
        Write-Host "✗ Failed to assign Sites.Selected. The logged-in account may need the 'AppRoleAssignment.ReadWrite.All' Graph permission or Global Admin role." -ForegroundColor Red
        if (-not (Get-Confirmation "Continue anyway?")) {
            exit
        }
    }
}

# ============================================================================
# STEP 6: Enter Site ID and Grant Permissions
# ============================================================================
Write-SectionHeader "STEP 6: Assign Permissions to Sites"

Write-Host "You will be prompted to enter a Site ID and Drive ID for each target." -ForegroundColor Yellow
Write-Host "Tip: Site IDs look like  'contoso.sharepoint.com,<guid>,<guid>'" -ForegroundColor Gray
Write-Host "     Drive IDs look like 'b!<base64-string>'" -ForegroundColor Gray

$sitesGranted = @()
$continueLoop = $true

while ($continueLoop) {
    Write-Host ""

    $enteredSiteId = (Read-Host "Enter the SharePoint Site ID").Trim()
    if ([string]::IsNullOrWhiteSpace($enteredSiteId)) {
        Write-Host "✗ Site ID cannot be empty." -ForegroundColor Red
        continue
    }

    $enteredDriveId = (Read-Host "Enter the Drive ID").Trim()
    if ([string]::IsNullOrWhiteSpace($enteredDriveId)) {
        Write-Host "✗ Drive ID cannot be empty." -ForegroundColor Red
        continue
    }

    Write-Host "`nGranting 'write' permission to UAMI '$uamiDisplay' on site '$enteredSiteId'..." -ForegroundColor White

    $permissionUrl  = "https://graph.microsoft.com/v1.0/sites/$enteredSiteId/permissions"
    $permissionBody = @{
        roles               = @("write")
        grantedToIdentities = @(
            @{
                application = @{
                    id          = $uamiAppId
                    displayName = $uamiDisplay
                }
            }
        )
    } | ConvertTo-Json -Depth 10

    $permissionResponse = Invoke-GraphApiRequest -Uri $permissionUrl -Method "POST" -Headers $headers -Body $permissionBody

    if ($permissionResponse) {
        Write-Host "✓ Permission granted on site '$enteredSiteId'" -ForegroundColor Green
        $sitesGranted += [PSCustomObject]@{
            SiteId  = $enteredSiteId
            DriveId = $enteredDriveId
        }
    }
    else {
        Write-Host "✗ Failed to grant permission on site '$enteredSiteId'" -ForegroundColor Red
    }

    if (-not (Get-Confirmation "`nDo you want to add another site?")) {
        $continueLoop = $false
    }
}

# ============================================================================
# Summary
# ============================================================================
Write-SectionHeader "Setup Complete!"

Write-Host "Summary:" -ForegroundColor Green
Write-Host "  UAMI          : $uamiDisplay" -ForegroundColor Gray
Write-Host "  Client ID     : $uamiAppId" -ForegroundColor Gray
Write-Host "  Sites granted : $($sitesGranted.Count)" -ForegroundColor Gray
Write-Host ""

foreach ($site in $sitesGranted) {
    Write-Host "  - Site ID : $($site.SiteId)" -ForegroundColor Gray
    Write-Host "    Drive ID: $($site.DriveId)" -ForegroundColor DarkGray
}

Write-Host "`n"
Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  Configuration completed successfully!                       ║" -ForegroundColor Green
Write-Host "║                                                               ║" -ForegroundColor Green
Write-Host "║  The UAMI now has site-level 'write' permissions.            ║" -ForegroundColor Green
Write-Host "║  No secrets were created or left behind.                     ║" -ForegroundColor Green
Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Green

# Clean up token from memory
$accessToken = $null
[System.GC]::Collect()

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
