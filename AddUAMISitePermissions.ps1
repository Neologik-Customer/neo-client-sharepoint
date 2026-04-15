$CLIENT_ID = ""  # UAMI client/app ID
$OBJECT_ID = ""  # UAMI service principal object ID
$siteUrl   = "" #EG https://company.sharepoint.com/sites/Test
$tenantId  = $(az account show --query tenantId -o tsv)
$TOKEN     = $(az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv)
$NAME      = $(az ad sp show --id $CLIENT_ID --query displayName -o tsv)
$headers   = @{Authorization="Bearer $TOKEN"; "Content-Type"="application/json"}

$graphSp           = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'" -Headers $headers
$sitesSelectedRole = ($graphSp.value[0].appRoles | Where-Object {$_.value -eq "Sites.Selected"})
$fullControlRole   = ($graphSp.value[0].appRoles | Where-Object {$_.value -eq "Sites.FullControl.All"})
$graphSpId         = $graphSp.value[0].id

# Step 1: Assign Sites.Selected to UAMI if not already assigned
$existingAssignments = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$OBJECT_ID/appRoleAssignments" -Headers $headers).value
if ($existingAssignments | Where-Object {$_.appRoleId -eq $sitesSelectedRole.id}) {
    Write-Host "Sites.Selected already assigned — skipping"
} else {
    Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$OBJECT_ID/appRoleAssignments" -Headers $headers -Body (@{principalId=$OBJECT_ID;resourceId=$graphSpId;appRoleId=$sitesSelectedRole.id} | ConvertTo-Json)
    Write-Host "Sites.Selected assigned"
}

# Step 2: Create temporary app + service principal
Write-Host "Creating temporary app..."
$tempApp   = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/applications" -Headers $headers -Body (@{displayName="Neologik Temporary Sharepoint Access"} | ConvertTo-Json)
$tempSp    = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" -Headers $headers -Body (@{appId=$tempApp.appId} | ConvertTo-Json)
$secretResp = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$($tempApp.id)/addPassword" -Headers $headers -Body (@{passwordCredential=@{displayName="temp"}} | ConvertTo-Json)

# Step 3: Assign Sites.FullControl.All to temp SP
Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($tempSp.id)/appRoleAssignments" -Headers $headers -Body (@{principalId=$tempSp.id;resourceId=$graphSpId;appRoleId=$fullControlRole.id} | ConvertTo-Json)
Write-Host "Sites.FullControl.All assigned to temp SP — waiting for propagation..."

# Step 4: Get token for temp SP via client credentials (retry until role propagates)
$tempToken = $null
$attempts  = 0
do {
    Start-Sleep -Seconds 15
    $attempts++
    try {
        $tokenResp = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
            grant_type    = "client_credentials"
            client_id     = $tempApp.appId
            client_secret = $secretResp.secretText
            scope         = "https://graph.microsoft.com/.default"
        }
        # Verify the token actually contains Sites.FullControl.All
        $payload = $tokenResp.access_token.Split('.')[1]
        $padded  = $payload + ('=' * ((4 - $payload.Length % 4) % 4))
        $claims  = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($padded)) | ConvertFrom-Json
        if ($claims.roles -contains "Sites.FullControl.All") {
            $tempToken = $tokenResp.access_token
            Write-Host "Token with Sites.FullControl.All obtained (attempt $attempts)"
        } else {
            Write-Host "Role not yet in token (attempt $attempts) — retrying..."
        }
    } catch {
        Write-Host "Token request failed (attempt $attempts) — retrying..."
    }
} while (-not $tempToken -and $attempts -lt 10)

if (-not $tempToken) {
    Write-Host "ERROR: Could not obtain token with Sites.FullControl.All after $attempts attempts"
    # Clean up before exiting
    Invoke-RestMethod -Method DELETE -Uri "https://graph.microsoft.com/v1.0/applications/$($tempApp.id)" -Headers $headers
    exit 1
}

$tempHeaders = @{Authorization="Bearer $tempToken"; "Content-Type"="application/json"}

# Step 5: Look up site ID and grant UAMI write access
$parsedUri = [System.Uri]$siteUrl
$site   = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($parsedUri.Host):$($parsedUri.AbsolutePath)" -Headers $tempHeaders
$siteId = $site.id

Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions" -Headers $tempHeaders -Body (@{
    roles               = @("write")
    grantedToIdentities = @(@{application=@{id=$CLIENT_ID;displayName=$NAME}})
} | ConvertTo-Json -Depth 10)

Write-Host "Write access granted to $NAME on $($site.displayName)"

# Step 6: Delete the temporary app
Invoke-RestMethod -Method DELETE -Uri "https://graph.microsoft.com/v1.0/applications/$($tempApp.id)" -Headers $headers
Write-Host "Temporary app deleted"
