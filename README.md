# neo-client-sharepoint

PowerShell script to assign scoped SharePoint site permissions to an Azure Service Principal.

## Overview

`AddPermissionToSites.ps1` is an interactive script that grants a Service Principal write access to specific SharePoint sites using the Microsoft Graph API **Sites.Selected** permission model. It can create a new Service Principal or work with an existing one. Elevated permissions (`Sites.FullControl.All`) and the temporary client secret are automatically cleaned up after setup.

## Prerequisites

- **Azure CLI** — [https://aka.ms/installazurecliwindows](https://aka.ms/installazurecliwindows)
- **PowerShell 5.1 or later** — pre-installed on Windows 10/11
- **Azure AD Admin Access** — Global Administrator or Privileged Role Administrator
- **Internet connection**

## How to Run

1. Open PowerShell and navigate to the script directory:
   ```powershell
   cd "path\to\neo-client-sharepoint"
   ```

2. Execute the script:
   ```powershell
   .\AddPermissionToSites.ps1
   ```

3. Follow the interactive prompts described below.

---

## Script Walkthrough

### Step 1 — Azure Authentication

The script checks whether you are already logged into Azure CLI and shows the current account and tenant. You can continue with that account or log out and re-authenticate.

### Step 2 — Service Principal Setup

You are asked whether to **use an existing** Service Principal or **create a new** one.

**Using an existing Service Principal:**
Enter the Client ID (Application ID). The script validates it against Azure AD and retrieves the display name.

**Creating a new Service Principal:**
You choose how to provide the name:

| Option | Input | Result |
|--------|-------|--------|
| 1 | Resource Group name (e.g. `rg-neo-abc-dev-uks-01`) | Strips `rg-` prefix, removes 2nd-last segment (location), condenses remaining parts → `Neologik SharePoint Access - neoabcdev01` |
| 2 | Short name (e.g. `neoabcdev01`) | Prepends standard prefix → `Neologik SharePoint Access - neoabcdev01` |
| 3 | Plain text | Used verbatim as the display name |

The derived name is shown for confirmation before anything is created.

### Step 3 — Check and Add Required Permissions

The script inspects the app's current Microsoft Graph API permissions and adds any that are missing:

- `Sites.Selected` — permanent, scoped site access
- `Sites.FullControl.All` — temporarily required to assign site-level permissions (removed in Step 9)

### Step 4 — Grant Admin Consent

Admin consent is granted automatically via Azure CLI. If that fails, a browser window opens for manual consent at the Microsoft admin consent URL. The script waits for your confirmation before continuing.

### Step 5 — Create Temporary Client Secret

A short-lived client secret named `test` is created on the app registration. This is used only to obtain an access token in the next step and is deleted at the end of the script.

### Step 6 — Generate Access Token

The script exchanges the client credentials for a Microsoft Graph access token using the OAuth 2.0 client credentials flow.

### Step 7 — List SharePoint Sites

All SharePoint sites in the tenant are retrieved and displayed as a numbered list with their URLs.

### Step 8 — Assign Permissions to Sites

Select sites by number. The script grants `write` access to each selected site via the Graph API. You can add multiple sites one at a time and stop when done.

### Step 9 — Security Cleanup

`Sites.FullControl.All` is revoked:
1. The active admin-consented role assignment is deleted.
2. All Graph permissions are removed from the app.
3. `Sites.Selected` is re-added and re-consented.

The app is left with only the minimum required permission.

### Step 10 — Delete Temporary Client Secret

The `test` client secret created in Step 5 is deleted. The app has no stored credentials after this point.

---

## Summary Output

At the end the script prints:

- Service Principal name and Client ID
- Number of sites configured and their display names

---

## Security Notes

- The temporary client secret is never displayed to the user and is deleted immediately after use.
- `Sites.FullControl.All` is automatically removed after site permissions are assigned.
- Sensitive values (`$clientSecretPlain`, `$accessToken`) are explicitly nulled and garbage-collected on exit.

---

## Troubleshooting

| Problem | Resolution |
|---------|------------|
| Azure CLI not found | Install from [https://aka.ms/installazurecliwindows](https://aka.ms/installazurecliwindows) and restart PowerShell |
| Login failed | Run `az logout`, then re-run the script |
| Service Principal not found | Verify the Client ID in Azure Portal > App registrations |
| Admin consent failed | Complete consent manually in the browser window that opens, then confirm when prompted |
| No sites listed | Permissions may not have propagated yet — wait a few minutes and re-run |
| Permission assignment failed | Verify admin consent shows "Granted" in Azure Portal and retry |

## Manual Verification

### Tenant-level permission (Azure Portal)
1. Go to **Azure Portal** > **Microsoft Entra ID** > **App registrations**
2. Find the application by name
3. **API permissions** — `Sites.Selected` should show status **Granted**
4. **Certificates & secrets** — the `test` secret should be absent after successful completion

### Site-level permission (Graph Explorer)
Permissions granted via `Sites.Selected` are **not visible in the SharePoint UI** (gear icon > Site permissions or SharePoint Admin Centre). The only way to confirm them is via the Microsoft Graph API:

1. Go to [https://developer.microsoft.com/en-us/graph/graph-explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in with your admin account
3. Run the following GET request, replacing `{site-id}` with the site's ID:
   ```
   GET https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
   ```
4. In the response, locate your app by its `clientId` — the `roles` field should show `["write"]`

---

For issues or questions, contact the Neologik team.
