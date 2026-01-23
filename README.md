# neo-client-sharepoint

PowerShell script to create a service principal with scoped access to customer's SharePoint specified sites

## Overview

This automated setup script creates and configures an Azure AD App Registration with Microsoft Graph API access for SharePoint using the Sites.Selected permission model. The script handles the complete setup process including authentication, app registration, permissions configuration, and site-specific access grants.

## Prerequisites

Before running the script, ensure you have:

- **Azure CLI**: Download and install from [https://aka.ms/installazurecliwindows](https://aka.ms/installazurecliwindows)
- **PowerShell 5.1 or later**: Pre-installed on Windows 10/11
- **Azure AD Admin Access**: You need one of the following roles:
  - Global Administrator
  - Privileged Role Administrator
- **Internet Connection**: Required for Azure authentication and API calls

## How to Run

### Step 1: Open PowerShell

1. Press `Win + X` and select "Windows PowerShell" or "Windows Terminal"
2. Navigate to the script directory:
   ```powershell
   cd "c:\Users\JashanpreetMagar\Documents\Qualyfi\neo-client-sharepoint"
   ```

### Step 2: Run the Script

Execute the script with:
```powershell
.\Setup-SharePointAPI.ps1
```

### Step 3: Follow the Interactive Prompts

The script will guide you through each step:

#### **Azure Authentication**
- A browser window will open for Azure login
- Sign in with your Azure AD admin credentials
- Confirm you want to proceed with the logged-in account

#### **App Registration**
- Enter a name for your application (e.g., "SharePoint API Access")
- The script will create the app registration automatically

#### **Admin Consent**
- A browser window will open for admin consent
- Review the requested permissions:
  - Sites.Selected
  - Sites.FullControl.All (temporarily, removed after setup)
- Check "Consent on behalf of your organization"
- Click "Accept"
- **Note**: You may see an error page saying "No reply address" - this is normal and means consent was successful

#### **SharePoint Site Configuration**
For each SharePoint site you want to grant access to:
- Enter the **SharePoint Hostname** (e.g., `contoso.sharepoint.com`)
- Enter the **Site Name** (e.g., `marketing`)
  - For URL `https://contoso.sharepoint.com/sites/marketing`, the site name is `marketing`
- Choose whether to add additional sites

#### **Save Configuration**
- Optionally save the configuration to a JSON file
- **Important**: The client secret is NOT saved to the file for security

## What the Script Does

1. **Validates Prerequisites**: Checks for Azure CLI installation
2. **Azure Authentication**: Logs you into Azure CLI
3. **Creates App Registration**: Registers a new Azure AD application
4. **Generates Client Secret**: Creates credentials for the app
5. **Configures Permissions**: Adds required Microsoft Graph API permissions
6. **Grants Admin Consent**: Obtains admin consent for the permissions
7. **Retrieves Site IDs**: Gets the unique ID for each SharePoint site
8. **Assigns Site Permissions**: Grants read access to specified sites
9. **Security Cleanup**: Removes Sites.FullControl.All permission, keeping only Sites.Selected

## Output

After successful completion, you'll receive:

- **Client ID**: Application (client) ID for your app
- **Tenant ID**: Your Azure AD tenant ID
- **Site IDs**: Unique identifiers for configured SharePoint sites
- **Configuration File**: Optional JSON file with setup details (excludes client secret)

## Important Security Notes

- The client secret is displayed only once during setup - save it securely
- The script automatically removes elevated permissions (Sites.FullControl.All) after setup
- Configuration files do NOT contain the client secret
- Sensitive data is cleared from memory after script completion

## Troubleshooting

### Azure CLI Not Found
- Install Azure CLI from [https://aka.ms/installazurecliwindows](https://aka.ms/installazurecliwindows)
- Restart PowerShell after installation

### Login Failed
- Ensure you have admin permissions in Azure AD
- Try logging out: `az logout` and run the script again

### Site ID Retrieval Failed
- Verify the site exists at the specified URL
- Check that the hostname and site name are correct
- Wait a few minutes for admin consent to propagate

### Permission Assignment Failed
- Ensure admin consent was granted successfully
- Verify both permissions show "Granted" status in Azure Portal
- Wait for permissions to propagate (up to 10 minutes)

## Manual Verification

To verify the setup in Azure Portal:

1. Go to **Azure Portal** > **Azure Active Directory** > **App registrations**
2. Find your application by name
3. Check **API permissions** - should show Sites.Selected as "Granted"
4. Check **Certificates & secrets** - should have an active client secret

## Support

For issues or questions, contact the Neologik team.
