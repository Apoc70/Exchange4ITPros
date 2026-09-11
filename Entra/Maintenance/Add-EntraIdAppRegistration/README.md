# Add-EntraIdAppRegistration.ps1

`Add-EntraIdAppRegistration.ps1` creates an Entra ID app registration configured for Microsoft Graph access. It assigns an owner, applies the permissions defined in `AppPermissions.json`, and creates either a client secret or a certificate-ready application.

## Requirements

- PowerShell 7.1 or later.
- The Microsoft Graph PowerShell SDK modules `Microsoft.Graph.Authentication` and `Microsoft.Graph.Applications`.
- The account running the script must be assigned the **Application Administrator** or **Global Administrator** role.
- `AppPermissions.json` must be available in the same folder unless another path is supplied with `-PermissionsConfigPath`.
- The owner supplied with `-AppOwnerEmailAddress` must be a valid user in the target tenant.
- To use `-GrantAdminConsent`, the account must also be assigned **Privileged Role Administrator** or **Global Administrator**.

Install the required Microsoft Graph modules in an administrative PowerShell session:

```powershell
Install-Module Microsoft.Graph -Scope AllUsers
```

## What the script does

1. Loads and validates the permissions from `AppPermissions.json`.
2. Connects to Microsoft Graph with the scopes required to create and configure the application.
3. Stops if an app registration with the same display name already exists.
4. Creates the app registration and assigns the specified owner.
5. Optionally uploads an application logo.
6. Creates a client secret or leaves the application ready for a certificate.
7. Applies the configured API permissions and the public client redirect URI.
8. Optionally grants tenant-wide admin consent.
9. Displays the application ID and, when created, the client secret.

The script does not overwrite an existing app registration with the same display name.

## Basic usage

Run the script from its directory, or provide the full script path:

```powershell
.\Add-EntraIdAppRegistration.ps1 `
  -TenantId "00000000-0000-0000-0000-000000000000" `
    -AppName "Entra Enterprise Apps Report" `
    -AppDescription "Application for Entra enterprise app reporting" `
    -AppOwnerEmailAddress "admin@contoso.com" `
    -AppSecretName "EntraEnterpriseAppsSecret" `
    -AppSecretValidityInMonths 12
```

`-TenantId` accepts a tenant GUID. If it is omitted, Microsoft Graph uses its normal tenant selection and sign-in behavior.

## Certificate-based authentication

Use `-UseCertificate` when the application should not receive a client secret:

```powershell
.\Add-EntraIdAppRegistration.ps1 `
  -TenantId "00000000-0000-0000-0000-000000000000" `
    -AppName "Entra Enterprise Apps Report" `
    -AppOwnerEmailAddress "admin@contoso.com" `
    -UseCertificate
```

After the application is created, upload the public certificate (`.cer`) manually to the app registration in the Entra portal. The script prints the credentials URL for the new application.

## Admin consent

By default, the script prints an Entra portal URL for granting admin consent. Use `-OpenBrowser` to open that URL automatically after a 30-second provisioning delay:

```powershell
.\Add-EntraIdAppRegistration.ps1 `
  -TenantId "00000000-0000-0000-0000-000000000000" `
    -AppName "Entra Enterprise Apps Report" `
    -AppOwnerEmailAddress "admin@contoso.com" `
    -GrantAdminConsent `
    -OpenBrowser
```

Use `-PrivateBrowserSession` together with `-OpenBrowser` to open the URL in a private browser window. The script tries Microsoft Edge, Google Chrome, and Firefox in that order.

`-GrantAdminConsent` grants the permissions from `AppPermissions.json` by code. This requires elevated directory permissions and can fail for individual permissions if the tenant or resource service principal does not support the requested grant. Review the console output after the run.

## Parameters

| Parameter | Description |
| --- | --- |
| `-TenantId` | Tenant GUID to use for the Microsoft Graph connection. |
| `-AppName` | Display name of the new app registration. Defaults to `AppRegistration Report`. The name must not already exist in the tenant. |
| `-AppDescription` | Description stored in the application notes. |
| `-AppSecretName` | Display name of the client secret. |
| `-AppOwnerEmailAddress` | User principal name or user ID of the application owner. |
| `-AppSecretValidityInMonths` | Number of months until the client secret expires. |
| `-UseCertificate` | Skips client-secret creation and requires a public certificate to be uploaded manually. |
| `-AppLogoPath` | Optional path to a `.png`, `.jpg`, `.jpeg`, or `.gif` logo no larger than 256 KB. |
| `-OpenBrowser` | Opens the admin-consent URL after the application is provisioned. |
| `-PrivateBrowserSession` | Opens the admin-consent URL in a private browser window. Requires `-OpenBrowser`. |
| `-PermissionsConfigPath` | Path to a permissions JSON file. Defaults to `AppPermissions.json` beside the script. |
| `-GrantAdminConsent` | Grants configured application and delegated permissions by code instead of requiring portal consent. |

## Permissions configuration

The default `AppPermissions.json` contains Microsoft Graph permissions. Each resource entry must include a resource application ID and one or more permission entries:

```json
{
  "requiredResourceAccess": [
    {
      "resourceName": "Microsoft Graph",
      "resourceAppId": "00000003-0000-0000-c000-000000000000",
      "resourceAccess": [
        {
          "name": "Application.Read.All",
          "id": "9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30",
          "type": "Role"
        }
      ]
    }
  ]
}
```

The script uses each permission `id` and `type` when configuring the app. `Role` entries are treated as application permissions; other types are resolved as delegated permissions when admin consent is granted by code. Permission IDs can be found in the [Microsoft Graph permissions reference](https://learn.microsoft.com/graph/permissions-reference#all-permissions-and-ids).

## Secret handling

The generated client secret is displayed only once by Microsoft Graph. Store it immediately in a secure secret store and never commit it, tenant-specific configuration, or credentials to source control.

The script prints the new **Client ID (App ID)** and the client secret value when applicable. The application object ID is not printed as a configuration value; it is used internally while the script runs.

## Related files

- [`Add-EntraIdAppRegistration.ps1`](Add-EntraIdAppRegistration.ps1) - Provisioning script.
- [`AppPermissions.json`](AppPermissions.json) - Default Microsoft Graph permission definitions.
- [`Get-EntraEnterpriseApps.ps1`](Get-EntraEnterpriseApps.ps1) - Generates enterprise application and app registration reports.
- [`README.md`](README.md) - Folder-level documentation for the reporting scripts.

## Links

- [Exchange for IT Pros community](https://exchangeforitpros.blog/) - Visit the website and engage with the community by sharing your experience and questions.
