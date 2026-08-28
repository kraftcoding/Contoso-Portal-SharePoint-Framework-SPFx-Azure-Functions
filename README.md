# Contoso Portal — Demo Project

> **DISCLAIMER:** This is a **fictitious, educational demo project**. All company names, tenant IDs, URLs, email addresses, GUIDs, certificates, and data are **entirely fictional** and created solely for educational and demonstration purposes. This project has no affiliation with any real organization, government, or institution. It is **not production-ready** and must not be used with real credentials or deployed to real tenants without a complete security review.

---

## Overview

**Contoso Portal** is a SharePoint Framework (SPFx) + Azure Functions demo solution that demonstrates how to build a corporate meeting-management platform on Microsoft 365.

The solution covers the full lifecycle of department meetings: creation, agenda management, attendance, minutes, approvals, certifications, delegations, notifications, and document publishing — all built on SharePoint Online and Microsoft Graph.

### Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                       Microsoft 365 Tenant                    │
│  ┌─────────────┐   ┌───────────────────────────────────────┐ │
│  │  SharePoint │   │          Azure Functions              │ │
│  │   Online    │◄──┤  WebApi (HTTP + Timer Triggers)       │ │
│  │  (SPFx WPs) │   │  ├── WebApi.Business (domain logic)   │ │
│  └─────────────┘   │  ├── WebApi.Data (SharePoint / Table) │ │
│                     │  └── WebApi.Model (domain models)     │ │
│                     └───────────────────────────────────────┘ │
│  ┌──────────────┐                                             │
│  │ Azure Storage│  (Azure Table Storage for async data)       │
│  └──────────────┘                                             │
└──────────────────────────────────────────────────────────────┘
```

### Main Components

| Component | Description |
|-----------|-------------|
| `SPFx/` | SharePoint Framework web parts and extensions |
| `Azure/WebApi` | Azure Functions host (HTTP + Timer triggers) |
| `Azure/WebApi.Business` | Domain services (meetings, departments, tasks, notifications) |
| `Azure/WebApi.Data` | Data access layer (SharePoint PnP, Azure Tables) |
| `Azure/WebApi.Model` | Domain model classes and enums |
| `Azure/WebApi.Tests` | Unit and integration tests |
| `Provisioning/` | PnP PowerShell provisioning scripts and templates |
| `ALM/` | Azure DevOps CI/CD pipeline definitions |
| `Azure/Infrastructure/` | ARM/Bicep infrastructure templates |

### Domain Model

```
BusinessUnit
  └── Department
        └── Meeting (Draft → Reserved → Published → InProgress → Held → Finished → Archived)
              ├── AgendaItem
              ├── Attendance (Attendee, Delegate)
              ├── Minutes (draft → approved → published)
              ├── Agreement
              └── Vote
```

---

## Prerequisites

- .NET 8 SDK
- Node.js 18+ and npm
- SharePoint Framework toolchain: `npm install -g @microsoft/generator-sharepoint`
- PnP PowerShell: `Install-Module PnP.PowerShell`
- Azure Functions Core Tools v4
- Azure CLI (for infrastructure deployment)
- A Microsoft 365 developer tenant (free: [developer.microsoft.com/microsoft-365/dev-program](https://developer.microsoft.com/microsoft-365/dev-program))

---

## Setup

### 1. Configure Azure App Registration

Register an application in Azure AD with:
- Certificate-based authentication (upload a self-signed certificate)
- **API Permissions:** `Sites.FullControl.All`, `User.Read.All`, `Mail.Send`, `TeamSettings.ReadWrite.All`

Update `appsettings.json` (or `local.settings.json` for local dev) with your real values:

```json
{
  "TenantId": "<your-tenant-id>",
  "ClientId": "<your-app-registration-client-id>",
  "CertificateThumbPrint": "<your-certificate-thumbprint>",
  "RootSiteUrl": "https://<your-tenant>.sharepoint.com/sites/contoso-portal"
}
```

> **Security note:** Never commit real secrets, certificates, or tenant IDs to version control.

### 2. Provision SharePoint Structure

```powershell
cd Provisioning
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com" -Interactive
.\Install.ps1 -ConfigFile ".\Config\Config-demo.xml"
```

### 3. Build and Run the Web API (Local)

```bash
cd Azure/WebApi
dotnet restore
dotnet build
func start
```

The API will be available at `http://localhost:7071`.

### 4. Build and Deploy the SPFx Package

```bash
cd SPFx
npm install
gulp bundle --ship
gulp package-solution --ship
```

Upload the generated `.sppkg` from `sharepoint/solution/` to your SharePoint App Catalog.

---

## ALM / CI-CD

The `ALM/CICD/` folder contains Azure DevOps pipeline templates:

| Pipeline | Description |
|----------|-------------|
| `az-pipe-base.yaml` | Main CI/CD pipeline (build + deploy) |
| `az-pipe-base-ci.yaml` | CI-only build |
| `az-pipe-deploy-script.yaml` | Script-based deployment |

Environments: `dev`, `test`, `prod`

---

## Configuration Reference

| Key | Description | Example |
|-----|-------------|---------|
| `TenantId` | Azure AD Tenant ID | `00000000-...` |
| `ClientId` | App Registration Client ID | `00000000-...` |
| `CertificateThumbPrint` | Certificate thumbprint | `AAAA1111...` |
| `RootSiteUrl` | SharePoint root site URL | `https://contoso.sharepoint.com/sites/contoso-portal` |
| `StorageAccountName` | Azure Storage Account | `dbdevcontoso001` |
| `DefaultLocale` | Default locale | `en-US` |
| `VotingEnabled` | Enable voting feature | `true`/`false` |
| `GraphMailboxUserId` | Graph API mailbox user ID | `00000000-...` |
| `TaxonomyRootSiteId` | Taxonomy site admin URL | `contoso-admin.sharepoint.com` |

---

## Legal & Safety Notice

This project is released for **educational and demonstration purposes only** under the MIT License.

- All entity names, GUIDs, URLs, email addresses, certificates, and configuration values are **entirely fictitious**.
- No real organizational data, personal data, or government data is present.
- The project is **not production-ready** without proper security hardening, real credentials, and compliance review.
- The mock integration services (`MockNotificationSmsService`, `MockDocumentArchiveService`) are stubs and do not connect to any real external system.

---

## License

MIT License — see [LICENSE](LICENSE) for details.
