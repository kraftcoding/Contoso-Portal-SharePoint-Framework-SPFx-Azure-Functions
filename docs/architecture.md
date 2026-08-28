# Contoso Portal — Architecture Overview

> **DISCLAIMER:** Fictitious demo project. All names, IDs, and data are fictional and for educational purposes only.

---

## 1. Solution Structure

```
DEV_Contoso/
├── git/
│   ├── ALM/              # CI/CD pipeline definitions (Azure DevOps)
│   ├── Azure/            # Backend: Azure Functions + class libraries
│   │   ├── Infrastructure/   # ARM/Bicep templates for Azure resources
│   │   ├── WebApi/           # Azure Functions host project
│   │   ├── WebApi.Business/  # Domain service layer
│   │   ├── WebApi.Data/      # Data access layer (SharePoint + Azure Tables)
│   │   ├── WebApi.Model/     # Domain model (DTOs, enums)
│   │   └── WebApi.Tests/     # Unit and integration tests
│   ├── Provisioning/     # PnP PowerShell provisioning
│   ├── SearchTemplates/  # SharePoint search result types
│   └── SPFx/             # SharePoint Framework frontend
└── docs/                 # Documentation
```

---

## 2. Domain Model

### Hierarchy

```
BusinessUnit  (e.g. Finance, HR, Technology)
  └── Department  (e.g. Finance-Budget, HR-Talent)
        └── Meeting  (a scheduled meeting instance)
              ├── AgendaItem   (items on the agenda)
              ├── Attendance   (who attends, in-person or online)
              ├── Delegation   (who delegates attendance or vote)
              ├── Minutes      (official record of the meeting)
              ├── Agreement    (agreed items / resolutions)
              ├── Vote         (voting on agenda items)
              └── Certificate  (certification document)
```

### Meeting Lifecycle States

```
Draft → Reserved → Published → InProgress → Held → Finished → Archived
                                                  └──────────────→ Cancelled
```

### Task Types (pending user actions)

| Task | Description |
|------|-------------|
| `AttendanceRequest` | User must confirm attendance |
| `DelegationRequest` | User must accept/reject a delegation |
| `MinutesApproval` | User must approve the meeting minutes |
| `MinutesModification` | User must review a minutes change request |
| `CertificationRequest` | Generate a certification document |

---

## 3. Backend (Azure Functions)

### Projects

| Project | Namespace | Role |
|---------|-----------|------|
| `WebApi` | `Contoso.Portal` | Functions host, DI wiring, triggers |
| `WebApi.Business` | `Contoso.Portal.Business` | Domain services |
| `WebApi.Data` | `Contoso.Portal.Data` | DAL, DAO, DTO classes |
| `WebApi.Model` | `Contoso.Portal.Model` | Shared domain models |

### Key Services (WebApi.Business)

| Service | Responsibility |
|---------|----------------|
| `BodiesService` | CRUD for Departments |
| `BodyRoleService` | Role-based permissions per department |
| `EventsService` | Meeting creation and management |
| `EventAgendaItemsService` | Agenda items |
| `EventAttendanceService` | Attendance registration |
| `EventMinutesService` | Minutes drafting and publishing |
| `EventPublishingService` | Meeting publishing workflow |
| `EventVotationService` | Voting management |
| `NotificationsService` | In-app and email notifications |
| `TasksService` | Pending-task management |
| `ConfigDepartmentsService` | Department taxonomy/config |
| `ManagementService` | Admin operations |
| `ProfileService` | User profile data |
| `TaxonomyTranslationService` | Multilingual taxonomy lookup |
| `DocumentToPdfService` | PDF conversion |

### Mock Integration Services

Removed external integrations are replaced with mock stubs:

| Interface | Mock Implementation | Original purpose |
|-----------|---------------------|-----------------|
| `INotificationSmsService` | `MockNotificationSmsService` | SMS notification gateway |
| `IDocumentArchiveService` | `MockDocumentArchiveService` | Document archive / ENI format conversion |

### Data Access Pattern

```
Service (Business) → DALProvider (Data) → DAO (Data) → SharePoint / Azure Table Storage
```

- **SharePoint Lists** — primary store for meetings, tasks, notifications, documents
- **Azure Table Storage** — async/offline data copy for reporting and high-volume reads

### Triggers

| Trigger | Type | Purpose |
|---------|------|---------|
| `WebhooksHttpTrigger` | HTTP | SharePoint change notifications webhook |
| `PingHttpTrigger` | HTTP | Health check |
| Various HTTP triggers | HTTP | REST API endpoints |
| Timer triggers | Timer | Background scheduled jobs |

---

## 4. Frontend (SPFx)

### Web Parts

| Web Part | Description |
|----------|-------------|
| `meetings` | Main meeting list and detail view |
| `myDepartments` | Current user's department memberships |
| `pendingTasks` | Tasks awaiting user action |
| `nextEvents` | Upcoming meetings widget |
| `recentDocuments` | Recently accessed meeting documents |
| `administrationApp` | Admin panel for department management |
| `certificationApp` | Certification document generation |
| `departmentManagement` | Department configuration |
| `businessUnitRelations` | Business unit relationship matrix |
| `directLinks` | Quick links widget |
| `documents` | Document library view |
| `agendaAnnouncement` | Agenda announcement banner |
| `delegateAdministrationBanner` | Delegation management banner |
| `welcomeMessage` | Personalized welcome widget |
| `myCollaborationSpaces` | Teams/SharePoint spaces |
| Master table web parts | CRUD for taxonomy master tables |

### Extensions

| Extension | Type | Description |
|-----------|------|-------------|
| `LaunchDataProcessCommandSet` | List Command | Trigger background data operations |
| `SendDocToBodyCommandSet` | List Command | Send documents to a department |

### Service Layer (SPFx)

- `BackendService` — Calls Azure Function REST API
- `Service` — SharePoint REST calls
- `RoleService` — Role/permission checks

---

## 5. Provisioning

PnP PowerShell scripts automate SharePoint site creation and configuration.

### Key Scripts

| Script | Description |
|--------|-------------|
| `Install.ps1` | Main provisioning entry point |
| `ApplyPnPHomeTemplate.ps1` | Apply the portal home PnP template |
| `CreateDepartmentsFromConfigList.ps1` | Create department sites from config list |
| `CreateHubSite.ps1` | Create the hub site |
| `ApplyTaxonomies.ps1` | Apply term store taxonomies |
| `ChannelConfiguration.ps1` | Configure Teams channels |

### PnP Templates

| Template | Description |
|----------|-------------|
| `templateHomePortal.xml` | Portal home site template |
| `templateDepartment.xml` | Department site template |
| `tenantTemplate.xml` | Tenant-level configuration |
| `TemplateDemoEntities.xml` | Demo entities registry template |

### Config Files

| File | Environment |
|------|-------------|
| `Config/Config-demo.xml` | Demo / local development |

---

## 6. ALM / CI-CD (Azure DevOps)

### Pipelines

| File | Stages |
|------|--------|
| `az-pipe-base.yaml` | Build → Deploy |
| `az-pipe-base-ci.yaml` | Build only (CI) |
| `az-pipe-deploy-script.yaml` | Script deployment |
| `az-pipe-build-sonar.yml` | Build + SonarQube analysis |

### Environments

| Name | Purpose |
|------|---------|
| `dev` | Developer environment |
| `test` | QA / testing environment |
| `prod` | Production environment |

---

## 7. Infrastructure (Azure)

ARM/Bicep templates in `Azure/Infrastructure/`:

- `contoso-template.json` — Main Azure Function App + Storage + App Insights
- `contoso-template.lowcost.json` — Low-cost variant (shared hosting plan)
- `contoso-azureStorage-template.json` — Azure Table Storage configuration
- `AppRegistration/` — App registration manifest templates

---

## 8. Security Model

- **Authentication:** Certificate-based (X.509) using Azure AD app registration
- **Authorization:** Role-based per department, defined in SharePoint group membership
- **Roles:**
  - `administrators` — Full admin
  - `schedulers` — Can schedule meetings
  - `schedulerAssistants` — Admin assistants for schedulers
  - `members` — Voting members
  - `memberAssistants` — Member assistants
  - `guests` — Read-only observers

---

*Document generated for the Contoso Portal demo project — all content is fictional and for educational purposes only.*
