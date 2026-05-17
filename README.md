# SharePoint Smart Copy

Website
**[https://sharepointsmartsolutions.com/smart-copy](https://sharepointsmartsolutions.com/smart-copy)**

[User Guide](SharePointSmartCopy/Docs/SharePointSmartCopy_UserGuide.docx)

A WPF .NET 8 desktop application that copies files with full version history and full metadata between SharePoint Online site collections.

## Features

Do what you can't do in Powershell, and have to pay for with migration tools:

- Full version history along with matching metadata gets copied for each version
- Bulk copy with 1-16 parallel file copies
- Easy to use UI to select which files or folders to transfer
- Detailed reporting done at the end

## Prerequisites

- Windows 10/11
- .NET 8 Desktop Runtime
- A Microsoft 365 tenant with SharePoint Online
- An Azure AD (Entra ID) app registration (see setup below)

## Azure AD App Registration

### 1. Register the app

1. Go to [Entra ID](https://entra.microsoft.com) → **App registrations** → **New registration**
2. Name: `SharePoint Smart Copy` (or similar)
3. Supported account types: **Single tenant** (your org only) or Multitenant as needed
4. Redirect URI: **Public client/native** → `http://localhost`
5. Click **Register**
6. Copy the **Application (client) ID** and **Directory (tenant) ID** — you will enter these in the app's Settings dialog

### 2. Add API permissions

Go to **API permissions** → **Add a permission**:

| API | Type | Permission | Purpose |
|-----|------|-----------|---------|
| Microsoft Graph | Delegated | `Sites.ReadWrite.All` | Browse and read SharePoint sites/files via Graph |
| Microsoft Graph | Delegated | `Files.ReadWrite.All` | Upload/download file content via Graph |
| SharePoint | Delegated | `AllSites.FullControl` | Required by the Migration API to submit migration jobs; also allows correct `IsSiteAdmin` evaluation in the OAuth context |

> **Why `AllSites.FullControl`?**
> SharePoint's Migration API (`CreateMigrationJobEncrypted`) performs a server-side site-collection-administrator check.
> With only `Sites.ReadWrite.All`, SharePoint caps the effective OAuth privilege below site-collection-admin level — meaning even a user who is explicitly a Primary Site Admin will be rejected.
> `AllSites.FullControl` raises the OAuth context to full control so SP recognizes the user's actual admin status.
>
> `AllSites.FullControl` is only required for Migration API mode. If your organization uses Enhanced REST mode exclusively, this permission can be omitted.

### 3. Grant admin consent

After adding the permissions, click **Grant admin consent for [your organization]** at the top of the API permissions list.
This pre-authorises the permissions org-wide so users are never prompted for individual consent.

Without admin consent, users will see an interactive consent dialog on first use.
As a Global Admin you can also check **"Consent on behalf of your organization"** in that dialog, which has the same effect.

### 4. Target site permissions (Migration API only)

The account running the copy must be a **Site Collection Administrator** on the **target** site:

> Site Settings → Site Collection Administrators → add your account

Being a Global Admin or SharePoint Admin grants effective access to all sites but does **not** automatically populate the Site Collection Administrators list for a specific site. You must add the account explicitly.

This requirement applies only to Migration API mode. Enhanced REST mode works with standard contributor access.

## Configuration

Launch the app and open **Settings** (gear icon):

- **Client ID** — Application (client) ID from the app registration
- **Tenant ID** — Directory (tenant) ID (leave blank for multi-tenant)

Source/target URLs and copy preferences are configured within the wizard and remembered between sessions.

## Copy Modes

### When to use each mode

| Scenario | Recommended mode |
|---|---|
| Large batch (50+ files or 200+ versions) | Migration API |
| Full version history fidelity required | Migration API |
| Small batch or a quick one-off copy | Enhanced REST |
| Copying current version only (no history) | Enhanced REST |
| User lacks Site Collection Admin rights | Enhanced REST |
| Need to see per-file progress in real time | Enhanced REST |

The copy mode option appears on the Options screen when **Copy versions** is enabled. Hover the ⓘ icon next to each mode name for a quick summary.

---

### Migration API

Uses SharePoint's built-in [Migration API](https://learn.microsoft.com/en-us/sharepoint/dev/apis/migration-api-overview). Files are packaged client-side, uploaded to SP-provisioned Azure Blob containers, then imported server-side by SharePoint.

**Advantages**
- Version numbers on target exactly match source (1.0, 2.0, 3.0, …)
- Modified By and Modified date correct per version in history
- Author and Created date preserved on the file
- Bypasses per-item throttling — SP processes the batch as a single job, not thousands of individual API calls
- Scales well: 500 files with 10 versions each has roughly the same client-side overhead as 50 files

**Limitations**
- Minimum ~1–2 minutes of overhead per run regardless of file count (container provisioning, manifest packaging, blob upload, SP processing)
- No per-file progress during SP's processing phase — results appear only after the full job completes
- Error reporting is at the job level; individual file failures may have limited detail
- Requires elevated permissions (see below)

**Required permissions for Migration API**

| Requirement | Where to configure |
|---|---|
| `AllSites.FullControl` delegated permission on the SharePoint API | Azure AD app registration → API permissions |
| Site Collection Administrator on the **target** site | Site Settings → Site Collection Administrators |

> SP's Migration API performs a server-side site-collection-administrator check on every job submission. Standard `Sites.ReadWrite.All` is not sufficient — even a user who is explicitly a Site Admin will be rejected unless the OAuth context carries `AllSites.FullControl`. The account also needs to appear in the Site Collection Administrators list on the target site, not just have a SharePoint Admin role at the tenant level.

---

### Enhanced REST

Uses the SharePoint REST and Microsoft Graph APIs directly. Each file version is uploaded individually, with metadata and timestamps patched immediately after.

**Advantages**
- Results appear per file as each one completes — you see progress in real time
- Low overhead for small batches: a 5-file copy completes in seconds
- No elevated permissions required beyond standard contributor access
- Per-file error messages are clear and immediate

**Limitations**
- Version numbers are 2× the source count (e.g. versions 2, 4, 6 for a 3-version source file) — a SharePoint REST constraint; the correct dates and editors are still preserved
- Subject to SharePoint throttling (HTTP 429) on large batches with high parallelism
- Slower than Migration API for large migrations with many versions

## NuGet Packages

| Package | Version | Purpose |
|---------|---------|---------|
| `Microsoft.Graph` | 5.x | Graph API client (site/file browsing, download, Enhanced REST upload) |
| `Microsoft.Identity.Client` | 4.x | MSAL — interactive sign-in and token management |
| `Azure.Storage.Blobs` | 12.x | Upload encrypted blobs to SP-provisioned containers (Migration API) |
| `Microsoft.SharePointOnline.CSOM` | 16.x | `EncryptionOption` type used in Migration API package |
| `CommunityToolkit.Mvvm` | 8.x | MVVM source generators for the WPF view models |
