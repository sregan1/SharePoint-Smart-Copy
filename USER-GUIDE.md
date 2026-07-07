# SharePoint Smart Copy — User Guide

*Copy files and folders between SharePoint Online sites*

**Version 3.3.0  ·  July 2026**

---

## Table of Contents

1. [Overview](#1-overview)
2. [Security](#2-security)
3. [System Requirements](#3-system-requirements)
4. [Setup — Azure App Registration](#4-setup--azure-app-registration)
5. [Step-by-Step Guide](#5-step-by-step-guide)
6. [Copy Modes — Detailed Reference](#6-copy-modes--detailed-reference)
7. [Copy History](#7-copy-history)
8. [Troubleshooting](#8-troubleshooting)

---

## 1. Overview

SharePoint Smart Copy is a Windows desktop application that copies files and folders between SharePoint Online site collections, preserving full version history and metadata. Whether you are migrating content to a new site, archiving documents, or reorganising your SharePoint environment, a guided wizard handles version history, metadata preservation, and parallel transfers — no scripting or PowerShell required.

Two copy engines are available. **Migration API** mode uses SharePoint's server-side import pipeline to achieve exact version fidelity — the same version numbers, dates, and editors appear on the target as on the source. **Enhanced REST** mode calls the Graph API directly for each version, preserving dates and editors but producing version numbers that are 2× the source count due to a SharePoint REST constraint. Both modes authenticate through your existing Microsoft 365 credentials.

### Key Capabilities

- Copy individual files, entire folders, complete document libraries, entire sites, or modern pages
- Preserve version history — copy all versions or limit to the N most recent
- **Migration API mode** — exact version numbers, exact dates, exact editors (requires Site Collection Admin on target)
- **Enhanced REST mode** — correct dates and editors per version; no admin rights required
- **Copy-if-newer** incremental mode — only copies files that are newer on the source, skipping files already up to date
- **Person/User, Managed Metadata, and Lookup columns** copied alongside content — no manual field re-entry
- Custom column mapping dialog — map source columns to matching target columns or create missing ones
- Automatic HTTP 429 throttle handling — large jobs complete without interruption
- Parallel transfers with 1–16 simultaneous file copies for faster bulk operations
- Real-time progress monitoring with per-file status updates (Enhanced REST) or job-level results (Migration API)
- Copy report with succeeded, failed, and skipped counts, inline permission status, and CSV export
- **Verification Report** — independently re-scans source and target after a run and produces an Excel workbook confirming what actually matches
- Full copy history stored locally — browse, re-export, or delete previous runs
- System sleep is blocked automatically for the duration of a copy, metadata update, or verification
- ← Back navigation available on every step, including the final report screen
- Light, Dark, and System-follows-Windows themes

---

## 2. Security

SharePoint Smart Copy is built on Microsoft's own identity platform and never handles your credentials directly. All communication is with Microsoft's servers only.

### Authentication — OAuth 2.0 / MSAL

When you click **Connect**, your default browser opens to Microsoft's login page — the exact same page you use to sign in to SharePoint, Teams, or any other Microsoft 365 service. The application never sees your password. After you authenticate, Microsoft returns a short-lived access token which the app uses to call the Graph API on your behalf.

### No Stored Credentials

Your Microsoft 365 password is never entered into or stored by the application. The access token is held in memory only for the duration of the session and discarded when the application closes. If you select **Remember me** in the browser during sign-in, the Microsoft Authentication Library (MSAL) stores a refresh token in the Windows Credential Manager — the same secure store used by Microsoft Office applications.

### Delegated Permissions — Acts as You

The application uses delegated permissions, meaning it acts as the signed-in user. It can only access content that your account already has permission to access through SharePoint. It cannot touch any site, library, folder, or file that you could not access in your own browser.

### AllSites.FullControl — Migration API Only

When Migration API mode is used for version history, the SharePoint API permission `AllSites.FullControl` is required. This is still a delegated permission — the application acts as the signed-in user and is limited to what that user can actually do. The permission raises the OAuth context to site-collection-admin level so SharePoint recognizes the user's existing admin status; it does not grant any access beyond what the signed-in account already holds.

> **Note:** `AllSites.FullControl` is only exercised during Migration API jobs. Enhanced REST mode does not require it. If your organization does not want to grant this permission, use Enhanced REST mode instead.

### Your Own App Registration

You create and own the Azure AD app registration in your organization's tenant (see [Section 4](#4-setup--azure-app-registration)). Your IT department controls which permissions are granted, who may use the application, and can revoke access at any time through the Azure portal — with no involvement from any third party.

### No External Servers or Telemetry

All data transfers occur directly between your computer and Microsoft's Graph API and SharePoint services. There are no intermediate servers, no analytics pipeline, and no telemetry. Your files never leave Microsoft's infrastructure.

### Local Storage Only

The only data written to your local machine is:

- **Application settings** (Azure AD registration IDs) — `%AppData%\SharePointSmartCopy\settings.json`
- **Copy history reports** — `%AppData%\SharePointSmartCopy\Reports\` (filenames and status only, no file content)
- **MSAL token cache** — managed by the Microsoft Authentication Library in the Windows Credential Manager

> **Note:** No file content is ever written to disk by this application. All copy operations stream data directly from the source SharePoint site to the target.

---

## 3. System Requirements

- Windows 10 (version 1809 or later) or Windows 11
- .NET 8 Desktop Runtime — download free from microsoft.com/dotnet
- Internet access to Microsoft 365 / SharePoint Online
- A default browser for the initial Microsoft 365 sign-in flow
- An Azure Active Directory / Entra ID app registration (see [Section 4](#4-setup--azure-app-registration))

### SharePoint Permissions Required

The signed-in Microsoft 365 account must have:

- Read access (Visitor or higher) on all source libraries and folders being copied from
- Contribute or higher access on the target SharePoint site and destination folder
- If copying version history: **Edit** or higher on the source library (required to read version history via the Graph API)
- If using Migration API mode: **Site Collection Administrator** on the target site (Site Settings → Site Collection Administrators)

> **Note:** Site Collection Administrator is distinct from SharePoint Administrator at the tenant level. Even a Global Admin must be added explicitly to the Site Collection Administrators list on the specific target site.

---

## 4. Setup — Azure App Registration

Before using SharePoint Smart Copy for the first time, a one-time setup is required: registering the application in your organization's Azure Active Directory (Entra ID). This grants the application the delegated permissions needed to call the SharePoint Graph API on behalf of the signed-in user. The process takes approximately five minutes.

### Step 1 — Create the App Registration

1. Sign in to [portal.azure.com](https://portal.azure.com) as a Global Administrator or Application Administrator.
2. Navigate to **Azure Active Directory** (or **Microsoft Entra ID**) → **App registrations** → **New registration**.
3. Enter a display name such as `SharePoint Smart Copy`.
4. Under **Supported account types**, choose **Accounts in this organizational directory only** (or the option that matches your organization's requirements).
5. Under **Redirect URI**, select **Public client/native (mobile & desktop)** from the drop-down, then enter: `http://localhost`
6. Click **Register**.

### Step 2 — Grant API Permissions

1. In your new app registration, go to **API permissions** → **Add a permission**.
2. Choose **Microsoft Graph** → **Delegated permissions**. Search for and add: `Sites.ReadWrite.All`
3. Still under **Microsoft Graph** → **Delegated permissions**. Search for and add: `Files.ReadWrite.All`
4. Click **Add a permission** again. Choose **SharePoint** → **Delegated permissions**. Search for and add: `AllSites.FullControl`
5. Click **Grant admin consent for [your organization]** and confirm. This requires the Global Administrator or Cloud Application Administrator role.

> **Note:** `AllSites.FullControl` (SharePoint delegated) is required for Migration API mode. If your organization uses only Enhanced REST mode, this permission can be omitted — but the app will display an error if Migration API is selected without it.

### Step 3 — Copy Your Application ID

On the **Overview** page of your new app registration, copy the **Application (client) ID** — a GUID in the format `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`. You will need this in the next step.

### Step 4 — Configure SharePoint Smart Copy

Launch SharePoint Smart Copy and click the **Settings** button (⚙) in the top-right corner of the application window.

*Settings dialog — add and manage Azure AD app registrations*

1. Click **Add** to create a new registration entry.
2. Enter a friendly **Name** for the registration, e.g. `Contoso Production`.
3. Paste the **Application (client) ID** from Step 3 into the **Client ID** field.
4. Enter your **Tenant ID** — use your organization's domain (e.g. `contoso.onmicrosoft.com`) or the tenant GUID. If unsure, enter `common`.
5. Ensure your registration is selected in the list on the left.
6. Click **Save**.

> **Note:** You can add multiple app registrations (for example, one for production and one for a dev tenant) and switch between them by selecting the active one before clicking Save.

---

## 5. Step-by-Step Guide

SharePoint Smart Copy guides you through a six-step wizard. The step indicator at the top of the window shows your current position. You can navigate back to any previous step at any time — including from the final report screen — using the **← Back** button at the bottom of the window, except while a copy is actively in progress.

### Step 1 — Connect to Source

Enter the full URL of the SharePoint site collection you want to copy from — the root URL of the site, for example:

```
https://company.sharepoint.com/sites/hr-documents
```

Click **Connect**. Your browser opens to the standard Microsoft 365 sign-in page. After signing in with an account that has Read access to the source site, the status bar confirms the connected user.

*Step 1 — Source site URL entered and connected*

Click **Next →** to proceed.

### Step 2 — Browse and Select Files

The application loads all document libraries from the source site into a tree view. Expand a library or folder by clicking the arrow to reveal its contents. Check the box next to any file, folder, or library you want to include in the copy.

*Step 2 — Browse and select files from the source site*

- Checking a folder automatically selects all files and sub-folders within it.
- Clicking a fully-checked folder again cycles it to an indeterminate state (shown with a dash) — the folder itself is deselected but its children remain selected, letting you exclude just the folder's own direct files while still copying its subfolders.
- Use **Select All** to check everything across all libraries, or **Deselect All** to clear all selections.
- File type icons indicate the format: 📝 Word, 📊 Excel/PowerPoint, 📄 PDF, and so on.
- File sizes are displayed to the right of each file name.

Click **Next →** once your selection is complete.

### Step 3 — Connect to Target

Enter the full URL of the SharePoint site you want to copy **TO**, then click **Connect**. If you already authenticated in Step 1, the connection is usually silent — no browser sign-in required.

*Step 3 — Target site connected and destination folder selected*

After connecting, the right panel shows the target site's document libraries as a folder tree. Click any folder to select it as the destination — the selected path is highlighted in blue. All selected source items will be copied into this folder.

Click **Next →** to proceed to copy options.

### Step 4 — Copy Options

Configure how the copy operation should behave:

*Step 4 — Copy options and copy preview*

- **Overwrite mode** — a three-way selector controlling what happens when a file already exists at the destination:
  - **Skip existing** — files already present at the destination are left untouched and recorded as Skipped in the report.
  - **Overwrite** — files with matching names are replaced unconditionally.
  - **If newer** — files are copied only when the source is more recently modified than the target. Files already up to date are recorded as Skipped.
- **Copy versions** — when checked, SharePoint version history is copied alongside each file. Requires versioning to be enabled on the source library.
- **Parallel copies** — controls how many files (or Migration API jobs) run simultaneously. The default of 4 is a good balance; raise to 8 or 16 on a fast connection for large batches.

When **Copy versions** is enabled, two additional controls appear:

*Step 4 — Version and copy mode options (visible when Copy versions is checked)*

- **Copy all versions** — copies the complete version history for every file.
- **Latest N versions** — copies only the N most recent versions of each file.

The **Copy mode** selector is always available in Advanced options and controls how files are transferred:

- **Migration API (Recommended)** — version numbers on the target exactly match the source (1.0, 2.0, 3.0…), with the correct Modified date and editor per version, and correct folder creation/modification dates and Created By/Modified By. With Copy Versions off, only the current version is imported via the same fast batched path. Requires Site Collection Administrator on the target site. Hover the ⓘ icon for a summary.
- **Enhanced REST** — correct dates and editors per version, but version numbers are 2× the source count (e.g. 2, 4, 6 for a 3-version file) due to a SharePoint REST constraint. No admin rights required. Hover the ⓘ icon for a summary.

> **Note:** Choose **Migration API** when exact version numbers matter or for large batches. Choose **Enhanced REST** when you do not have Site Collection Admin rights, or for small quick copies where sequential numbering is not critical.

The right panel shows a preview of the top-level items to be copied. Click **Start Copy →** to begin.

### Step 5 — Copy in Progress

The progress bar shows the overall percentage complete. The counter shows completed / total files and the elapsed time updates every 400 milliseconds.

*Step 5 — Copy in progress with real-time file status*

Each file appears in the list with its current status:

| Icon | Status | Description |
|------|--------|-------------|
| ⏳ | Pending | Waiting in the transfer queue |
| ⟳ | Copying | File transfer is in progress |
| ✅ | Success | Copied successfully (version count shown when version copying is enabled) |
| ❌ | Failed | An error occurred; the reason appears in the Error column |
| ⏭ | Skipped | File already exists at the destination and Overwrite was not enabled |

Use the **filter chips** (All / Success / Failed / Skipped) above the file list to focus on a subset of results — useful when copying thousands of files.

Click **Cancel** to stop the copy at any time. Files already transferred remain in the destination — the operation is not rolled back. When all files are processed, the **Next →** button becomes active.

### Step 6 — Copy Report

The report screen summarises the completed copy with four summary cards:

*Step 6 — Copy report with summary cards and per-file results*

| Card | Meaning |
|------|---------|
| Green — **Succeeded** | Number of files copied successfully |
| Red — **Failed** | Number of files that could not be copied (see the Error column for details) |
| Yellow — **Skipped** | Files not copied because they already existed at the destination |
| **Duration** | Total elapsed time for the entire copy operation |

The full per-file results table is shown below the summary cards. Use the **filter chips** (All / Success / Failed / Skipped) to narrow the list. Available actions:

- **← Back** — returns to Step 5 (progress screen) to review the run details, or navigate further back through the wizard
- **Export CSV** — saves the complete report to a comma-separated file
- **Start New Copy** — begins another copy operation; your sign-in session is reused automatically

---

## 6. Copy Modes — Detailed Reference

### Choosing a Mode

The **Copy mode** option appears on the Copy Options screen when **Copy versions** is enabled.

| Scenario | Recommended Mode |
|----------|-----------------|
| Large batches (50+ files or 200+ versions) | Migration API |
| Exact version numbers required | Migration API |
| Full migration fidelity needed | Migration API |
| Small batches or quick one-off copies | Enhanced REST |
| No Site Collection Admin rights | Enhanced REST |
| Per-file progress in real time | Enhanced REST |

### Migration API

Uses SharePoint's built-in Migration Import API (SPMI). Files are packaged client-side, uploaded to SP-provisioned Azure Blob containers, then imported server-side by SharePoint.

**Advantages:**

- Version numbers on target exactly match source (1.0, 2.0, 3.0, …)
- Modified By and Modified date correct per version in history
- Author and Created date preserved on the file
- Bypasses per-item throttling — SP processes the batch as a single job
- Scales well: 500 files with 10 versions each has roughly the same client-side overhead as 50 files

**Limitations:**

- Minimum 1–2 minutes of overhead per run regardless of file count (container provisioning, manifest packaging, blob upload, SP processing)
- No per-file progress during SP's processing phase — results appear only after the full job completes
- Error reporting is at the job level; individual file failures may have limited detail
- Requires `AllSites.FullControl` delegated permission and Site Collection Administrator membership on the target site
- Any single file version larger than 2 GB cannot be copied in this mode (each version is buffered fully in memory for encryption) — use Enhanced REST for files with a version this large

### Enhanced REST

Uses the SharePoint REST and Microsoft Graph APIs directly. Each file version is uploaded individually, with metadata and timestamps patched immediately after.

**Advantages:**

- Results appear per file as each one completes — you see progress in real time
- Low overhead for small batches: a 5-file copy completes in seconds
- No elevated permissions required beyond standard contributor access
- Per-file error messages are clear and immediate

**Limitations:**

- Version numbers are 2× the source count (e.g. versions 2, 4, 6 for a 3-version source file) — a SharePoint REST constraint; the correct dates and editors are still preserved
- Subject to SharePoint throttling (HTTP 429) on large batches with high parallelism
- Slower than Migration API for large migrations with many versions

---

## 7. Copy History

Every completed copy run is saved automatically to the local history. Access it by clicking the **History** button in the top-right corner of the application window.

*Copy History — list of previous runs*

The left panel lists previous runs in reverse chronological order. Each entry shows the source and target site URLs, a summary of file counts (succeeded / failed / skipped), and the total duration. Click any run to select it and view its details.

*Copy History — run selected, showing per-file details*

With a run selected, the right panel shows the summary cards and the complete per-file results table for that run. Available actions:

- **Export CSV** — saves the selected run's per-file report to a comma-separated file
- **Verify** — runs an independent Verification Report for the selected run (see below)
- **Delete Run** — permanently removes the selected run from the history
- **Close** — returns to the main application window

> **Note:** History is capped at 50 entries. When the limit is reached, the oldest entries are automatically pruned. All history is stored locally in `%AppData%\SharePointSmartCopy\Reports\` — it is never uploaded or shared.

### Verification Report

The copy report itself only reflects what the app *attempted* — the Verification Report independently confirms what is actually present, by re-scanning both the source and target with fresh Graph API calls rather than reusing any data collected during the copy.

Select a run in History and click **Verify**. You will be prompted for a location to save an Excel (`.xlsx`) workbook, then the app re-walks the source and target folder trees. This can take some time on large libraries — the status line shows files found on each side as the scan progresses, and shows a throttle notice if the Microsoft Graph API is rate-limiting the scan, so a slow scan is never mistaken for a hang. Click **Cancel** at any point to stop the verification.

The resulting workbook contains:

| Sheet | Contents |
|---|---|
| **Overview** | A one-glance headline (✓ all files match, or ⚠ content does not match) followed by summary counts — matched, content mismatch, date mismatch, only in source, only in target |
| **Source** | Every file found on the source side, with its relative path |
| **Target** | Every file found on the target side, with its relative path |
| **Comparison** | Every relative path with its match status (Match, Content Mismatch, Date Mismatch, Only in Source, Only in Target, or Unverified), plus the Source Value/Target Value that were compared and a Note explaining anything Deep Verify found |
| **Scan Errors** | Present only if a source or target root could not be scanned (e.g. it was deleted or renamed since the copy) |

> **Note:** Whether a file went missing is confirmed by relative path (does a file with the same name and location exist on both sides). Whether its *content* actually matches uses a different signal depending on file type, because a single approach doesn't work for everything:
>
> - **Most file types** (PDFs, images, archives, and other non-Office formats) get a genuine content comparison — Microsoft Graph's content hash for the source and target files must match exactly. A difference is reported as **Content Mismatch**.
> - **Office and Outlook files** — modern Word/Excel/PowerPoint/Visio formats (`.docx`, `.xlsx`, `.xlsb`, `.pptx`, `.vsdx`, and their template/add-in variants) and legacy binary formats (`.doc`, `.xls`, `.ppt`, `.vsd`, `.pub`, `.mpp`, `.msg`) — can't rely on a hash comparison. SharePoint routinely re-serializes these files' internal container (the ZIP structure behind modern formats, or the OLE metadata streams behind legacy ones) for indexing, thumbnails, and co-authoring, which changes the file's size and hash without changing its actual content. If the hash happens to match anyway, that's still trusted as a genuine match; otherwise these files fall back to a **modified date** comparison (within a few seconds' tolerance) — the app is already responsible for preserving the source's modified date onto the target, so a mismatch here means that didn't happen. This is reported as **Date Mismatch**. If neither a hash nor a modified date is available on both sides, the file is reported as **Unverified** rather than assumed to match.
>
> **Preserve Metadata must have been enabled on the original copy** for the Office-file date check to be meaningful — if it was off, the target's date was never set to match the source, and a Date Mismatch does not necessarily indicate a real problem.
>
> Verification can only be run for saved reports that recorded their source/target scope. Runs from before this feature was added do not have that information, and the **Verify** button is disabled for them.

**Deep verify Office files** (optional checkbox next to Verify): the date-based check above proves SharePoint preserved the modified date, not that the file's actual content matches — a rewritten-but-corrupt Office file could still pass. Checking this box downloads both copies of every modern Word/Excel/PowerPoint/Visio file (`.docx`/`.xlsx`/`.pptx`/`.vsdx` and variants) whose hash or date disagreed, and compares their actual internal content, ignoring only the specific parts SharePoint itself rewrites (Document ID stamps, custom document properties, and similar metadata). A file that's byte-identical in content is reported as an ordinary **Match** — with a Note on the Comparison tab explaining it was confirmed by deep verify — even though its raw hash differed. This is slower — it downloads files rather than just listing them — so it only runs on files that need it, but every flagged file is always checked; there's no cap or time limit that would leave some unverified. Off by default; legacy binary Office formats and non-Office files are unaffected by this option.

---

## 8. Troubleshooting

### Cannot connect to source or target site

- Verify the URL is the full root URL of the site collection, e.g. `https://company.sharepoint.com/sites/sitename` — not a document library URL.
- Confirm the signed-in account has at least Read access (Visitor permission level) on the site.
- Check that admin consent has been granted for `Sites.ReadWrite.All`, `Files.ReadWrite.All`, and `AllSites.FullControl` in the Azure portal.
- If the browser shows **Need admin approval**, ask your Azure AD administrator to grant tenant-wide consent via **Azure Active Directory** → **Enterprise Applications**.

### Migration API — Permission Denied

- Ensure `AllSites.FullControl` delegated permission is added to the Azure AD app registration and admin consent has been granted.
- Ensure the signed-in account is listed in **Site Settings** → **Site Collection Administrators** on the target site. Tenant-level SharePoint Admin or Global Admin roles are not sufficient on their own.
- If you cannot obtain these permissions, switch to **Enhanced REST** mode — it does not require admin rights.

### Files are failing with errors

- **File too large** — SharePoint Online supports files up to 250 GB. Files above this limit cannot be copied.
- **"File version is X GB — larger than the 2 GB this mode can buffer" (Migration API mode)** — a single version of the file exceeds what this mode can hold in memory during encryption. Switch to **Enhanced REST** mode for that file (or the whole batch).
- **File is checked out** — check the file back in at the source site before retrying.
- **Permission denied** — verify the signed-in account has Contribute or higher access on the target folder.
- **Path too long** — SharePoint Online has a 400-character limit on the full URL path. Shorten folder names or use a shallower structure.
- **Throttling (429 Too Many Requests)** — the Microsoft Graph API rate-limits heavy usage. The application retries automatically, but very large batches may take longer.
- **"SharePoint aborted the batch after N name conflicts" in the activity log (Migration API mode)** — this is expected automatic recovery, not a failure. SharePoint's Migration API cancels an entire import batch once enough files at the destination already exist, which would otherwise discard every other valid file in that batch. With **Skip existing** selected, the app marks the conflicting files Skipped (matching what Skip means) and resubmits the rest once. With **Overwrite** or **If newer**, the app clears the conflicting targets and resubmits the whole batch once. If a retry still fails, the affected files are reported individually in the final results — re-running the copy again typically clears them, since a fresh pre-flight scan re-detects the current state of the target.

### Version history not copying

- Versioning must be enabled on the source document library (**Library Settings** → **Versioning Settings**).
- The signed-in account needs **Edit** or **Design** permission on the source library to read version history via the Graph API.
- If version copying repeatedly fails for specific files, try reducing **Latest N versions** to a small value (e.g. 3) instead of copying all versions.

### Slow copy performance

- Increase the **Parallel copies** slider on the Options screen. Values of 8 or 16 significantly improve throughput on fast connections.
- For very large batches (thousands of files), **Migration API** mode typically outperforms Enhanced REST significantly.
- SharePoint throttles requests during peak hours. Running large operations outside business hours typically yields better performance.

### Token expired or re-authentication required

Access tokens are valid for approximately one hour. If the application has been idle, click **Connect** on Step 1 — MSAL will attempt a silent token refresh. If silent refresh fails, your browser will open for re-authentication. Previous selections and settings are preserved.

### Settings not persisted

Settings are written to `%AppData%\SharePointSmartCopy\settings.json`. On managed corporate machines, Group Policy may restrict writes to the AppData folder. Contact your IT department to verify write access to the `AppData\Roaming` location for your user profile.
