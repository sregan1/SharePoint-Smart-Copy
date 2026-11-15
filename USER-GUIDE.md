# SharePoint Smart Copy — User Guide

*Copy files and folders between SharePoint Online sites*

**Version 4.1.4  ·  November 2026**

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

SharePoint Smart Copy is a Windows desktop application that copies files and folders between SharePoint Online site collections, preserving full version history and metadata. Whether you are migrating content to a new site, archiving documents, or reorganizing your SharePoint environment, a guided wizard handles version history, metadata preservation, and parallel transfers — no scripting or PowerShell required.

Two copy engines are available. **Migration API** mode uses SharePoint's server-side import pipeline to achieve exact version fidelity — the same version numbers, dates, and editors appear on the target as on the source. **Enhanced REST** mode calls the Graph API directly for each version, preserving dates and editors but producing version numbers that are 2× the source count due to a SharePoint REST constraint. Both modes authenticate through your existing Microsoft 365 credentials.

### Key Capabilities

- Copy individual files, entire folders, complete document libraries, entire sites, or modern pages
- Preserve version history — copy all versions or limit to the N most recent
- **Migration API mode** — exact version numbers, exact dates, exact editors (requires Site Collection Admin on target)
- **Enhanced REST mode** — correct dates and editors per version; no admin rights required
- **Copy-if-newer** incremental mode — only copies files that are newer on the source, skipping files already up to date
- **Person/User, Managed Metadata, Lookup, Hyperlink, and Currency columns** copied alongside content in both copy modes — no manual field re-entry
- Custom column mapping dialog — map source columns to matching target columns or create missing ones
- **Custom Fields status column** flags any column value that couldn't be written (e.g. a target column that looks identical but doesn't actually match by internal name), so a partial metadata mismatch is never mistaken for a fully clean copy
- Automatic HTTP 429 throttle handling — large jobs complete without interruption
- Parallel transfers with 1–16 simultaneous file copies for faster bulk operations
- Real-time progress monitoring with per-file status updates (Enhanced REST) or job-level results (Migration API)
- Copy report with succeeded, failed, and skipped counts, inline permission status, and CSV export
- **Total Size** shown alongside item counts on the preview, in-progress, and report screens, in the Excel verification report, and in History
- **OneNote notebooks** copy correctly as a single native operation, preserving their special notebook association instead of being rebuilt as a plain folder of files
- **Structure of new lists and libraries** — views, list settings, column rules and formatting, content types (with items, files and folders put on their type), the library document template and folder default values, and, in Site scope, site columns, site content types, navigation and site settings (see *What a Library, List or Site copy brings along* in Section 5)
- **Structure to copy** — on the Options step, one checkbox per kind of structure (views, list settings, content types, column rules, templates, and in Site scope site columns, navigation and settings), and the report records what a run copied
- **Update existing lists** (optional, off by default) — a list or library that already exists on the target can be brought in line with the source: same-named views (and the default view), list settings and column rules and formatting are replaced, and the results mark the row as updated instead of skipped
- **Flows** — rewrite exported packages (zip, including solution exports) or copy flows live between Power Automate environments, with a report of everything that was changed, left out or needs a look
- **Lookups and managed metadata** — lookup values are matched by the original item, the lists a lookup points at can be copied along, and a managed-metadata term set can be copied to another tenant through Microsoft Graph (optional `TermStore.ReadWrite.All` permission)
- **Verification Report** — independently re-scans source and target after a run and produces an Excel workbook confirming what actually matches
- **Compare** — generate the same kind of difference report for any source and target you connect to, without needing a prior copy or a saved history entry
- **Copy between two tenants** — Multi-tenant mode signs in to source and target separately, with a user mapping table and **Check in target**; the report lists people the target couldn't find so you can map them (see Section 5)
- **Resume / run again** and **crash recovery** — a run cut short is recovered as an Interrupted History entry and can be resumed
- Full copy history stored locally — browse, re-export, resume, or delete previous runs
- System sleep is blocked automatically for the duration of a copy, metadata update, or verification
- ← Back navigation available on every step, including the final report screen
- Light, Dark, and System-follows-Windows themes
- 30 interface languages — follows your Windows display language by default; change it under Settings → Language (takes effect the next time the app starts). Translations are machine-generated and pending native-speaker review

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

You create and own the Microsoft Entra ID app registration in your organization's tenant (see [Section 4](#4-setup--azure-app-registration)). Your IT department controls which permissions are granted, who may use the application, and can revoke access at any time through the Azure portal — with no involvement from any third party.

### No External Servers or Telemetry

All data transfers occur directly between your computer and Microsoft's Graph API and SharePoint services. There are no intermediate servers, no analytics pipeline, and no telemetry. Your files never leave Microsoft's infrastructure.

### Local Storage Only

The only data written to your local machine is:

- **Application settings** (Microsoft Entra ID registration IDs) — `%AppData%\SharePointSmartCopy\settings.json`
- **Copy history reports** — `%AppData%\SharePointSmartCopy\Reports\` (filenames and status only, no file content)
- **MSAL token cache** — managed by the Microsoft Authentication Library in the Windows Credential Manager

> **Note:** The application keeps no copy of your files. Content is streamed from the source SharePoint site to the target. In **Enhanced REST** mode, files of 100 MB or more (and files whose size is not known in advance) are held in a temporary file in your Windows temp folder while they transfer. That file is readable only by your Windows account and is deleted automatically as soon as the file's transfer ends. In **Migration API** mode, file content is encrypted before it is uploaded to the Azure storage container that SharePoint provides for the import.

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
5. Click **Grant admin consent for [your organization]** and confirm. This requires the Global Administrator role (Cloud Application Administrator can add permissions but may not be able to fully consent to all of them).

> **Note:** `AllSites.FullControl` (SharePoint delegated) is required for Migration API mode. If your organization uses only Enhanced REST mode, this permission can be omitted — but the app will display an error if Migration API is selected without it.

> **Optional — only for the Forms tool:** to copy Microsoft Forms, also add **Microsoft Forms** → **Delegated permissions** (search "Microsoft Forms" under *APIs my organization uses*) and grant consent. It is not needed for any other feature. See [Copying Forms and Flows](#copying-forms-and-flows).

> **Optional — only to copy managed metadata to a different tenant:** also add **Microsoft Graph** → **Delegated permissions** → `TermStore.ReadWrite.All` and grant admin consent on the registration used for the **target** (the copy writes the term set there). In different-tenant mode the registration used for the **source** needs only the read permission `TermStore.Read.All`; with one registration (single-tenant mode) `TermStore.ReadWrite.All` covers both sides, and without `TermStore.Read.All` you are not asked for it. The signed-in target account also needs a term store role, for example *Term Store Administrator*. It is not needed for any other feature, and a same-tenant copy never uses it. See [Managed metadata](#managed-metadata--how-it-copies).

#### Which permissions does each feature need?

Every sign-in asks for the two Microsoft Graph permissions, so they are needed for anything that connects to SharePoint. The other two are only needed by the features marked below. If you do not use a feature, you can leave its permission out.

| Feature | Graph `Sites.ReadWrite.All` | Graph `Files.ReadWrite.All` | SharePoint `AllSites.FullControl` | Microsoft Forms | Graph `TermStore.ReadWrite.All` |
|---|:-:|:-:|:-:|:-:|:-:|
| Files & Folders, Libraries & Lists, Entire Site — **Enhanced REST** mode | ✔ | ✔ | — | — | optional† |
| Files & Folders, Libraries & Lists, Entire Site — **Migration API** mode | ✔ | ✔ | ✔ | — | optional† |
| Site Pages (always Enhanced REST) | ✔ | ✔ | — | — | — |
| **Compare**, and verifying a copy | ✔ | ✔ | — | — | — |
| **Forms** tool | ✔ | ✔ | — | ✔ | — |
| **Flows** tool (packages) | ✔ | ✔ | — | optional* | — |
| **Flows** tool (live copy between environments) | ✔ | ✔ | — | optional* | — ‡ |
| Browsing saved History, Settings (local only; verifying a run signs in like Compare) | — | — | — | — | — |

‡ Live copy also needs, on the app registration, **Power Automate Service** → delegated `Flows.Read.All` and `Flows.Manage.All`, and **PowerApps Service** → delegated `User` (to list the connections to bind). To create flows **inside a solution**, add **Dataverse** (shown as *Dynamics CRM* in older tenants) → delegated `user_impersonation`; without it the flows are created outside any solution. In an environment with Dataverse the signed-in account also needs a Dataverse security role that can create flows and read connection references (for example *Environment Maker* or *System Administrator*); a missing role is reported by name.

\* The Flows tool only uses Microsoft Forms to look up form titles and the target account's forms. Without the permission those lookups report an error and you paste each new form's ID or edit-page link instead; everything else in Flows works.

† Only for copying a managed-metadata column to a different tenant (the term set is copied through Microsoft Graph, delegated, admin consent). The signed-in target account also needs a term store role, for example *Term Store Administrator*. Without the permission the column is not created and the run report says why. See *Managed metadata* in the User Guide.

For example, a company that only wants to move Power Automate flows needs the two Graph permissions and nothing else: no Microsoft Forms permission, and no `AllSites.FullControl`. Flows reads your packages from disk and looks up lists and columns on the source and target sites with the same sign-in. Only live copy (choosing flows in a Power Automate environment instead of exporting packages) needs the extra Power Automate and PowerApps permissions marked ‡.

> **Important — admin consent is per-tenant:** Consenting to these permissions in one Microsoft 365 tenant has no effect on any other tenant. If you're deploying to a new customer or a second tenant, repeat this entire Azure App Registration section — a new app registration, new permissions, and a fresh **Grant admin consent** click — inside *that* tenant. Consent can never be reused across tenants.

### Step 3 — Copy Your Application ID

On the **Overview** page of your new app registration, copy the **Application (client) ID** — a GUID in the format `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`. You will need this in the next step.

### Step 4 — Configure SharePoint Smart Copy

Launch SharePoint Smart Copy and click the **Settings** button (⚙) in the top-right corner of the application window.

*Settings dialog — add and manage Microsoft Entra ID app registrations*

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

**Disconnect** also signs the account out, so the next **Connect** opens the Microsoft sign-in again with the account picker instead of silently reusing the remembered account. (In Single tenant mode the target shares that sign-in, so it is disconnected too.)

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
- **Parallel copies** — controls how many files (or Migration API jobs) run simultaneously. The default of 8 is a good balance; raise to 8 or 16 on a fast connection for large batches.
- **Preserve metadata** — when checked, folder Created/Modified dates, Author/Editor, and custom folder color are copied to the target. The **Re-apply folder metadata every run** option underneath it (on by default) controls whether that repair runs for every folder on every run, or only for folders receiving new files this run — turning it off speeds up repeated incremental runs on large libraries where most files are already up to date.

When **Copy versions** is enabled, two additional controls appear:

*Step 4 — Version and copy mode options (visible when Copy versions is checked)*

- **Copy all versions** — copies the complete version history for every file.
- **Latest N versions** — copies only the N most recent versions of each file.

*List item version history* (Library and Site scopes) follows the same options: with **Copy version history** off, list items copy with their current values only; otherwise they copy all versions or the latest N, exactly as files do. There is no separate setting for list items.

History is replayed for items the run creates: the item is created from its oldest version, then each later version is written with its own values, Modified date and editor, so the target shows the same version labels (1.0, 2.0, 3.0…). Items that already exist on the target keep their own history and are updated in place as before. Text, number, choice, yes/no and date columns follow the history; people, lookup, taxonomy and link columns copy their current value only. If a version cannot be written, the item row says which part did not copy.

The **Copy mode** selector is always available in Advanced options and controls how files are transferred:

- **Migration API (Recommended)** — version numbers on the target exactly match the source (1.0, 2.0, 3.0…), with the correct Modified date and editor per version, and correct folder creation/modification dates and Created By/Modified By. With Copy Versions off, only the current version is imported via the same fast batched path. Requires Site Collection Administrator on the target site. Hover the ⓘ icon for a summary.
- **Enhanced REST** — correct dates and editors per version, but version numbers are 2× the source count (e.g. 2, 4, 6 for a 3-version file) due to a SharePoint REST constraint. No admin rights required. Hover the ⓘ icon for a summary.

> **Note:** Choose **Migration API** when exact version numbers matter or for large batches. Choose **Enhanced REST** when you do not have Site Collection Admin rights, or for small quick copies where sequential numbering is not critical.

The right panel shows a preview of the top-level items to be copied. Click **Start Copy →** to begin.

### Copying only the structure (no files or items)

In Library and Site scope, untick **Copy files and list items (content)** at the top of the Options step. A line under it confirms *Structure only*: the copy then creates the libraries and lists with their columns, views, settings and content types (and, in Site scope, the site columns, site content types, navigation and site settings), and copies no files or list items. Tick it again to copy the content as well. A library or list that already exists on the target only gets the content types and views it lacks.

### Choosing what structure is copied

In Library and Site scope the Options step has a **Structure to copy** section (open by default, above *Advanced options*). Each checkbox is one kind of structure, all on by default:

![The Options step of a Libraries and Lists copy with the Structure to copy section](SharePointSmartCopy/Docs/screenshots/04_options_library.png)

| Checkbox | What it covers |
|---|---|
| **Views** | Shared views, including calendar, Gantt, datasheet and board views and the default view |
| **List settings** | Content approval, check-out, attachments, folder creation, item-level security and the list's own validation rule |
| **Content types and their document templates** | Custom content types on the list, their columns and templates, and the type of copied items, files and folders |
| **Column validation rules and formatting** | A column's validation formula and message and its JSON formatting |
| **Library template and folder default values** | The library's document template and its folder default values |
| **Update existing lists to match the source** | Off by default. For lists and libraries that already exist on the target: replaces same-named views and the default view, list settings, and column rules and formatting (see below) |
| **Site columns and site content types** (Site scope) | The site-wide definitions the lists build on |
| **Navigation** (Site scope) | Quick Launch and top navigation |
| **Site settings** (Site scope) | Description, quick launch, tree view and the other settings that can be written |
| **Also copy lists that my lists look up** (Library scope) | Adds the custom lists that a selected list's Lookup columns point at |
| **Copy managed-metadata term sets** (different-tenant mode) | Copies a managed-metadata column's term set to the target term store |

Columns themselves, projected lookup fields and the data are not switches: columns are always created, and *Copy files and list items (content)* above the section controls the data. The choices are remembered for the next run, and the run summary on the report (and in History) lists the structure that run copied, for example *Structure: Views, List settings, Content types*.

### What a Library, List or Site copy brings along
When a Libraries & Lists or Site copy **creates** a list or library on the target, it copies the list's structure as well as its content:

| What | How it is copied |
|---|---|
| **Columns** | Created on the list, with your column mapping applied. Column validation rules (formula and message) and column formatting are copied onto the matching column |
| **Views** | Shared views are recreated: filter, sorting, grouping, columns and their order, row limit, paging, totals, formatting, and the default view. A view with the same name as one the new list already has (its built-in *All Items* / *All Documents*) is updated instead of duplicated. Calendar, Gantt, datasheet and modern board views are created with their own type. Personal and hidden views are skipped |
| **List settings** | Content approval, check-out, attachments (lists), open in browser (libraries), folder creation, search visibility, draft version visibility, item-level security, and the list's own validation rule |
| **Content types** | Custom content types are created on the target site (SharePoint derives every type created this way from *Item*) and added to the list; the list's columns that the source type carries are recreated so they sit on the new type. Copied list items, files and folders are put on the matching type, a file only when the type carries every column being written to it |
| **Document template** | The library's template file is copied into the target's Forms folder and set as its template |
| **Folder default values** | The library's folder default column values, with their folder paths pointed at the target library, and the rule that applies them to new files |

**Site scope** also copies the site-wide definitions the lists build on, before any list is created: custom **site columns**, custom **site content types** (without their own columns, which cannot be linked over REST), the **left and top navigation** (links whose title already exists on the target are left alone, and addresses into the source site point at the target site), and **site settings** (description, quick launch, tree view, horizontal quick launch, mega menu and minimal download options). The site title, address, time zone and regional settings are not changed. Site permissions use the existing **Copy permissions** option. A *Site columns, content types, navigation and settings* row in the results lists anything that could not be copied.

What to know:

- Everything above is copied to lists and libraries the copy **creates**. A list or library that already exists on the target only gets what it lacks: the content types it does not have and the views whose title it does not have are added. Its existing views, default view, settings, column rules, template and default values are not changed, unless you tick **Update existing lists to match the source**: then a view of the same name is rewritten to the source's definition (filter, sort, grouping, columns, row limit, paging; a view with no filter on the source loses the target's filter), the source's default view becomes the default, and list settings and column rules and formatting are applied again. Content types stay add-only (recreating a column on a list that has data would delete its values), and so do the library template and folder defaults; views that exist only on the target are left alone. A row that was updated is shown as **Success** with the note *Existing list updated to match the source*. The option only applies when custom columns are being copied. A view's filter, sort, grouping and totals that name a column the target does not have are left out of the view, with a note, so the view still opens (SharePoint otherwise accepts it and then fails to show it: *One or more field types are not installed properly*).
- A custom content type's own **document template** (the file new documents of that type start from) is copied into the new library's Forms folder and set on the library's copy of the type. It is not copied for the site-level type in Site scope (a site content type's own folder cannot be written over REST).
- Anything that cannot be copied is listed on that item's result row (for example *views, settings or content types: …*), and the rest of the copy carries on. A view's missing columns are left out of the view, and a filter that names a missing column is left out so the view is still created.
- **Content types on files and folders** are written together with the custom column values, so files get their type only when custom columns are being copied and the library has at least one custom column (custom columns are always copied in Library and Site scope). Only custom types are carried; ordinary Documents and Folders need no extra write. A folder's type is set before its dates and authors are re-applied.
- **Document sets** are folders on a special content type. The copy sets a copied folder to the type when the target list has it, which needs the *Document Sets* feature turned on for the target site collection. It does not create the document set's welcome page or shared columns.
- **Lookup site columns** are created after the lists, bound to the target list of the same title; a lookup into another web of the site collection needs the same web path on the target, and the extra columns a lookup shows from its list are added too. A column whose list is not on the target is listed in the results. Managed-metadata site columns are not copied in a cross-tenant copy (list columns are handled by the term set copy, see *Managed metadata*).
- **List settings are applied in two steps:** most when the list is created, but content approval, require check-out and draft visibility only after all files and the metadata passes are done, so they cannot interfere with the copy's own writes.
- A custom column on a copied content type is recreated on the new (still empty) list so it sits on the type. A managed-metadata column, a projected (extra) lookup field, or any column on a list that already holds data, is listed in the report instead.
- Modern **board** views are created and set to the board type, but whether they render as a board has not been proven. Document Set types need the feature on the target site collection.
- The links under the *Recent* navigation node are generated by SharePoint and are not copied.

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
| 🚫 | Cancelled | Still in progress when the run was stopped or the app was closed — not attempted, and not counted as a failure |

Use the **filter chips** (All / Success / Failed / Skipped) above the file list to focus on a subset of results — useful when copying thousands of files. Cancelled items appear under the **Failed** chip, since they're still unfinished work, but they're counted separately in the summary so an interrupted run doesn't look like a mass failure.

Click **Cancel** to stop the copy at any time. Files already transferred remain in the destination — the operation is not rolled back. When all files are processed, the **Next →** button becomes active.

### Step 6 — Copy Report

The report screen summarizes the completed copy with four summary cards (a fifth, **Cancelled**, appears only if the run was stopped early):

*Step 6 — Copy report with summary cards and per-file results*

| Card | Meaning |
|------|---------|
| Green — **Succeeded** | Number of files copied successfully |
| Red — **Failed** | Number of files that could not be copied (see the Error column for details) |
| Yellow — **Skipped** | Files not copied because they already existed at the destination |
| **Duration** | Total elapsed time for the entire copy operation |
| **Cancelled** *(hidden when zero)* | Items still in progress when the run was stopped or the app was closed — not attempted, so not counted as Failed |

The full per-file results table is shown below the summary cards. Use the **filter chips** (All / Success / Failed / Skipped) to narrow the list. Available actions:

- **← Back** — returns to Step 5 (progress screen) to review the run details, or navigate further back through the wizard
- **Export CSV** — saves the complete report to a comma-separated file. The suggested filename is prefixed with the source and target site names (e.g. `Marketing-Archive-CopyReport_Files_....csv`); turn this off in **Settings** if you'd rather not have it
- **Start New Copy** — begins another copy operation; your sign-in session is reused automatically
- **Compare source & target** and **Open History** — open the Compare tool and the run history without leaving the report
- **Failed / Skipped cards** — click a card to filter the list to just those items

### Copying Between Two Tenants (Multi-tenant Mode)

On the first screen, choose **Multi-tenant** next to **Connection** when the source and the target are in **different** Microsoft 365 tenants. In the default **Single tenant** mode both sites share one sign-in, exactly as described above.

> **Note:** Multi-tenant mode has been tested against two tenant pairs, but other setups (guest accounts, conditional access, very large libraries) may behave differently. Try it on a small library first and check the result with a Verification Report, and please report anything unexpected on GitHub.

1. **Create an app registration in each tenant** (Section 4) and add both in **Settings**. Each tenant's administrator must grant consent for its own registration.
2. On the first screen, choose the **Source app registration** and the **Target app registration**. You sign in to each site separately, with an account from that tenant. The title bar shows both accounts (*Source: … · Target: …*).

   ![The first screen in Multi-tenant mode, with a separate app registration for each side](SharePointSmartCopy/Docs/screenshots/10_multitenant_source.png)
3. Click **Identity mapping…** (on the first screen or the Target screen) to say who is who in the two tenants:
   - **Source / Target email domain** — a shortcut: an address such as `jane@contoso.com` is written to the target as `jane@fabrikam.com`. Guest accounts (`…#EXT#…`) are not rewritten.
   - **The table** — one row per person or group: the source user on the left, the matching target user on the right (users by email address, groups by name). A row always wins over the domain swap. Use **Add** and **Remove**, **Import CSV…** / **Export CSV…**, and **Fill targets from domain swap** to propose a target for rows you have left empty. Rows with a problem (a blank cell, a duplicate source, source equal to target) are highlighted and must be fixed before you can save.
   - **Check in target** looks each target email up in the target tenant and shows **✓ Found** or **✗ Not found** next to it. It needs the target site's URL (enter it on the Target screen first). Looking a user up also adds them to the target site's user list, which the copy does anyway before it can make them an author or give them access.

   ![The identity mapping table with Check in target results](SharePointSmartCopy/Docs/screenshots/11_identity_mapping.png)
4. Authors, editors, person columns and permissions are written using that mapping. Anyone who cannot be found in the target is left out — not given access, not set as author or editor, dropped from person columns — and recorded.
5. When a run finishes and anyone could not be found, the report shows a banner with **Map these users…**, which opens the table already filled in with those people (with a suggested target from the domain swap where possible). **History → Map unresolved users…** does the same for an earlier run. Saving a mapping applies to the **next** run: Copy if newer skips files that are already up to date, so the authors on files that were already copied do not change. Run again with **Overwrite** to re-stamp them. Permissions are refreshed on skipped files either way.

   ![The report banner listing people the target could not find](SharePointSmartCopy/Docs/screenshots/13_unresolved_report.png)

   ![The mapping table opened with those people already added](SharePointSmartCopy/Docs/screenshots/12_identity_prefill.png)

What differs from a same-tenant copy:

- **Managed-metadata (taxonomy) columns and their values are not copied** — term IDs only exist in the tenant that owns the term store.
- **OneNote notebooks and document sets cannot be copied** between tenants yet; they are reported as failed with an explanation.
- **Entra groups do not carry over** unless you map them with an override.
- The source and target must be on **different SharePoint hosts**. For two sites in the same tenant, use Single tenant mode.
- Switching between Single tenant and Multi-tenant disconnects both sites (you are asked first).

### Copying Forms and Flows

On the **Browse** step, below the four copy-scope tiles, an **Other tools** row has two buttons, **Forms** and **Flows**, that open their own tools. They are not copy scopes: they use the sign-ins you already made (Single tenant: one shared sign-in; Multi-tenant: the source and target sign-ins).

> **Unofficial API.** The Forms tool talks to the Microsoft Forms web API, which Microsoft does not document or support. It works today but may change without notice. Try a few forms first.

**Forms** — four steps, like the main copy wizard:

1. **Forms** — your forms load as soon as the dialog opens (you may be asked to sign in); **Reload forms** reads them again, and **Load my forms** is the retry if the first attempt failed or was cancelled. Your own forms are listed; forms owned by Microsoft 365 groups and forms shared with you are not listed yet. Tick the forms to copy; the **filter box** narrows the list by title or type (**Select all** then takes just the visible forms), and the **Already copied** column shows forms you copied before. **Export responses…** (bottom left) writes one Excel file per ticked form, or one workbook with a sheet per form if **One workbook** is ticked; responses can't be re-submitted into a copy. If a selected form was already copied to the same account, you are asked before a second copy is made.

   ![Choosing forms to copy](SharePointSmartCopy/Docs/screenshots/14_forms_select.png)

2. **Target** — choose where the copies go. In *Single tenant* mode they are created in the same account, next to the originals. In *Multi-tenant* mode this step has its own **Sign in to the target account** button (separate from the SharePoint sign-ins) and confirms that the account can use Microsoft Forms before you can continue. You can add a suffix to the copied titles here.

   ![The Forms Target step](SharePointSmartCopy/Docs/screenshots/15_forms_target.png)

3. **Copy** — shows progress form by form. **Cancel** stops after the current form (Cancel is also available while loading, signing in and exporting). Closing the window during a copy asks first, because it can leave a half-built form in the target.
4. **Results** — a dashboard like the file copy's report: how many forms copied, copied with warnings, or failed, how many questions copied, and how long it took. The table lists each form with the reason for any warning or failure. **Open selected copy** and **Open Microsoft Forms** jump to the copies, **Export results…** saves the table to Excel, and **Show form ID map** opens the file the Flows tool uses. **Copy more forms** starts again.

   ![Forms results with a warning](SharePointSmartCopy/Docs/screenshots/17_forms_results.png)


Each copy keeps all questions (with their original IDs), choice branching, quiz type, settings, and theme. After each copy the tool checks that the copy has as many questions as the source. Ranking options and Likert scales are copied too, and so are the form's logo, background, and header images: each is uploaded to the new form, then checked on the copy. If an image cannot be added, a copy of the file is saved in `Documents\SharePointSmartCopy\Forms images` and the warning names it, so you can add it in Microsoft Forms by hand. Quiz scoring is not verified; that is noted in the results.

**Flows** — six steps. The same layout runs through all of them: **SOURCE** (grey) is where your flows come from and **TARGET** (blue) is where they go, each with its tenant address. Every step opens with one line saying what happens, and a **What happens here** link for the longer reason. A small **What the tool is doing** strip at the bottom shows the latest action; click it to see the whole log (it opens by itself when something fails).

1. **Packages** — a choice at the top: **Copy live between environments** or **Use exported packages (.zip)**. *Live*: click **Sign in and load environments**, pick the **From** and **To** environments, tick the flows to copy (**Select all** / **Select none**; solution-aware flows are listed too), and press Next; the tool reads the flows and takes them through the same steps as packages. In Multi-tenant mode only the source loads first; **Sign in to target tenant** then signs in to the target and loads its environments. *Zip*: export your flows in Power Automate (**Export → Package (.zip)**) and add the files, or drag them onto the window. Each package is listed with the flows found in it and how many sites and lists they use; **✕** removes one. Add the **original** exports from the source, not packages this tool wrote earlier (the tool recognizes those and warns). Solution exports (.zip files with a `Workflows` folder) are read too. **Have a saved mapping? Load it…** fills the following steps from a mapping you saved earlier. **What these flows use** (closed by default) lists the services the flows use and what the tool does about each.

   ![The Flows Packages step](SharePointSmartCopy/Docs/screenshots/18_flows_packages.png)

2. **Sites** — shows where each flow's site moves to. The **Source** and **Target** lines show each tenant's address, the account it is connected as, and a state chip. In *Multi-tenant* mode, **Change source…** signs in to the source again, and **Sign in to target…** signs in to the target (the window always opens, so you can switch account or browser; the button turns into **Cancel** while it waits). The tool reads the target tenant's address from the account and fills in every empty target site with the same path on it. The table is **From** (the address in your flows), **To** (the target site, editable) and a status chip. A site left empty is not changed, and the report lists it. In *Single tenant* mode there is one sign-in for both sides, so the Target line shows the same account as the Source line; if nobody is signed in yet, **Sign in…** opens that sign-in. A flow site that is the same as the source site you connected on the Browse step starts out pointed at your target site, and **Sign in…** fills every other empty one with the same path on your tenant. Look through the **To** column and type the full address of any target that is still empty (a placeholder in the empty cell reminds you). Next stays off only while every target is empty, so a package that keeps its sites needs one typed. If a flow's sites are on a different tenant than your source connection, **Advanced: read lists from a different address** opens by itself with the corrected address. One warning at a time is shown under the connections.

   ![The Flows Sites step](SharePointSmartCopy/Docs/screenshots/19_flows_sites.png)

3. **Lists & columns** — **one card per list** the flows use, however many flows use it. A card has a header (the list's name, which target it goes to, how many flows use it, and a status chip) and opens into a **SOURCE** pane (tenant, list name and ID, site) and a **TARGET** pane (tenant, a dropdown of every list on the target site, or paste an ID). The same-named target list is chosen for you. Below is the **column mapping** from source to target: the columns the flows use (tick "Show all columns" for the rest) with the target column chosen for you. You only need to act on cards marked **Needs a look** and rows marked *Similar name* or *No match*; a row you fix stops asking for attention. A column that the flows **write** (in Create item or Update item steps) and the **target list does not have** is set to **Leave out of the flows (not on the target)** by default: those steps stop writing that field (importing a flow that writes a field the list does not have fails), the card says which columns are left out, and the report lists each one under **To review**. A column the flows only filter on, sort by or read in an expression does not stop the import, so it does not block **Next**; the row is marked and the report lists it, because that step will not find the column when the flow runs. The value that field carried is not saved anywhere on the target. Add the column to the target list and click **Check again**, or pick another column, to keep it. If you map two columns of the same step to one target column, Next stays off and the card names them (the second would overwrite the first's value, and the import would reject the field that was left unchanged). If the source list cannot be read, the card says why and still lets you map, using the column names found in the flow. If the list's ID is already a list on the target (a package this tool wrote), the card says **Already on target**. The summary line at the top ("7 lists · 5 ready · 2 need a look") always matches the cards.

   ![A list card with its column mapping](SharePointSmartCopy/Docs/screenshots/20_flows_lists.png)

**Next stays off until the flows would work.** On Sites (Multi-tenant: target signed in, every site has a target), Lists & columns (every list has a target list, every column the flows **write** exists there (or is left out), no two columns of one Create item step go to the same target column, and every column the target list requires is filled in by the flow's Create item steps, or by a value you type on the list card, which the tool adds to those steps) and Forms & email (a valid target form), the footer says what is in the way. Fix it, use **Check again** on a list card after changing the list in SharePoint, and Next turns on.

4. **Forms & email** — each form in your flows is shown with its target. A flow only stores a form's ID, form names load automatically when you arrive (**Load form names…** retries): the tool reads each form's title from Microsoft Forms (source account, including forms shared with it) and the forms the target account owns, and reports what it found on each side right under the button (or why it could not). A form whose title matches exactly one target form (or "title (copy)") is matched for you; otherwise pick the new form from the dropdown, or paste its ID or edit-page link. Either works, and **Cancel** stops a pending sign-in. A form's ID carries its tenant, and the copy must be in the **target's** tenant: across tenants that is a different tenant from the form in the flow; in Single tenant mode it is the tenant you are signed in to (a package exported from another tenant can be imported here, and then needs a target form). The dropdown lists only forms the target account owns, not group-owned forms. **Other values** (email addresses, the tenant ID, library paths) are grouped by kind. Values the tool filled in from your sign-ins are marked **✓ filled in**; **Clear values filled in automatically** undoes them. Email domains are one sentence: *Move every address on [domain] to [domain]*.

   ![The Flows Forms and email step](SharePointSmartCopy/Docs/screenshots/21_flows_forms.png)

5. **Write** — *Live mode*: the step is headed **Copy the flows to the target** and its button reads **Copy flows**. Nothing is saved as a package: each flow is created **off** (stopped) directly in the target environment, and nothing runs until you have reviewed it. **Bind each flow's connections to my existing connections** (on by default) points each connection reference at one of *your* connected connections in the target; a connection that reports any error is skipped, and the report lists which connection each reference was bound to. In an environment with Dataverse, **Create in** lists its unmanaged solutions: choose one to create the flows inside it (they are then solution-aware; the target's connection reference for each connector is used, and one is added to the solution when there is none), or leave **My flows (not in a solution)**. Copying into the same environment you read from asks first. The **Report folder** is the only thing written to your computer: an HTML page and a Markdown file listing what was changed in each flow, the columns left out, the connections matched and what still needs your review. Keep it with your migration records. If one flow fails, press **Copy flows** again: flows created earlier are skipped, not duplicated. Cancelling still writes the report with what was created. *Zip mode*: a **Ready to write?** list says where everything stands (packages, sites, lists, columns, forms, email); a line that needs a look has a **Go to step** link. Choose the output folder (it is remembered) and the name of each updated package in the **Names of the updated packages** card (type a new name per package, or use **Keep the original names** or **Add ' - migrated'**). When you click **Write updated packages**, a Save As dialog opens for each package with that name filled in; you can change the folder or name there, the tool refuses to save over an original, and Cancel writes nothing. **Save mapping…** keeps your choices as a JSON file.

   ![The Flows Write step](SharePointSmartCopy/Docs/screenshots/22_flows_write.png)

6. **Report** — what was written, how many replacements were made, and **To review**: each item the tool could not settle, in words, with a **Go to step** link where one step fixes it. **Open output folder**, **Open report**, **Open Power Automate** and **Save mapping…** are one click away, and **What to do next** is numbered. `rewrite-report.html` (also saved as `rewrite-report.md`) lists what the flows use, what to do after import, and every change.

   ![The Flows Report step](SharePointSmartCopy/Docs/screenshots/23_flows_report.png)

After importing a package into Power Automate, **recreate each connection and re-select it** — connections can't be moved between tenants. Import flows turned off, review them, then turn them on. Your original packages are never modified.

**Requirements and limits**

- Forms needs the delegated **Microsoft Forms** permission on the app registration (see Step 2 of the Azure setup). If it is missing, the dialog shows a hint.
- Legacy flow packages are fully supported. For solution exports only the flow definitions are rewritten; environment variables and connection references in the solution are not, and the report says so. Rewritten solution exports import cleanly (tested by importing one back into Dataverse). Live copy reads solution flows through the same API as other flows.
- A column that a flow **writes** and the target list does not have is left out of those steps (the import would otherwise fail on it) and listed under **To review**. A column that is only filtered on, sorted by or read in an expression does not block the copy, but that step fails when the flow runs; the report names those. Whole steps are never removed: a list with no target list, or a required column nothing fills, keeps Next off.
- Column reads inside a loop or of the trigger's item (`items('Apply_to_each')?['Column']`, `item()?['Column']`, `triggerBody()?['Column']`) are renamed like `body/Column` reads, and a list a flow names by its title is matched by that title on the target. A flow file that cannot be read is copied unchanged and flagged in the report.
- **Create flow fails with "connection has error status"** when the connection a flow is bound to in the target is broken (an expired sign-in, a changed password, a Conditional Access policy). Open **Connections** in Power Automate for the target environment, choose **Fix connection** (or create the connection again), and save the flow again.
- Forms: images on questions and on their choices are copied as links to the source form's images, so they stop showing if the source form is deleted or when the copy is in another tenant; the results note says how many questions are affected. Only forms owned by the signed-in account are listed (group-owned forms and forms only shared with you are not). **Sign in again** on the target (Multi-tenant) opens the account picker, so another account can be chosen.
- The tool edits the same values you would change by hand in a package's `definition.json` (form ID, site URL, list GUIDs, column names), and writes a new package for you to import.
- Column names inside complex flow expressions are matched by pattern; review the report.

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

- **Export CSV** — saves the selected run's per-file report to a comma-separated file, with the same source/target site name filename prefix as Step 6 (see Settings to turn it off)
- **Resume / run again** — reconnects and re-runs the selected run's selection with **Copy if newer**, so files that were already copied are skipped and only what is missing or changed is transferred. Available for runs that copied a file selection. Custom column mappings are not restored; set them again on the Options step if you used them.
- **Verify** — runs an independent Verification Report for the selected run (see below)
- **Delete Run** — permanently removes the selected run from the history
- **Close** — returns to the main application window

A run that was cut short — the app was closed, crashed or lost power mid-copy — is recovered the next time you start the app and appears here marked **INTERRUPTED**, with the items that had finished. Use **Resume / run again** to carry on.

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

### Compare (Source vs. Target)

On the **Browse** step, click **Compare** in the **Other tools** row (below the copy-scope tiles) to generate a difference report between any two locations — without needing a prior copy or a saved History entry first. This is useful for checking whether a site copied by another means (or edited independently on both sides since a copy) still matches, or for a one-off audit that was never meant to be a full copy job.

*Compare — connect to any source and target and pick a library or folder on each side*

Connect a **Source** and a **Target** independently, the same way you would in the main wizard, then select a library or folder in each side's tree. Click **Compare with Excel Report** to scan and save the resulting workbook — it's produced by the same engine and has the same sheet layout (Overview/Source/Target/Comparison/Scan Errors) as the Verification Report above, including the optional **Deep verify Office files** checkbox. A Compare run is never saved to History.

---

## 8. Troubleshooting

### Switching to another tenant still signs in to the first one

In **Single tenant** mode one sign-in serves every site. If you change the source URL to a site on a different tenant, the app now clears the sign-in and prompts again. If the prompt still offers (or forces) the first tenant, look at the card at the top of the Connect step: its **Sign in with app registration** dropdown shows (and lets you change) the registration the sign-in uses, and **Show more** under it names that registration's tenant and explains how to connect a site on another one: a registration with a specific **Tenant ID** can only sign in to that directory, so another tenant needs either a multi-tenant registration (leave Tenant ID empty to use `common`) or **Multi-tenant** mode, which signs in to the source and target separately. Closing and reopening the app also clears the sign-in.

### A copied view says "One or more field types are not installed properly"

The view's filter, sort, grouping or totals name a column the target list does not have, or the column exists under a different internal name (a column created by an older copy was named after its display title, for example `Assigned_x0020_to` instead of `AssignedTo`). New copies keep the source's internal names, and when a view is copied the tool matches such a column by its display title and leaves out anything that still has no column. To repair a list that was copied earlier, copy it again with **Update existing lists to match the source** ticked; the results list each view it rewrote or trimmed.

### Cannot connect to source or target site

- Verify the URL is the full root URL of the site collection, e.g. `https://company.sharepoint.com/sites/sitename` — not a document library URL.
- Confirm the signed-in account has at least Read access (Visitor permission level) on the site.
- Check that admin consent has been granted for `Sites.ReadWrite.All`, `Files.ReadWrite.All`, and `AllSites.FullControl` in the Azure portal.
- If the browser shows **Need admin approval** / **Approval required**, admin consent is missing or incomplete for *this specific tenant*. Have a Global Administrator for that tenant do the following:
  - Confirm they're signed in as **Global Administrator** — Cloud Application Administrator or Application Administrator may not be able to consent to all of these permissions.
  - Open the app registration's **API permissions** page and verify every row — both the Microsoft Graph block and the SharePoint block — shows a green "Granted for [tenant]" status. If a permission was added after consent was last granted, click **Grant admin consent** again to cover it.
  - As a one-click alternative, visit `https://login.microsoftonline.com/{tenant-id}/adminconsent?client_id={app-client-id}` while signed in as Global Admin — this grants tenant-wide consent directly without navigating the portal.
- **Admin consent does not transfer between tenants.** If this is a new Microsoft 365 tenant, repeat the full Azure App Registration setup in Section 4 — including a fresh admin consent grant — inside that tenant; consent granted for another tenant's app registration has no effect here.

### Migration API — Permission Denied

- Ensure `AllSites.FullControl` delegated permission is added to the Microsoft Entra ID app registration and admin consent has been granted.
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
- **A file shows "⚠ Warning" in the Custom Fields column** — the file itself copied successfully; only one or more custom column values could not be written. Check the Custom Fields Details column for the reason. The most common cause is a target column that looks identical to its source counterpart (same display name and values elsewhere) but doesn't actually share the same internal name — SharePoint auto-suffixes a column's internal name at creation time if it collides with another field, so "Notes" on the source and "Notes" on the target can silently be different fields underneath. Open the Column Mapping dialog and confirm the source column is mapped to the target's real internal name (visible in **Library Settings** → click the column → the `Field=` value in the URL), rather than relying on the display name to match automatically.

### Lookup columns — how they copy and where they can't

A Lookup column stores a pointer to an item in another list, and item IDs differ between source and target. So the copy writes the matching **target item's ID**. For an item copied in the same run, the match is made by the **original item** (so two items with the same title are told apart); for an item that was already at the target, the copy falls back to the column's **display text** (its "show field", usually Title). The run report ends with one *Lookup values* line: how many values were matched by item, by text, and not found. This works in both copy modes and between tenants.

In **Library** scope, the option **Also copy lists that my lists look up** (on by default) adds the custom lists that a selected list's Lookup columns point at, repeating until nothing new is needed, so the column can be created and its values resolved in one run. A lookup to a list in a **child web** of the same site collection is recreated against the same-named list in the same child web at the target (the child web must already exist there), and its values are matched against that child web. **Projected (additional) fields** — the extra columns a lookup shows from the other list — are recreated on a new list once its lookup column exists. The Column mapping dialog shows under each lookup column which list it looks up, and under each managed-metadata column whether its term set will be copied.

For a value to come across:

- The target must have a **Lookup column** of that name, bound to a list that has an item with the same display text. When a library, list or site is copied, the column is recreated and bound to the target list with the **same title** as the source's lookup list. In **Files** scope, "Create in target" does the same.
- The lookup list's **items must already be at the target.** In Site scope this is done for you: every library and list is created first, columns that couldn't bind yet are added once all lists exist, custom-list items are copied (lists that others look up come first), and library files are copied last.

**Known limitations** — these are not bugs to report; they are not supported today:

- **Duplicate display text** — for items that were **not** copied in this run, two target items with the same display text (two items titled "Q1") resolve to the one with the lowest ID.
- **Lookups to a different site collection** — a column that looks up a list on another site collection is not recreated (the target list can't be found), and its values are skipped.
- **Projected fields on a list that already exists** — they are added only when the app creates the list (or adds its missing structure); a projected field already on the target is left alone.
- **Lookups to a document library's files** — libraries are copied after custom lists, so a lookup that points at a library's items finds nothing to match on the first run. Run the copy again after the library has been copied.
- **Library scope with a mix of libraries and lists** — only **Site** scope creates all lists and libraries up front. In Library scope, libraries are copied before the lists you selected, so a library column that looks up one of those lists may need a second run after the lists have been copied. Among the selected lists themselves (and the lists added by *Also copy lists that my lists look up*), ones that others look up are copied first. Lists added that way are custom lists only, not libraries.
- **A same-named column that isn't a Lookup** — if the target already has a column with that name but a different type (for example Text), the values are not written and the Custom Fields Details column says so. Rename or remove it and copy again.
- **Items missing at the target** — a value whose display text matches nothing is reported as "unresolved"; the file or item itself still copies.

### Managed metadata — how it copies

A managed-metadata column is bound to a **term set** in the tenant's term store, and a term is identified by a GUID.

- **Same tenant:** the column is recreated with the same term set, and each value is written with its term GUID. Nothing else is needed.
- **Different tenant:** the term store of the target does not have those GUIDs. The app copies the **term set** (group, set, terms with their hierarchy, labels in every language, synonyms, descriptions and custom properties) to the target through Microsoft Graph, builds a map from every source term to its new term, binds the column to the copy, and writes each value with the new term. The copy is repeatable: a group, set or term that is already at the target (matched by name) is reused. The run report gets a *Managed metadata* line: term sets copied (created / already there) and how many values were remapped or left out.
- **The switch:** on the Options screen (different-tenant mode) **Copy managed-metadata term sets to the target tenant** is on by default. Turn it off to skip the term store entirely; the column is then reported as not created. The Column mapping dialog says under each managed-metadata column whether its term set will be copied.
- **What it needs:** the optional delegated Microsoft Graph permission `TermStore.ReadWrite.All` on the registration used for the target (admin consent; the source registration needs `TermStore.Read.All`) and a term store role (for example *Term Store Administrator*) for the signed-in target account. The first term set copy may show a consent prompt. Without them the column is **not created** and the report says so, instead of being skipped silently.
- **Not copied:** term GUIDs themselves (Graph always assigns new ones), tagging availability, deprecation, term order, term store roles and Enterprise Keywords. A term set that belongs to only the source site collection is copied into the target site's own site-collection group.
- A value whose term is not in the copied set (a term outside the column's branch, or one deleted at the source) is left out and counted.
- The term set copy is checked by the live test harness by copying a term set into a second group of the same term store (every GUID changes, so the same remapping is exercised). A copy into another tenant has not been tested live.

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

### The connection error tells me to open Settings

Connection problems that come from the app registration (no Client ID yet, the wrong Tenant ID, a missing redirect address, consent not granted) are explained in plain language under the **Connect** button, with an **Open Settings** button when the fix is there. The technical detail is kept in brackets at the end of the message — include it if you ask for help.

### Sending diagnostics to support

In **Settings**, click **Export diagnostics…** to save a zip with recent logs, basic system information and your settings with organization details removed. It does not include copy reports, file names or site URLs.

### Settings not persisted

Settings are written to `%AppData%\SharePointSmartCopy\settings.json`. On managed corporate machines, Group Policy may restrict writes to the AppData folder. Contact your IT department to verify write access to the `AppData\Roaming` location for your user profile.
