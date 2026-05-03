# SharePoint Smart Copy

A WPF desktop app for copying files between SharePoint Online site collections with support for version history and parallel transfers.

## Prerequisites

- **[.NET 8 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8)** (not just the runtime)
- **Azure AD App Registration** (see below)

## Azure AD App Setup

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Give it a name (e.g. "SharePoint Smart Copy")
3. Under **Authentication** → **Add a platform** → **Mobile and desktop applications**
   - Add redirect URI: `http://localhost`
4. Under **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**:
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
5. Click **Grant admin consent** (requires Global Admin or SharePoint Admin)
6. Copy the **Application (client) ID** from the Overview page

## Build & Run

```bash
cd SharePointSmartCopy
dotnet restore
dotnet run
```

Or open in Visual Studio 2022 and press F5.

## First-Time Configuration

1. Launch the app and click **⚙ Settings**
2. Paste your **Client ID** (from step above)
3. Set **Tenant ID** to your tenant domain (e.g. `contoso.onmicrosoft.com`) or leave as `common`
4. Click **Save**

## How to Use

| Step | What to do |
|------|-----------|
| **1 — Source** | Enter the source site URL and click Connect. A browser window will open for sign-in. |
| **2 — Browse** | Expand document libraries. Check files/folders you want to copy. Checking a folder selects its entire contents. |
| **3 — Target** | Enter the destination site URL, connect, then click the target library or folder in the tree. |
| **4 — Options** | Configure overwrite and version settings, set parallel copy count, review the copy plan. |
| **5 — Copy** | Click "Start Copy". Watch real-time progress per file. Cancel at any time. |
| **6 — Report** | See success/failure summary. Export to CSV. |

## Options

| Option | Description |
|--------|-------------|
| **Overwrite existing files** | If a file exists at the destination, replace it. If unchecked, the copy fails for that file. |
| **Copy all versions** | Downloads and re-uploads each historical version in chronological order. Note: original author and timestamp metadata is not preserved (Graph API limitation). |
| **Parallel copies** | Number of files copied simultaneously. 4–8 is a good balance. Very high values may trigger SharePoint throttling. |

## Notes

- Files under 4 MB use a simple upload; larger files use an upload session in 320 KB chunks.
- Both source and target must be accessible by the signed-in user.
- The app authenticates once and reuses the token for both sites (same tenant assumed).
- Token cache is in-memory per session; re-login is required on restart (silent refresh is attempted first).
