# Changelog

All notable changes to SharePoint Smart Copy are documented here.

---

## 3.1.1 — 2026-06-28

### Fixed

- **Migration API always engages when selected** — Copy Versions off no longer silently fell back to Enhanced REST. Migration API now runs whenever the mode is selected; with Copy Versions off it produces a current-only copy through the same fast batched path.
- **"Missing file info for list item" import errors eliminated** — Standalone `SPListItem` manifest objects were always rejected by SharePoint's importer because it cannot link them to their `SPFile`. `SPFile` alone imports the file with correct metadata. Removing the redundant `SPListItem` emission eliminated the primary source of import failures and freed the full 250-item batch size.
- **Special-character folder names** — Folder GUID resolution now uses `GetFolderByServerRelativePath(decodedurl=...)` instead of `GetFolderByServerRelativeUrl`, fixing 404 failures for folders whose names contain `#`, `%`, `&`, or other URL-reserved characters.
- **Accurate per-file import reporting** — Each file's SPMI result is now attributed by GUID. "0 of N imported" is no longer reported when files actually landed; the error column shows the specific file and reason for any genuine failure.
- **Folder dates and authors preserved** — Source folder created/modified dates and created-by/modified-by users are now written into the SPMI manifest. Previously all migrated folders showed a hardcoded 1999/2000 placeholder date on the target.

### Changed

- **SPMI batch ceiling raised to 250 items** — With per-item SPListItem errors eliminated, batches now operate reliably at SharePoint's package limit (250 items / 250 MB), reducing the number of jobs and total overhead on large runs.
- **Upfront metadata and version cache** — All file metadata and version counts are fetched in a single parallelised pass before import begins, eliminating per-batch Graph API calls and preventing stale-data failures mid-run.
- **Parallel blob uploads within each package** — File blobs are now uploaded concurrently rather than one at a time, reducing package preparation time on multi-file batches.
- **AIMD throttle tuning** — The adaptive parallelism ceiling now starts at the configured soft-start value (not the slider maximum), halves on a throttle event, and probes upward every 45 s only when the semaphore is under active load. The slider acts as a safety cap rather than a target.
- **Large-copy UI responsiveness** — File rows are added to the progress list in batches of 200 via a non-blocking async dispatcher call instead of one synchronous UI round-trip per file. Auto-scroll is coalesced to a single background-priority scroll per burst. Copies with 47,000+ files no longer freeze the progress display.
- **Success filter chip** — A "Success" chip appears alongside "All / Failed / Skipped" on both the copy log (Step 5) and report (Step 6) screens, making it easy to confirm what landed in a large run.

---

## 3.1.0 — 2026-06-13

### Added

- **Person/User column copy** — Person and People columns are now read from the source (with full user detail via `$expand`) and written to the target using the SharePoint `ValidateUpdateListItem` endpoint with claims-format keys (`i:0#.f|membership|email`). Works for both single and multi-value user fields. Applied in all three write paths: Enhanced REST, Migration API manifest, and list item copy.
- **Managed Metadata column copy** — Taxonomy (single) and TaxonomyMulti columns are now read from the source and written using the `Label|TermGuid` format. Because source and target share the same tenant term store, term GUIDs are identical — no term mapping required.
- **Copy-if-newer incremental mode** — A third overwrite option sits alongside Skip and Overwrite: "If newer" compares the source file's last-modified date against the target and copies only when the source is more recent. Files that are already up to date are reported as Skipped ("Up to date") and still receive a permission refresh when Copy Permissions is enabled. Supported by both Migration API and Enhanced REST.
- **HTTP 429 throttling backoff** — SharePoint throttle responses are now handled automatically: the app reads the `Retry-After` header (supporting both seconds and date formats), waits the instructed delay (capped at 120 s) with ±10 % jitter, and retries up to 5 times. A "Throttled, retrying in N s…" status message appears during the wait. Retry count increased from 3 to 5.
- **Column mapping: auto-match on open** — When the Column Mappings dialog opens for the first time (no saved state), source columns are automatically pre-matched to target columns that share the same internal or display name and a compatible field type. Previously the user had to click "Auto-match" manually.
- **Column mapping: type-filtered dropdowns** — Each source column's dropdown now only lists target columns whose type is compatible with the source. A Person source column no longer shows Text targets; a Date column no longer shows Lookup targets; etc. Incompatible pairings are removed from the list entirely.
- **Light/Dark/System theme** — A theme selector in Settings switches between Light, Dark, and System-follows-Windows palettes at runtime without restarting.
- **Copy log filter chips** — "All / Failed / Skipped" radio chips on the Copy progress and Report screens filter the file list in place, making it easy to focus on failures in a large run.
- **Custom title bar** — The blue app bar now doubles as the window chrome (drag-to-move, double-click to maximize). Standard caption buttons (minimize / maximize / close) are embedded in the bar.

### Changed

- **Overwrite control** changed from a checkbox to a three-way horizontal radio group: **Skip existing** / **Overwrite** / **If newer**. The previous `OverwriteFiles = true` setting is automatically migrated to `Overwrite` on first launch.
- **Permissions UI** refactored: the separate Permissions tab in the report is removed. Permission result (status + details) now appears as two inline columns on each file row — Perm Status and Perm Details. The columns are hidden when Copy Permissions is off.
- **Copy Options layout** compacted so all options fit a standard 720 px window without a scrollbar: option captions moved into ⓘ tooltips, radio groups converted from stacked to single horizontal rows, Preserve Metadata moved to Advanced.
- **Destination Name in Copy Preview** is now read-only when individual list items are selected (the name was already authoritatively chosen on the Target step). A tooltip explains this when the field is locked.
- **Navigation bar** given a distinct background (`#E9E9E9` in light mode) so it is always visually anchored and the Back/Next buttons never appear to float against white content surfaces.
- **Release packaging**: GitHub release now uses `PublishSingleFile=true` — the distributed `.exe` is a true self-contained binary and no longer an apphost wrapper that requires `SharePointSmartCopy.dll` alongside it.

### Fixed

- Fixed: phantom document library appearing in Copy Preview when only a custom list was selected. Root cause: WPF's tri-state checkbox left non-list nodes in `IsChecked = null`; the app misread `null` as "items-only mode" and included the library. Non-list nodes are now always two-state.
- Fixed: New List name not populating in Copy Preview. `OverrideName` was never set when advancing from the Target step; it is now assigned from `EffectiveDestinationListName` during step navigation.
- Fixed: custom column writes corrupting preserved metadata timestamps. `ValidateUpdateListItem` updates the `Modified` and `Editor` fields as a side effect. Fix: custom columns are now applied before the metadata preservation stamp, so the stamp always runs last.
- Fixed: column mapping auto-match silently pairing incompatible types (e.g. a Person column matching a same-named Text column). `AreTypesCompatible()` now gates all name-based matches.

---

## 3.0.0

### New Features

**Libraries & Lists copy scope**
- Copy entire document libraries to a target site — recreates the library with matching versioning settings, custom column schema, and all content.
- Copy generic lists (Tasks, Calendars, Announcements, custom lists) — recreates structure and copies list item data including all custom column values.
- Overwrite behavior: if a library or list already exists at the target, it is skipped gracefully (shown as ⏭ Skipped in the report) rather than failing.
- System libraries (Site Assets, Style Library) that are not returned by the Graph Drives API are resolved via a SharePoint REST fallback and handled correctly.

**Individual list item selection**
- Custom lists in the Libraries & Lists scope are now expandable in the browse tree — expand a list to see its items, each with a checkbox.
- Check individual items to copy only a subset; leave the list node itself checked (without expanding) to copy the whole list as before.
- Items are loaded on demand via a lightweight REST query (`$select=Id,Title`) and sorted by ID. Lists with more than 5,000 items cannot be expanded for item-level selection (the whole-list copy still works for those).

**Destination list picker for item-level copies**
- When individual list items are selected, the Target step shows a "Destination List" dropdown populated with the custom lists on the target site.
- The Next button is held disabled until a destination list is chosen, preventing accidental misconfiguration.

**Column mapping for list item copies**
- The Configure Mappings button is now accessible in the Options step when copying individual list items, with source and target columns loaded from the respective lists.
- Column mappings are applied when writing item field values to the target list, so items land in the correct columns even when source and target schemas differ.

**Site copy scope**
- Copy all document libraries and custom lists from a source site to a target site in a single operation.
- Navigation links (Quick Launch) are optionally copied alongside the content.
- Each library and list appears as its own row in the progress screen and report.

**Pages copy scope**
- Copy modern SharePoint pages (.aspx) between sites.
- Optional web part URL remapping: internal URLs within web part properties are rewritten from the source site domain to the target site domain.

**Custom column mapping**
- When copying files between libraries with different column schemas, a Configure Mappings dialog lets you map source columns to target columns or mark them for creation on the target.
- The mapping dialog is accessible from the Copy Options step when "Copy custom column values" is enabled.

### UI Improvements

- New mode selection tile on the Browse step: choose Files, Libraries & Lists, Site, or Pages before browsing.
- Copy Preview panel on the Options step is now scope-aware: shows a library/list summary for Libraries & Lists and Site scopes, and the file list for Files and Pages scopes.
- Copy Preview summary now distinguishes between whole libraries/lists selected and individual list items selected.
- Version history sub-options (Copy all / Latest N) are shown inline under the "Copy version history" checkbox with ⓘ tooltips.
- Migration API is pre-selected by default on the Options step.
- Versions column in the progress and report grids is hidden for library/list creation rows (where it is not applicable).
- The Configure Mappings button is disabled while column metadata is loading, preventing the dialog from opening with an empty column list.
- Removed splash screen on startup.

### Bug Fixes

- Fixed: Events and other custom lists showed ❌ Failed instead of ⏭ Skipped when the list already existed at the target during a site copy.
- Fixed: Migration API radio button was not pre-selected when navigating to the Options step for the Libraries & Lists scope due to a `GroupName` conflict between the two copy-mode radio button groups.
- Fixed: `CopyLibraryContent` and `RemapPageWebPartUrls` settings were not restored from disk on startup.
- Fixed: column mappings were not cleared when switching copy scope or navigating back from the Options step, causing stale mappings from a previous run to be applied.
- Fixed: opening the column mapping dialog for a custom list source caused a "malformed drive ID" Graph error. Custom lists have no drive; the fix uses the list GUID directly to load column definitions, bypassing the drive-based lookup entirely.
- Fixed: individual list item copy reported success even when no items were copied because per-item errors were silently swallowed. Item-level errors are now caught individually and surfaced in the copy report.

---

## 2.1.0

### Bug Fixes

**Copy correctness**
- Fixed: overwriting a file using the Migration API appended imported versions to the existing version history instead of replacing it. The root cause is that SPMI `UPDATE` (when a file GUID already exists) extends history rather than replacing it. The fix pre-deletes the target file before submitting the SPMI job so the import always performs a fresh `INSERT`.
- Fixed: "zombie" files — AllDocs rows left behind when a previous import failed partway through, where Graph returns 404 but the SharePoint content database still holds a record — caused subsequent SPMI imports at the same URL to fail. Zombie files are now detected via `/_api/web/GetFileByServerRelativeUrl` and permanently deleted before import.
- Fixed: Graph `DeleteAsync` only soft-deletes (moves to recycle bin). Soft-deleted records still interfere with SPMI imports at the same URL. Files are now fully purged by recycling and then deleting the resulting recycle bin entry.
- Fixed: selecting a destination folder more than one level deep placed files at the library root instead of inside the selected folder. The relative path was computed from `ServerRelativePath`, which is only populated on library-root nodes. Path computation now walks the node's parent chain, which is always populated. `TargetParentItemId` is anchored to the library root and `TargetSubFolderPath` carries the full path from library root, keeping REST and Migration API behaviour consistent.

**Authentication**
- Fixed: SharePoint REST requests could fail with 401 after a token expiry mid-session. All SharePoint REST calls now share a helper that automatically retries once with a force-refreshed token on a 401 response.

### Performance

- Folder metadata (Created By, Modified By, dates) is now applied in the background after the file copy completes rather than blocking the completion signal. For large copies with many subfolders this eliminates a multi-minute post-copy wait.
- Subfolder metadata within each folder job is applied in parallel (up to the configured parallel copies limit) instead of sequentially.

### UI Improvements

- The progress screen now shows a "Updating folder metadata…" spinner after files finish copying, and "✔ Folder metadata updated" once complete.
- The Next → button on the copy progress screen is held disabled until folder metadata finishes applying, so the report is not shown before the operation is truly complete.
- "✅ Copy complete! Click Next to see the full report." now appears only after metadata is done, not as soon as files finish.
- Copy Preview: the From / To site URL banners now wrap instead of truncating with ellipsis.
- Copy Preview: hovering over a truncated source or destination path in the preview list shows the full path in a tooltip.
- Settings dialog: increased height and fixed the layout so the Azure App Registration setup instructions are always fully visible.

### Other

- Copy settings (Overwrite, Copy Versions, Max Versions, Parallel Copies) are now persisted between sessions.
- `TargetLibraryServerRelativeUrl` is now propagated to child file jobs created during folder enumeration, avoiding a redundant API call to look up the library URL per-batch.

---

## 2.0.0

### Overview

Full rewrite of the application. The core copy engine has been rebuilt around SharePoint's Migration API for high-fidelity bulk migrations, with the Enhanced REST path retained for small batches and quick copies.

### New Features

- **Migration API mode** — imports files server-side via SharePoint's Migration Import API (SPMI). Version numbers, dates, and per-version editors are preserved exactly as on the source. Requires `AllSites.FullControl` permission and Site Collection Administrator membership on the target site.
- **Parallel Migration jobs** — the "Parallel copies" slider (1–16) now controls the number of concurrent SPMI jobs (capped at 5) when using Migration API mode, with blob uploads within each job also running in parallel.
- **Full site URL in paths** — source and target paths displayed in the copy screen, report screen, and CSV export now include the full site collection URL for clarity.
- **Run history** — completed copy runs are saved locally and viewable in the History dialog, with per-file status and CSV export.
- **← Back button on all steps** — the Back button is now available on every step including the final Report screen, so you can review options or re-run without restarting the wizard.
- **ⓘ info icons on copy modes** — the Migration API and Enhanced REST options on the Options screen now show an inline ⓘ icon; hover to see a description of each mode without cluttering the UI.

### Enhanced REST improvements (carried forward from v1)

- Modified By and Modified date correct per version in version history
- Author and Created date preserved on the file and folders
- Overwrite and skip-existing options
- Configurable max-versions limit per run

---

## 1.1.2

### Bug Fixes

- Fixed: application crashed on launch with `System.DllNotFoundException` when published as a single-file executable. WPF's native runtime DLLs do not extract reliably from a single-file bundle. The release now ships as a zip containing all files alongside the executable, eliminating the self-extraction step entirely.

### Documentation

- Added Installation section to both READMEs explaining how to download, extract, and run the zip release.
- Removed .NET 8 Desktop Runtime from Requirements — the runtime is now bundled in the release zip.
- Added SmartScreen note to Installation section.

---

## 1.1 fixes

### Bug Fixes

**Copy behaviour**
- Fixed: files were still copied when "Overwrite existing files" was unchecked. The small-file upload path (`PutAsync`) ignores conflict behaviour entirely; the fix adds an explicit existence check before any upload begins, so existing files are correctly skipped and marked `Skipped` in the results grid.
- Fixed: sequential version strategy showed today's date in the SharePoint library "Modified" column instead of the source file's date. The root cause is that SharePoint's library view reads `fileSystemInfo.lastModifiedDateTime`, not the `listItem.Modified` field set by VULI. A `PatchFileSystemDateAsync` call is now made after VULI for both the sequential strategy and the single-version copy path.
- Fixed: copied files showed the SharePoint "new item" glimmer badge. The badge is driven by `fileSystemInfo.createdDateTime`; uploading a file sets it to today. The fix extends `PatchFileSystemDateAsync` to also write `createdDateTime` so the source's original creation date is preserved.
- Fixed: folder metadata (Created By, Created Date, Modified By, Modified Date) did not match the source after a folder copy. VULI was already being called but `fileSystemInfo` dates were not being patched. `PatchFileSystemDateAsync` is now called for every folder — root and sub-folders — after VULI. Folders have no version history so patching `fileSystemInfo` has no phantom-version side-effect.

**Authentication / Connect flow**
- Fixed: the Source Connect button was permanently disabled after clicking Disconnect. The button's `IsEnabled` binding used `InvBoolToVis` (returns `Visibility`) instead of `InvBool` (returns `bool`). WPF silently ignores the type mismatch on first render, then converts `Visibility.Visible` (integer 0) to `false` via `Convert.ToBoolean`, permanently disabling the button.
- Fixed: the Target Connect button had the identical `InvBoolToVis` / `IsEnabled` mismatch and was corrected the same way.
- Fixed: after clicking the red Disconnect button, the Connect button reappeared but remained unclickable until MSAL finished its auth flow. `AsyncRelayCommand` keeps `CanExecute = false` while its task is running, and `GetAccessTokenAsync` / `GetSharePointTokenAsync` were not receiving a `CancellationToken`, so MSAL kept blocking even after cancellation was requested. Both auth methods now accept and forward a `CancellationToken` to every MSAL `ExecuteAsync` call, so cancellation is immediate.

### UI Improvements

**Splash screen**
- Changed from maximised fullscreen to a fixed 960 × 720 window matching the main application window size.
- Replaced the centred logo image with a full-bleed splash image (`splash.png`, `UniformToFill`).

**Step 2 — Connect to Target**
- Fixed a layout jump where selecting a target folder caused the TreeView to shift upward. The selected-folder banner now uses `Visibility.Hidden` (reserves space) instead of `Visibility.Collapsed` (removes space), so the layout is stable whether or not a folder is selected.
- Aligned the Target Site URL text box with the top of the folder tree by adding a compensating top margin to the field label.

**Step 3 — Copy Options**
- Reworked the panel layout: the options column is now wrapped in a `ScrollViewer` (prevents expanded version options from overflowing the window), and the Copy Preview list is placed in a `Grid` with a star row so it fills all remaining vertical space and scrolls independently.
- Replaced the three verbose description text blocks under the version options with inline `ⓘ` glyphs. Hovering over the glyph shows the full description in a tooltip, keeping the panel compact. Tooltip appears instantly with a 30-second display duration.
- Widened the options column from 260 px to 320 px so all option labels fit on one line without wrapping.
- Tightened vertical margins and border padding throughout the options box (~38 px saved) so the full panel fits on screen without requiring the scroll bar.
- Extended the Copy Options white box bottom padding so text at the bottom of the box was not clipped.

### Documentation

- Updated user guide (SharePointSmartCopy_UserGuide.docx) to v1.1: added Version strategy section to the Step 4 Copy Options chapter, explaining the Preserve metadata and Keep sequential radio options and their tradeoffs.
- Regenerated all 10 user guide screenshots to match the current UI.
- Updated both READMEs (root and Docs) to document the Version strategy options and corrected the parallel copies range from 1–8 to 1–16.
- Added a `http://localhost` redirect URI step to the root README setup instructions.
- Removed stale debug screenshot (`check_state.png`).
