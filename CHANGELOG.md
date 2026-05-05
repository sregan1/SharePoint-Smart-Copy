# Changelog

All notable changes to SharePoint Smart Copy are documented here.

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
