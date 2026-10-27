# Changelog

All notable changes to SharePoint Smart Copy are documented here.

---

## 4.1.1 — 2026-10-26

### Added

- **List item version history** — copying a list used to give every item one version, although the Copy version history option and the README implied history copied for everything (only file versions in document libraries did). A new item is now created from its oldest version and each later version is replayed with its own values, Modified date and editor, so the copy shows the same version labels (proven on a live tenant: 4 source versions gave 1.0–4.0 with matching values and dates). A **List item version history** option (Library and Site scopes) chooses *Same as files* (the default: follows Copy versions and All / Latest N), *All versions* or *Current only*. Items that already exist on the target keep their own history. Text, number, choice, yes/no and date columns follow the history; people, lookup, taxonomy and link columns copy their current value only. Each replayed version is one extra write, so a large list takes longer with history on. Replaying needs `ValidateUpdateListItem` with `bNewDocumentUpdate=false` (the value `true` creates no version) and an item created with its first version's values, so no empty extra first version appears.

### Changed

- **Lists show in Site scope on the Options screen** — only Library scope loaded custom lists, so a Site copy never listed them and the list item version option could not apply. Site scope now loads them too (as top-level entries without expandable items).
- **List item version option appears only when it applies** — it is shown when list items will be copied (a custom list or some of its items is selected in Library scope, or, in Site scope, the site has lists and Copy library content is on), not for every Library or Site copy.
- **Target step no longer jumps while typing a URL** — the step's grid sized itself to its content, so a longer target URL widened the columns and shifted the whole layout. The grid now stretches to a fixed maximum width and the connected-site banner wraps.

### Fixed

- **Activity logs were empty** — every `activity-*.log` stopped at its two header lines. The writer read `ChannelReader.Count`, which throws on a single-reader channel, and the silent catch ended the log on the first line, so a failed run left nothing to diagnose (a 14-minute run with 5,897 "Invalid request" failures had no log). The writer now drains the queue with `TryRead` and flushes each burst.

---

## 4.1.0 — 2026-10-26

### Added

- **Forms tool (Browse step → Other tools → Forms)** — four guided steps (Forms, Target, Copy, Results) like the file copy: the target has its own step with an explicit sign-in, and a results dashboard shows copied / with warnings / failed counts, questions copied, duration, and a per-form table with reasons, plus Open selected copy and Export results (Excel). copies Microsoft Forms to the target account with every question type, choice branching, quiz type, settings, and theme, keeping the original question IDs. Exports a form's responses to Excel (responses can't be re-submitted). It uses the unofficial Microsoft Forms web API, which Microsoft does not document or support, so the dialog says so. Ranking, Likert and the form's logo, background and header images are copied too. Each copy is recorded in a local form ID map (`%APPDATA%\SharePointSmartCopy\forms-id-map.json`) for the Flows tool.
- **Flows tool (Browse step → Other tools → Flows)** — six guided steps (Packages, Sites, Lists & columns, Forms & email, Write, Report) with an automatic scan, automatic list and column matching on entry, and a report dashboard with a what-to-do-next checklist. Rewrites exported Power Automate flow packages (.zip) so they point at the target: it reads each source list's name through your connection, finds the same-named list on the target, matches columns by internal then display name, and also rewrites sites, copied form IDs, and a whole email domain. Writes patched packages and a report (including anything still pointing at the source) to a separate folder; the originals are never modified. Connections can't be moved between tenants and must be recreated and re-selected after import. After scanning, a table shows which services the flows use (SharePoint, Forms, Outlook, Teams, Planner, Excel, OneDrive, Approvals...) and which of them the tool rewrites or only reports. The column mapping is an editable grid, pre-filled by automatic matching (same name, same display name, then a similar name), with a dropdown of the target list's columns per row; the same column can map differently on different lists. Target lists and forms can also be chosen by hand: paste a list GUID or list-settings link, or a form ID or form edit-page link.
- **Forms tool refinements** — a filter box and an **Already copied** column on the forms list (from the local form ID map, which now also records the target account and date), a confirmation before copying a form to an account it was already copied to, Cancel for loading, sign-in and export (not only the copy), a confirmation when closing during a copy, and response export with real date cells, a score column when the responses carry one, and a **One workbook** option (one sheet per form). Forms API calls retry on throttling and on an expired token.
- **Flows tool refinements** — a **Load mapping…** button restores a saved mapping, an **Other values** table maps individual email addresses, the tenant ID and library paths, each form ID is checked against the source tenant, solution exports (a `Workflows` folder) are read as well as legacy packages, and the report gains **What these flows use** and **To do after import** sections plus an HTML copy (`rewrite-report.html`). List lookups run four at a time. The Sites step has a **Connections** card with explicit source and target sign-in, every step explains what it does and why, an activity list shows progress, and **Cancel** stops a pending browser sign-in in the Flows and Forms tools so it can be retried in another browser. The Forms & email step fills in the target tenant ID, the target account's address and the email domains from your sign-ins (empty cells only, with a note), lists the target account's forms to pick from (or paste an ID), the Lists & columns step shows one card per list (used by which flows, its name and every column on the source, a target-list dropdown defaulting to the same-named list, and a column mapping for the target list), and the Write step lets you name the updated package(s) with a `{name}` pattern and a live preview.
- **Flows tool redesign** — all six steps follow one layout: SOURCE (grey) and TARGET (blue) panes with tenant addresses, status chips (Ready, Needs a look, Can't read, Info), one warning at a time, a one-line "what happens" with an expandable "why", and a collapsible one-line activity strip. Lists & columns shows the source list's name (not its ID), its columns, a target-list dropdown and a column mapping side by side; a list that cannot be read still offers a mapping from the flow's own column names, and a list whose ID is already on the target says so. Packages is a list with per-package flows and drag-and-drop; Sites is simpler (the advanced "read from" column is hidden unless needed); Forms & email groups values and marks what was filled in from sign-ins; Write opens with a "Ready to write?" checklist with links to the step that fixes each item; Report lists what to review. Fixed: a card stayed "needs a look" after its columns were fixed, "✓" shown when columns could not be read, a stale "Reading…" footer, and unreadable text in the dark theme.
- **Flows tool checks before you can go on** — the Sites, Lists & columns and Forms & email steps keep **Next** off, and say why in the footer, while something would make the imported flows fail: the target is not signed in (Multi-tenant) or a site has no target, a list has no target list, a column the flows use is not on the target list, a column the target list requires is not filled in by a "Create item" step (import error "missing required property"), or a form has no valid target ID. When a required column is left out by a Create item step, the list shows a box to type the value for it, and the tool adds that value to those steps in the updated package (saved in the mapping file as `RequiredValues`). A list that would fail shows the reasons and a **Check again** button, which re-reads the target list after you fix it in SharePoint and keeps your column choices.
- **Flows tool follow-ups** — **Write updated packages** now opens a Save As dialog for each package (name pre-filled from the Save as card), so you choose where and under what name every updated copy is saved; an original can never be overwritten, and Cancel writes nothing. each package now has its own **Save as** name on the Write step (with **Keep the original names** and **Add ' - migrated'**), form names load automatically when you reach Forms & email (the button remains as a retry), the Report step has a **Back** button, and "What these flows use" is open by default. When every site in a loaded package is already on the target tenant (a package this tool wrote), one warning explains it and offers **Remove these packages**; the tool also remembers the packages it wrote. Other values no longer offer target emails or the target tenant ID as the source value.
- **Browse step: "Other tools" row** — the four copy-scope tiles are back to one row. Below them, a lighter **Other tools** row holds **Compare** (moved out of the app bar, so the header now has just History and Settings), **Forms** and **Flows**. Compare is still available on the Report step.
- **Permissions by feature** — the README and User Guide now include a matrix showing which API permissions each feature needs (for example, Flows needs neither the Microsoft Forms permission nor `AllSites.FullControl`).
- **Microsoft Forms permission** — the Forms tool needs the delegated Microsoft Forms permission on the app registration (see the User Guide). Existing features are unaffected without it.

### Changed

- **Translations complete** — all 741 strings that were English-only (the whole Forms and Flows tools, sign-in and consent errors, dialog text, the new Other tools row) are now translated into all 29 other languages. `Tools/CheckResxKeys.ps1` lists any key missing from a language and any `{0}`-style placeholder that differs from English, so a gap can no longer slip in unnoticed.
- **CSV exports open correctly in Excel** — the copy report and the saved-run report are written as UTF-8 with a byte-order mark, so accented, Cyrillic, CJK, Arabic and Hebrew file names no longer turn into garbage. One shared writer now serves both.
- **Dialog buttons are consistent** — Forms and Flows use the same **← Back** and **Close** labels as the rest of the app, and Esc closes them (it asks first if work is still running). The column mapping dialog also closes on Esc.
- **Multi-line text boxes start at the top** — the copy log in the Forms dialog (and the activity log in Flows) centered a short text vertically, so it floated in the middle of an empty box. Every text box that holds several lines or wraps now aligns to the top; one-line inputs stay centered.
- **Report summary tiles reachable by keyboard** — the Failed and Skipped tiles can be focused with Tab and activated with Enter or Space; Space on a tree checkbox now follows the same unchecked / all / items-only cycle as the mouse.
- **Toast notifications stand out in the dark theme**, and "Up to date" / "Already exists" in the Details column are localized.
- **Diagnostics export** no longer includes activity logs: those hold file names, site URLs and SharePoint's own error reports, which the export promised to leave out.

### Fixed

- **Migration API: a failure after overwrite deletes now says what happened** — the manifest upload shares the data upload's retry policy, and if preparing a batch still fails after existing target files were removed, those files report that they are missing from the target and need a re-run, instead of a bare network error.
- **Migration API: one failing library no longer fails the rest** — a library-level error is contained to that library's files.
- **Migration API: other hardening** — a job that reports no progress for two hours is given up on instead of being polled forever; a resumed download only continues if the source file is unchanged (ETag / If-Match), and starts over if it changed; throttle waits, connection resets and permanent errors (403/404/410) have separate retry budgets; a non-GET request is no longer repeated after a bare 503; download responses are released on errors; overwrite deletes run four at a time; cancel now reaches the folder lookup, delete and job-submit calls.
- **Crash recovery** no longer records not-yet-started files as finished (they stayed "Pending" after a recovered run).
- **List item permissions** are shown on the item's own row (every item used to overwrite the list's row), and a failure to read whether the source has its own permissions is reported on the item instead of silently skipping the permission copy.
- **Flows: stale matches after changing a site target** — going back and changing a site's target now re-matches the lists on the new site.
- **Flows: adding or removing a package keeps what you entered** (site targets, "read from" edits, form IDs, replacement values) instead of resetting every table; **Start over** also clears the email-domain mapping.
- **Flows: rewrite accuracy** — renaming columns can no longer overwrite a value when one rename's new name is another's old name (and a clash is skipped and reported); bare column names in `$filter` are renamed only outside quoted values; `contoso.com` no longer matches inside `contoso.com.au`; a tenant-root site mapping no longer rewrites sub-sites; list IDs written as `%7B…%7D` are rewritten; an unwritable output folder or a package with an unexpected shape fails that package instead of the whole run.
- **Forms load automatically** — the Forms dialog reads your forms as soon as it opens (you may be asked to sign in) instead of waiting for a click on **Load my forms**. The button remains as **Reload forms** and as the retry when the first attempt fails or is cancelled.
- **Copy text from anywhere** — right-click any text in the app (a grid cell, a status line, a hint, a title) and choose **Copy**. Grid cells copy the whole cell, so a long Details message can be pasted into an email. Text boxes, buttons and combo boxes keep their own behavior.
- **Forms: logo and images** — a form's logo, background and header cannot be set by patching the form (the service accepts the request and changes nothing; the files belong to the source form). The copy now uploads each image to the new form the way the Forms web app does (the file itself, sent to the form's image endpoint), checks that it really is on the form, and points the copied theme at the new background instead of the source form's. If an image still cannot be added, a copy of the file is saved under `Documents\SharePointSmartCopy\Forms images` and the warning names the file so it can be added by hand. The result of every attempt is in `formscopy-<date>.log`.
- **Forms: image warnings are now accurate** — the copy no longer says an image "was not copied" just because the source had one (the background was being carried over all along). After the copy the target form is read back; an image that is missing is patched once more on its own and checked again, and only an image that still did not arrive is reported, with a new wording that asks you to set it again in Microsoft Forms. The source and target image values are written to the `formscopy-<date>.log` file so a logo or header that does not copy can be diagnosed.
- **Forms: Ranking and Likert questions** — the repair that re-applies their options now uses the question ID the target actually assigned, compares the options as JSON instead of text, reads the question back after the repair and warns if it still differs from the source (instead of a silently wrong question). The full source, request and response for these questions is written to `%LOCALAPPDATA%\SharePointSmartCopy\Logs\formscopy-<date>.log` (left out of the diagnostics export) so a copy that still comes out wrong can be diagnosed from the real payloads. The options turned out not to be part of the question at all: Ranking options and the Likert answer scale are a separate "choices" collection of the question, which creating the question never copied. They are now read from the source and written to the copy (the default choice is updated in place, the rest are added, extra defaults removed).
- **Forms: branching and cancel** — branching rules are applied after the questions exist; cancelling mid-copy lists the partial form (with its ID) in the results; a gateway error no longer retries the create or question POST (which could duplicate a form); a damaged ID-map file is kept aside as `.corrupt-<time>` instead of being overwritten; response export clips cells over Excel's size limit instead of failing.

---

## 4.0.2 — 2026-10-22

### Changed

- **Each list now gets its own destination on the Target step** — when several lists had items selected, a single "Destination List" picker applied to all of them (every list's items landed in the one list you chose), with nothing saying which list it was for. The Target step now shows one card per source list, naming the list and letting you pick an existing target list or create a new one for it. Each defaults to the same-named target list if there is one, otherwise to creating a new list with the source name. The Options-step mapping is still shared and is based on the first existing destination.

### Fixed

- **Unchecking items under an "items only" list is now honored** — moving a list's checkbox to the middle (items-only) state ticks every loaded item, but the copy then ignored the item checkboxes entirely, so a user who unchecked all but one item still got every item copied and reported. When some items are checked and some aren't, only the checked ones are copied; with all checked, or the list never expanded, it still copies everything.
- **Lists and libraries with Lookup columns no longer fail to read (HTTP 400)** — the source read asked SharePoint for `Column/LookupId` and `Column/LookupValue`, which aren't fields on the lookup's target list, so any list with a lookup column was created at the target but copied zero items (the same query broke custom-column reads for libraries). Lookups are now read as `Column/Id` plus the column's show field, so lookups that show a column other than Title work. If SharePoint still rejects the show field, the read retries with IDs only instead of failing the whole list.

---

## 4.0.1 — 2026-10-21

### Fixed

- **Lookup values on copied list items now resolve** — custom-list item copy sent the lookup's display text, which `ValidateUpdateListItem` doesn't accept (it needs the target item's ID), so lookup values on list items failed or came out blank. They are now resolved to target item IDs the same way file columns are, and unresolved or partly-resolved values are reported on the item. The same write also stopped reporting read-only system fields as field errors.
- **"Create in target" for a missing Lookup column now works in Files scope** — the source site wasn't passed along, so the matching target list couldn't be found and the column was never created.
- **Site scope copies lists before files, and creates every list and library first** — a library's Lookup column that pointed at a custom list (or a list that pointed at a later list) couldn't be created because that list didn't exist yet, and its values had nothing to resolve against. Site scope now creates all libraries and lists, re-adds any column that couldn't bind the first time, copies custom-list items in lookup-dependency order, and only then copies library files. Lookups that point at a *document library's* files still depend on copy order, since libraries are copied last.
- **A lookup value is no longer remembered as "not found" across runs** — the lookup cache lived for the whole app session and also kept "no such value" forever, so after copying a lookup list's items and re-running, files kept failing until the app was restarted. The cache is now cleared at the start of every run, and a miss is remembered for only 90 seconds.
- **Lookups with duplicate display text resolve consistently** — the query had no ordering, so which of two same-titled items won was arbitrary; it is now always the lowest ID.
- **A clearer message when the target column isn't a Lookup** — a same-named Text (or other) column already at the target used to report only "lookup unresolved"; it now names the target column's type and says to rename or remove it. The message is translated in all 30 languages.
- **Library scope copies lookup lists first** — selected custom lists are now copied in lookup-dependency order, as Site scope already does.

### Known limitations (lookups)

Documented in the User Guide under *Troubleshooting → Lookup columns*; deliberately not addressed:

- Lookups to a list on a **different site** are not recreated and their values are skipped.
- Lookups to a **document library's** files depend on copy order (libraries are copied last).
- **Duplicate display text** can't be told apart — matching is by text, not by original item identity (that would need a source-to-target item-ID map).
- In **Library scope**, libraries are still copied before the selected lists, so a library column that looks up one of those lists may need a second run; only Site scope creates everything up front.

---

## 4.0.0 — 2026-10-13

### Added

- **Copy between two Microsoft 365 tenants (Multi-tenant mode, beta)** — a **Connection: Single tenant / Multi-tenant** switch on the first screen. In Multi-tenant mode the source and the target each sign in with their own app registration and account (chosen right under the switch; the title bar shows "Source: … · Target: …"), and every Graph and SharePoint request is routed to the right tenant's token by site host or drive ID. Single-tenant mode is unchanged: the router has nothing registered and always returns the one sign-in. A request for a host or drive that was never connected now fails with a clear "reconnect" message instead of quietly using the other tenant's token. Tested against two tenant pairs; other setups (guest accounts, conditional access, very large libraries) may differ, so try a small library first and report anything unexpected.
- **Identity mapping with a user table** — people and groups are different accounts in each tenant, so the **Identity mapping…** dialog is now a table of *source user → target user* (groups by name) on top of an optional email-domain swap. It supports Add/Remove, CSV import and export (quoted fields, header row), and **Fill targets from domain swap**. Rows with a problem (blank cell, duplicate source, source equal to target) are highlighted and block saving. **Check in target** looks each target email up in the target tenant and shows ✓ Found / ✗ Not found. The mapping is applied wherever an identity is written: migration-package users, file and folder author/editor, person columns and permission principals. Earlier text-box pairs convert automatically.
- **Unresolved users are reported, and can be mapped** — a cross-tenant run records every person the target tenant could not find (needed as an author, editor, permission or person column), including a pre-check of the authors and editors named in a Migration API package, which was never validated before. The report shows a banner with **Map these users…**, which opens the table pre-filled with them; **History → Map unresolved users…** does the same for older runs. Saving a mapping affects the next run — Copy if newer skips files that are already up to date, so authors already stamped are only re-stamped by an Overwrite run.
- **Crash recovery and Resume** — every run now journals its finished rows as it goes, so a crash, kill or power loss no longer leaves no History entry and nothing to verify; the next launch recovers the run as an **INTERRUPTED** entry. **History → Resume / run again** reconnects and re-runs a saved file selection with Copy if newer, so only what is missing or changed transfers.
- **Large Migration API versions stream through encrypted temp files** — a version of 500 MB or more is downloaded to a temp file, encrypted by streaming (MD5 computed over the ciphertext as it is written, the same bytes the in-memory path hashed) and uploaded from the file. The 2 GB-per-version ceiling is gone, and a multi-GB version no longer costs its size twice in memory. Smaller versions take the original path unchanged. The temp files delete themselves, including if the app dies.
- **Hyperlink and Currency columns are copied** — both were silently skipped before. (Image and Location columns still are.)
- **Diagnostics** — **Settings → Export diagnostics…** saves a zip of recent logs, system info and settings with organization details redacted (no reports, file names or site URLs). Unhandled background-thread and unobserved-task exceptions now write a `crash-*.log` instead of vanishing, logs carry full timestamps and the build/culture, and logs older than 60 days are pruned.
- **Friendlier everyday UI** — a first-run banner that points to Settings; connection errors in plain language with an **Open Settings** button (missing consent, wrong tenant ID, missing redirect, 403/404…) instead of a clipped raw message; the footer says *why* Next is disabled; a confirmation before closing or cancelling a running copy and before switching connection mode; taskbar progress (red after a run with failures) and a taskbar flash when a run finishes in the background; a **View report** button in the completion line; a resizable activity feed; Failed/Skipped summary tiles that filter the list; **Compare source & target** and **Open History** on the report; column mapping surfaced outside the collapsed Advanced options; Enter connects from the URL boxes and Esc closes dialogs; the ⓘ help icons are reachable by keyboard.

### Changed

- **Saving Settings no longer signs you out** unless the app registrations actually changed — changing only the theme or language used to reconfigure sign-in (and, in Multi-tenant mode, disconnect both sites).
- **Folder metadata is written in the order the file path already used** — the dates PATCH first, author/editor last — because the Graph `fileSystemInfo` PATCH re-stamps the item as the migrating account; the old order left every Enhanced REST folder's Modified By set to the copying account. Failures in either write are now reported in the activity log instead of discarded.
- **Copy-log sorting works in every language and covers the custom-field columns** — it matched on the English header text, so it silently did nothing in the other 29 languages.
- **Layouts wrap instead of clipping** — the Options column, label-plus-help rows, report toolbar and Settings footer wrap, and Settings, Identity mapping and History are resizable, so long translations (German, Finnish…) fit. Progress sentences are single translatable strings built from each language's existing wording.
- **Dark-theme text contrast** — links, glyphs and the setup text in Settings use a lighter accent that clears 4.5:1; the standard accent is only about 3:1 on dark surfaces.
- **The title bar no longer runs the account text under the app name** — the signed-in account(s) sat in a fixed-width box on top of the logo text, which the longer Multi-tenant text ("Source: … · Target: …") made visible on a narrow window. The text now takes only the room the buttons leave and trims with an ellipsis (the full text is the tooltip).
- **"Azure AD" is now "Microsoft Entra ID"** in the interface and the User Guide.
- **Internal** — `CopyService.ExecuteAsync` takes a named `CopyRunOptions` record instead of 26 positional parameters (the 3.4.0 "missing argument at one of four call sites" bug class), and `SharePointService` and `MainViewModel` are split into partial-class files by area (a pure move, verified line for line).

### Fixed

- **Permissions were silently lost on large libraries** — a failed page of the permission-flag read now throws (as intended), but all four callers swallowed it with an empty `catch`, so one throttled page meant every file copied with inherited permissions, reported Success, and logged nothing. The failure is now a visible warning.
- **Migration API Overwrite / If-newer deleted the target before the source was downloaded** — the delete is permanent (recycle bin purged too), so any later download failure, size limit, packaging error or cancel left the file gone from both sides. The delete now happens only after every blob for that file is uploaded, and a file that already failed keeps its existing target.
- **Unattributed import errors could be reported as Success** — the check compared SharePoint's error count to *every* failed row, including rows that failed before submission, so five download failures hid three real import failures and those three were promoted to Success.
- **If-newer's conflict retry purged targets without comparing dates** — after an "already exists" abort every conflicting target was deleted and re-imported, even one newer than the source. It now keeps a target that is already up to date.
- **Folder and file permission failures never showed anywhere** — folders have no result row and per-file errors were swallowed by bare `catch` blocks, so a failed break-inheritance or an unresolvable principal read as a clean Success. They now appear in the activity log or on the row's Permissions column.
- **A throttled user or group lookup was remembered as "not found" for the whole run**, dropping that person's grant on every later item; "not found" is now remembered for 90 seconds only.
- **List items were matched by Title** — two items titled "Q1" overwrote each other on a re-run, and blank-title items were created again every run even in Skip mode. Duplicates now pair up by position.
- **List-item Created/Modified skipped the site-local date conversion** the file and folder paths use, so the date was rejected or shifted by the site's UTC offset.
- **A failed time-zone lookup was cached as the UTC fallback for the rest of the session**, pinning that instant (often the unchanged Created date reused across every version) to a value off by the site's offset; failures are no longer cached and are retried once.
- **Deep Verify turned a date mismatch into a Match** — an Office file with identical content but a lost Modified date showed green; it now stays a Date mismatch with a note. A 0-byte target with a stamped date no longer passes as Match.
- **Enhanced REST opened large-file downloads before waiting for a transfer slot**, so queued files held idle connections the server dropped (and that path has no resume). The slot is taken first, and downloads honor cancellation.

---

## 3.6.0 — 2026-10-02

### Added

- **Localization in 30 languages** — the same 30 as the Smart Permissions web part (Arabic, Chinese Simplified/Traditional, Czech, Danish, Dutch, English, Finnish, French, German, Greek, Hebrew, Hindi, Hungarian, Indonesian, Italian, Japanese, Korean, Norwegian Bokmål, Polish, Portuguese Brazil/Portugal, Romanian, Russian, Spanish, Swedish, Thai, Turkish, Ukrainian, Vietnamese). The app follows the Windows display language by default; pick a specific language under Settings → Language (applies on next launch). Arabic and Hebrew switch the windows to right-to-left layout. Translations are machine-generated and pending native-speaker review.

---

## 3.5.3 — 2026-08-28

### Fixed

- **Migration API mode's custom-column write was overwriting each file's real Modified date and Editor with the migrating account and the current time** — applying custom column values goes through `ValidateUpdateListItem`, which always re-stamps `Modified`/`Editor` as a side effect of the write, even though SPMI's own import had already set the correct values from the source file. The post-import custom-fields pass now re-stamps `Modified`/`Created`/`Editor`/`Author` from the source file's own metadata in the same write, so the correct values survive the custom-column write instead of being clobbered by it.
- **That re-stamp landed on the wrong time of day** — `ValidateUpdateListItem`'s plain `"yyyy-MM-dd HH:mm:ss"` date format (no timezone designator) is parsed by SharePoint in the *site's* regional-settings time zone, not UTC, so submitting a UTC instant directly shifted it by the site's offset (e.g. a real 6:33am UTC modified time came back as 1:33pm on a UTC-7 site). Dates are now converted through the site's own `regionalsettings/timezone/UTCToLocalTime()` REST function before being submitted, so there's no separate zone/DST table to get out of sync with what SharePoint itself will use to parse the value back.

### Changed

- **Migration API custom columns + metadata restamp now cost far fewer REST/Graph calls per file** — the Modified/Editor re-stamp above is folded into the *same* `ValidateUpdateListItem` call as the custom-column write instead of a separate call afterward (one list-item lookup and one write instead of two), and the source file's Modified/Created/Editor/Author values are read once per library in the same bulk scan that already reads custom column values, instead of a live per-file lookup.

---

## 3.5.2 — 2026-08-26

### Added

- **Migration API mode now applies custom column values** — previously this mode explicitly skipped custom columns for non-office files entirely and told the user to re-run in Enhanced REST mode to stamp them (the SPMI manifest's per-item `<Fields>` support had been removed earlier after it caused "Missing file info" import failures). A new post-import REST pass now applies custom columns after SPMI import completes, resolving each target item by its server-relative path since the migration API doesn't surface Graph item IDs for imported files — the same underlying mechanism Enhanced REST mode already used per file.
- **Custom column failures are now visible instead of silently reported as "Success"** — a new "Custom Fields" status column (with its own Details column) appears in both the live Copy Log grid and the Copy Report grid, alongside a "Metadata Warnings" summary card on the Copy Report screen and a single end-of-run activity log line. A custom-field mismatch is tracked as its own amber "Warning" status, deliberately separate from the main Failed/Success status — the file itself still copied correctly, only some metadata didn't make it across, so treating it as a full copy failure would overstate the problem.

### Fixed

- **A single unreadable custom column emptied the entire bulk field cache for every file in the run** — `BulkReadCustomFieldsAsync` chunks columns into batches of 20 for one REST call each; if any one column in a batch wasn't actually selectable at the item level (observed with `_ExtendedDescription`, a field that appears in the fields list — not hidden, not read-only — but which SharePoint's item REST API rejects with "The field or property does not exist"), the whole batch's HTTP 400 wiped out every other column's values for every item, with the failure silently swallowed by the caller. The read now drops just the offending field and retries, and reports what was skipped instead of failing silently.
- **A single mismatched custom column blocked every other column from being applied to the same item** — if a submitted field name doesn't exist on the *target* list (e.g. a column with the same display name on both libraries but a different auto-suffixed internal name — observed with a "Notes" column), `ValidateUpdateListItem` rejects the entire request with an HTTP 400, not just that one field. The write now drops just the offending column and retries with the rest, so one column mismatch no longer blocks otherwise-valid columns like Status, Count, or Owner from being written.
- **A custom field silently absent from `ValidateUpdateListItem`'s response was reported as a successful copy** — SharePoint only returns an entry for fields applicable to the item's actual content type, so a submitted field for a column the content type doesn't include (e.g. a PDF assigned a different content type than Office files in the same library) was previously dropped from the response with no error, rather than flagged with `HasException`. A submitted field now missing from the response is treated as a failure.
- **The "Custom Columns" checkbox didn't persist across sessions** — every other file-copy option (versions, permissions, preserve metadata, etc.) is restored from saved settings on startup; this one was saved on exit but never read back in, so it silently reset to unchecked every time the app was reopened.

---

## 3.5.1 — 2026-08-04

### Changed

- **Folder metadata lookup failures are now reported with their real cause** instead of a generic "Graph lookup failed after retries" guess — the fetch pass used to swallow every exception silently, so a genuinely throttled/missing item and a folder the scan simply never identified looked identical in the log. The activity log now distinguishes "no identity from the scan" (the folder was never part of this job's own walk — expected for some manually-retargeted single-folder copies) from an actual Graph exception, and shows the real exception text for the latter.
- **The source-scan phase now reports its own throttle overhead** — total throttle events, time waited, and a wait/elapsed ratio alongside "Source scan complete," plus concurrency step-down/step-up lines matching the ones already shown for Analysis/Downloads/Uploads — so an unusually slow scan is diagnosable from its own log instead of requiring after-the-fact timestamp arithmetic.
- **Estimated time remaining no longer jumps between wildly different values tick-to-tick** — a momentary rate spike (e.g. one large file finishing) previously could swing the displayed estimate from "1 minute" to "48 minutes" and back on the very next tick. Both the main and packaging ETAs are now smoothed with an exponential moving average tied to the same 20-second window already used for the underlying rate calculation.

### Fixed

- **A folder with no subfolders of its own never received its date/author/color correction in Migration API mode**, even with "Re-apply folder metadata every run" on — a leaf folder holding 11 files with no nested directories was entirely absent from the scan's per-job folder-identity map, because that map was only created the first time a *subfolder* was found during the job's walk. A job whose walk found files but zero subfolders never got an entry at all, so its own top-level folder — not just a missing ancestor — silently kept the placeholder "Dec 31, 1999" date on every run. This is a different gap from the 3.5.0 fix (which covered a folder whose walk *did* find subfolders elsewhere but never recorded the folder's own identity); together they should cover every folder shape.

---

## 3.5.0 — 2026-08-02

### Added

- **Total Size shown alongside item counts** — the total bytes involved are now shown next to the existing item counts throughout the app: the Copy Preview header, the in-progress copy line, a new "Total Size" tile on the Copy Report screen, the Excel verification report's Overview sheet, and each run's History summary. Once a run finishes, this figure also accounts for every version actually copied when "Copy versions" is on, not just each file's current version — reusing the same per-file size data the app already tracked for its time-remaining estimate, so no extra Graph calls are needed. Left blank whenever too many rows lack a knowable size (e.g. Library/Site-scope or permission-only copies) rather than showing a misleading partial number.
- **Report filenames now recognize OneDrive personal sites** — a source or target URL under `/personal/{username}/` is now prefixed onto default report filenames the same way `/sites/`/`/teams/` site names already were.

### Changed

- **Site/user names used in report filenames are now percent-decoded** — a non-ASCII name (e.g. an accented or non-Latin site or OneDrive user name) previously showed up in filenames as a raw `%XX`-escaped string instead of its real characters.

### Fixed

- **OneNote (.one/.onetoc2) false-positive "Date Mismatch"** — a live verification report showed every target file's Modified date landing at the copy run's own timestamp regardless of what was preserved, because OneNote's server-side consolidation re-stamps Modified along with the bytes whenever it rewrites a section file, unlike other Office formats it's grouped with for the same hash-unreliable symptom. Date isn't a stable fallback signal for these two extensions either, so they're now reported "Unverified" with an explanatory note instead of a false alarm.
- **A pre-existing destination folder's Modified date could be corrupted to "Dec 31, 1999"** in Migration API mode — copying into an existing container folder that wasn't itself part of the copied source tree left no source-side metadata to stamp onto its SPMI manifest entry, so it fell back to a hardcoded placeholder date that SharePoint's import then applied for real, overwriting the folder's actual date. The container folder's own current dates are now used instead, making that manifest entry a no-op.
- **History's per-run summary could silently clip its own text** — the Total Size figure above was invisible for any run whose summary line grew long enough to overflow the fixed-width run list, since the list's default horizontal scroll region clips overflow with no wrapping or ellipsis to indicate it. The summary now wraps instead of overflowing.

---

## 3.4.1 — 2026-07-29

### Added

- **Folder color preserved** — when Preserve Metadata is on, a source folder's custom color (SharePoint's "Customize folder > color" feature) is now read and stamped onto the corresponding target folder, alongside the existing Created/Modified/Author/Editor repair. Applied via SharePoint's dedicated color-stamping endpoint, before the date/author correction pass runs — reversed order silently reverted Author/Editor to the importing account on every colored folder, since the color endpoint's own write also stamps Editor/Modified. Read/write are best-effort: an unsupported tenant or renamed internal field name never disables the rest of the metadata repair, and color failures are counted and logged separately so they don't inflate "could not correct metadata."

### Changed

- **Estimated time remaining is now weighted by file size (bytes), not item count** — a run with a burst of small files followed by a few large ones previously read as "almost done" right up until those large files started, since finishing 90% of the *files* said nothing about how much data was actually left. The estimate now tracks a recent-throughput window (falling back to the full-run average only when that window has no data yet) and shows "Calculating…" instead of a number until at least ~15 seconds and ~2% of the job have been observed — long enough to get past the upload concurrency controller's deliberate ramp-up, which previously got extrapolated into wildly high early estimates (e.g. a 20-30 minute job briefly reading "3 hours"). Falls back to the previous item-count method for scopes that don't carry per-file sizes (library/site copy, permission-only rows).
- **Copy-If-Newer re-runs are dramatically cheaper in Enhanced REST mode** — a per-target-folder existence/date snapshot (one listing per folder) replaces two Graph calls per *file*; on an all-skip re-run of a 100,000-file library this cuts roughly 200,000 round trips down to about 10,000.
- **SharePoint REST/CSOM calls needed to correct a folder's metadata are now cached instead of re-fetched per folder** — the form digest (site-scoped, ~30 min valid) and each resolved user's SharePoint ID are now reused across the whole run instead of being re-requested for every single folder, cutting a documented 39-minute correction phase by roughly 40%.
- **SPMI folder metadata now reuses folder identities the source scan already discovered**, instead of a separate re-walk of the whole tree (Enhanced REST) or an ancestor "hops up" derivation from a nearby file (Migration API) — eliminates the redundant walk entirely and removes several thousand extra Graph calls per run on a large, deeply-nested library. This also fixes a real correctness bug: the old ancestor-derivation could occasionally resolve the wrong folder — including, in one observed case, stamping the *target library root itself* with an unrelated folder's metadata.
- **Enhanced REST's folder-metadata pass is now throttle-aware** — it previously ran at fixed width with no backoff at all, immediately after the copy phase had already depleted the tenant's throttle budget.
- Several large-scale bookkeeping structures (per-file result counters, per-folder path strings, version-list lookups during replay, the 250,000-task startup for a very large copy) were replaced with versions that don't re-scan or re-allocate at that same wasteful scale — cumulatively, meaningfully less CPU/memory overhead on very large (100,000+ file) runs, with no behavior change.
- **"Copy all versions" (or a version cap above 250) is now clamped to Migration API's real per-file version ceiling** instead of silently forming a batch that fails every import attempt and reproduces itself identically on retry; the activity log now says so explicitly when it happens.
- **Migration API's prep/import pipeline no longer leaves work running unobserved after a failure** — if either side of the download→package or package→import pipeline failed unexpectedly, the other side previously kept running (or hung waiting on a channel nobody was draining anymore) after the failure had already started tearing down the resources it depended on.
- **Library and Site-scope copies now correctly track "updating metadata" while their background folder-correction passes are still running** — previously the wizard could report a run complete (and allow navigating away or closing the app) while folder dates/authors/colors for one or more libraries were still being written in the background.

### Fixed

- **A migration job's own top-level folder no longer shows a placeholder "Dec 31, 1999" date in Migration API mode** — the source scan only ever recorded folder identities for folders it found *underneath* the selected item, never the selected folder itself, so that one folder could never be matched up with its real Created/Modified metadata and silently fell back to the SPMI manifest's default placeholder date, even though every descendant folder underneath it got its correct date.
- **Permission-flag bulk read no longer silently truncates on a failed page** — on a library with 5,000+ items, a throttled page partway through previously made every item past that point copy with inherited permissions and no error, while the run still reported Success.
- **A page copy that fails partway (throttled REST call, or a missing site URL) is no longer reported as Success** — it previously left an empty stub at the target that counted as a success and was skipped on a later Copy-If-Newer re-run.
- **A fully-copied file whose post-copy permission step is interrupted no longer gets overwritten to "Cancelled"** — the file's real Success/Skipped outcome is now preserved, and the cancellation (if any) is reported separately.
- **A single throttled lookup-column resolution no longer poisons that value for the rest of the session** — every other file referencing the same lookup value previously lost that column silently, for the remainder of the run.
- **The folder-metadata pass no longer reports "complete" after abandoning every remaining folder following one failure** — a single folder's throttle-exhausted correction used to silently end the whole pass early with no error.
- **Permission read failures are now distinguished from "this item genuinely has no unique permissions"** — previously reported identically as a clean, zero-item success.
- **A failed lookup during per-version replay is now surfaced as an error** instead of silently leaving a duplicate phantom version behind with no indication anything went wrong.
- **Existing-file listings during an Overwrite pre-flight no longer silently truncate on a failed page** — a partial listing could previously cause SharePoint's Migration API to reject files as "already exists" that the pre-flight should have already cleared.
- **Folder existence/delete checks now distinguish "genuinely not there" from a transient failure**, matching a fix already applied to the equivalent file-level checks — a locked or checked-out special folder (e.g. a OneNote notebook) now reports its real cause instead of a misleading "already exists."
- **Site-scope Site Pages copies no longer silently drop custom column values** due to a cache-key mismatch specific to that scope; Library and Files-scope copies were unaffected.
- **A truncated page during version-count batching can no longer silently under-count a file's versions**, which could otherwise push that file's batch over the Migration API entry ceiling and fail its import.
- **"Re-apply folder metadata every run" now actually has an effect in Enhanced REST mode** — it previously only worked for Migration API mode; an Enhanced REST re-run with the option checked silently did nothing.
- **Enhanced REST file transfers are now bounded by the same memory budget and large-file gate Migration API already had** — previously unbounded, so many large files copying in parallel could exhaust process memory the same way a documented Migration API incident once did.
- **Newly-created empty folders in Enhanced REST mode now receive their metadata pass** — previously created with today's date and the copying account's name, never corrected afterward.
- Fixed as a side effect of the counter/lookup rework above: a folder-level permission failure can no longer be misattributed to an unrelated file that happens to share its name.
- **Copying a page whose filename contains `#`, `&`, or `+` no longer fails to be found on the source** — the lookup wasn't URL-encoding the filename before building the request.
- **The custom-column definition cache is now cleared at the start of every run** — a column added or retyped on the target between two runs (without restarting the app) was previously still resolved against the stale definition from the earlier run.
- **Date parsing during folder-metadata reads is now pinned to invariant culture**, closing a theoretical misparse on a non-Gregorian regional setting.
- **Cancelled rows now show a status in the copy log grid** instead of a blank cell.

---

## 3.4.0 — 2026-07-23

### Added

- **Compare dialog** — a new "Compare" button in the app bar opens a standalone dialog that connects to any source and target location, lets you pick a library or folder on each side, and generates a difference report — independent of running a copy or re-verifying a History entry. Produces the same Excel report shape (Overview/Source/Target/Comparison/Scan Errors) as the existing verification report.
- **"Re-apply folder metadata every run" option** — a new Options checkbox (on by default) controls whether folder Created/Modified dates and Author/Editor are re-checked for every folder on every run, or only for folders that actually receive new files this run. Turning it off makes repeated "mostly skipped" incremental runs significantly faster on large libraries. The run-settings summary now reports which mode was used.
- **Cancelled status** — items still in progress when a run is cancelled or the app is closed mid-copy are now reported as Cancelled rather than Failed, so an interrupted run's saved report no longer looks like a mass failure (e.g. thousands of untouched items previously showed up as "failed").

### Changed

- **Large-file transfer reliability overhaul** (Migration API mode): a shared, self-tuning upload-concurrency gate replaces the old fixed per-batch cap, which allowed dozens of simultaneous multi-GB uploads and caused a high failure rate ("Error while copying content to a stream") on large libraries; a new global in-memory byte budget (sized to ~40% of machine RAM, 2–16 GB) bounds total in-flight file payloads across every concurrent batch, preventing multi-GB heap growth and connection-reset storms on libraries of very large files; the first couple of batches are now deliberately small so the first SharePoint import starts within minutes instead of after hours of packaging; and a periodic resource snapshot is now logged so a native WPF crash still leaves a trail of what the process looked like beforehand.
- **Wider prep pipeline** — up to 8 batches (was 3) can now package concurrently, keeping the shared download slots fed even when a batch is stuck waiting on one slow large file.
- **User-Agent decoration** — Graph and SharePoint REST calls now send a structured, company-neutral User-Agent (`NONISV|SharePointSmartCopy|<version>`), since undecorated traffic is throttled more aggressively by SharePoint Online.
- **Throttle-aware retries** — download/upload throttling (429/503) is now handled with patience proportional to the server's Retry-After, and a throttle window discovered during one phase now carries over to the next phase instead of each one re-discovering it independently.
- **Deep Verify report naming** — exported verification reports get a "Deep" filename prefix and on-screen status/activity text says "Deep verification" when the deep Office-file comparison pass is enabled, so a deep run is distinguishable from a standard one at a glance.
- The main wizard's Step 1 (Browse Source) and Step 2 (Target folder) now scroll as a whole on small windows, instead of clipping the bottom of the folder tree.

### Fixed

- **Target-site mismatch** — editing the Target URL after already connecting, without an explicit Disconnect first, could copy into the previously-connected (and now unrelated) site while the screen showed the newly-typed URL.
- **OneNote (.one/.onetoc2) false-positive comparisons** — these files were flagged as content mismatches based on size alone despite being unchanged; they now use the same date-based fallback already applied to other Office formats SharePoint re-serializes internally.
- **Incorrect "Modified By" stamp on a directly-selected source folder** — caused by an unconditional ProgID probe that ran against the source root even for ordinary, non-special folders.
- **Empty folders not created at the target** — a folder with no files anywhere in its subtree was previously never created, since every other folder was only provisioned as a side effect of copying a file into it.
- **Empty folders always reported "Success," even on a fully up-to-date re-run** — they now correctly report "Skipped" when the target folder already exists.
- **"Re-apply folder metadata" toggle had no effect for ordinary Files/Folders-scope copies** — it only worked for Library/Site-scope and Pages copies; a missing argument at one of four call sites meant Files-scope runs always repaired every folder regardless of the setting.
- **Large heap growth / connection-reset storms on large-file libraries** — files whose size fetch failed under throttling fell back to a size of 0, letting multi-GB files bypass the large-file memory gate entirely; they now fall back to the size already captured during the initial scan.
- **WPF render-thread crash (`UCEERR_RENDERTHREADFAILURE`) on long migration jobs** — high-concurrency background threads posting UI updates via a blocking dispatcher call could overwhelm WPF's composition engine; updates are now posted asynchronously instead.

---

## 3.3.1 — 2026-07-14

### Added

- **Default report filenames now include the source and target site names** — e.g. `Marketing-Archive-CopyReport_Files_20260714_101500.csv` instead of a bare timestamp, for both the live Copy/Verification report exports and History's CSV/Verification exports. Falls back to the previous timestamp-only name if either site's URL doesn't contain a `/sites/` or `/teams/` path segment. Configurable via a new **Settings** checkbox (on by default).

### Changed

- **Migration API throttle events are now always logged** — a throttle landing inside the adaptive gate's cooldown window, or one Kiota's own retry handler quietly absorbs before `GraphThrottleNotifyHandler` reports it, previously produced no log line at all, showing up only as an unexplained stall. Rate-limited to one line per 5 seconds so a throttle storm doesn't flood the activity log.
- **Azure AD setup docs now spell out that admin consent is per-tenant** — USER-GUIDE.md and README.md previously gave inconsistent single-/multi-tenant guidance and never mentioned that consent granted in one Microsoft 365 tenant has no effect on another; both now recommend single-tenant, and the Troubleshooting sections include a concrete checklist (confirm Global Administrator role, confirm both permission blocks show "Granted," the direct admin-consent URL shortcut) for the "Approval Required"/"Need admin approval" screen.

### Fixed

- **OneNote notebooks no longer copied as a broken plain folder** — SharePoint identifies a OneNote notebook (and similar container items) via a `package` facet that Microsoft Graph often returns with no populated `folder` facet alongside it; the app's folder detection previously missed these entirely, so scanning, browsing, and file-by-file copying all treated a notebook as an ordinary file or a misclassified folder, silently losing the notebook association that SharePoint's own "Copy to" preserves. Special folders are now detected up front and copied as a single native server-side Graph operation instead — in Migration API mode, a post-import correction pass (mirroring the existing folder Author/Editor/date fix) additionally repairs the folder's `ProgID` via CSOM, since SPMI's manifest doesn't honor it on import.
- **Migration API job-submission failures now appear in the activity log** — unlike every other failure path in `MigrationJobService`, an exception thrown while submitting the import job itself (before a job ID was even obtained) previously marked the batch's files Failed with no explanation written anywhere, making a dead-on-arrival batch look identical to a silent hang.
- **Settings dialog no longer silently resets "Deep Verify Office files" to off** — saving Settings for any reason (e.g. just switching the theme) rebuilt the persisted settings object without carrying that preference over, discarding the user's choice on the very next save.

---

## 3.3.0 — 2026-07-10

### Added

- **Deep Verify Office files** — An opt-in checkbox next to Verify in the History dialog downloads both copies of any modern Office file (`.docx`/`.xlsx`/`.pptx`/`.vsdx` and variants) whose hash or modified date disagreed during verification, and compares their actual internal content part-by-part, ignoring only the specific parts SharePoint itself rewrites on upload (Document ID stamps, custom document properties, Excel's `calcChain.xml`, and similar). A file that's genuinely unchanged is reported as an ordinary **Match** — with a Note explaining it was confirmed by deep verify — rather than a separate status a user has to learn. Off by default; always verifies every flagged file with no count cap or time limit, so a run can be left to fully complete.
- **Verification Overview sheet leads with a one-glance headline** — "✓ ALL FILES MATCH" or "⚠ CONTENT DOES NOT MATCH", with the matched count and a plain-language summary right below it, ahead of the detailed count breakdown. Previously the sheet led with eight separate counts (including deep-verify-specific rows) and buried the pass/fail read at the bottom.
- **Verification report distinguishes "unverifiable" from "match"** — a new **Unverified** status covers rows where neither a content hash nor a fallback file size is available on both sides. Previously a missing signal was silently reported as Match, which could let a 0-byte or corrupt file pass green. Non-Office files missing a hash (Microsoft Graph omits `quickXorHash` for a nontrivial share of listed items) now fall back to a file-size comparison before giving up.
- **Verification's Office-reserialization exception list greatly expanded** — now covers Excel's binary-sheet `.xlsb` variant, all modern Visio formats (`.vsdx`/`.vsdm`/`.vssx`/`.vssm`/`.vstx`/`.vstm`), and legacy binary Visio/Publisher/Project (`.vsd`/`.vst`/`.vss`/`.pub`/`.mpp`) — confirmed via a real run that any Office container format missing from this list produces the same false Content Mismatch pattern as the `.xls`/`.msg` case that motivated the original list.
- **Column-creation failures now appear in the run report** — previously only written to the Debug output, invisible to the user; a library or list copy with a column that failed to create now says so in its result row.

### Changed

- **Migration API batching now also budgets by total bytes, not just item/version count** — a 250-item batch of large or version-heavy files could previously produce a multi-GB package; a 10 GB per-batch cap now applies alongside the existing count limits, reducing import time, memory pressure, and the blast radius of a fatal batch abort.
- **Migration API concurrent-import width tuned from 6 to 4** — SharePoint queues rather than truly parallelizes beyond roughly 2 concurrent imports per site collection, and 6-wide showed more server-side "Operation canceled" soft-aborts with a larger blast radius per conflict retry.
- **A single file version larger than 2 GB now fails fast with an actionable message** ("copy this file with Enhanced REST mode") instead of being downloaded three times and eventually surfacing a confusing "Stream was too long" as if it were a connection problem — Migration API buffers each version fully in memory, which has a hard 2 GB ceiling.
- **Large-file download throttling now considers a file's total size across all its versions, not just the current version** — all versions download and are held in memory concurrently until upload, so a small current version with gigabytes of prior history was exactly the out-of-memory case this gate exists to prevent, and previously slipped through it.
- **Large-file upload slice size increased from 320 KiB to 10 MiB** — a multiple of Graph's required 320 KiB granularity; cuts the round-trips for a 1 GB file from roughly 3,200 to about 100.
- **Folder/file date fields written via `ValidateUpdateListItem` now use unambiguous ISO 8601 instead of a locale-dependent format** — the old `M/d/yyyy H:mm:ss` format transposed day and month on non-US regional settings for dates on or before the 12th, and a bare time was interpreted in the target site's local time zone, silently shifting stored Created/Modified dates by the UTC offset and corrupting later Copy-If-Newer comparisons.
- **Excel report writing and History dialog run-detail loading now happen off the UI thread** — both previously froze the window, sometimes for many seconds, on a 100,000+ item run.
- **Graph 504 Gateway Timeout responses now trigger the same throttle step-down as 429/503** — previously absorbed invisibly by the underlying retry handler with no adaptive backoff. The default retry delay when no `Retry-After` header is present also dropped from 60s to 10s, since a bare 503/504 without one is usually a transient blip, not a real sustained throttle, and used to freeze all new work for a full minute regardless.
- **The run-settings summary now reflects what actually ran, not just what was selected** — Pages scope silently forces Enhanced REST with version copying off regardless of the Copy Mode/Copy Versions controls' displayed state; the summary line now reports that override instead of the pre-override selection.

### Fixed (Deep Verify)

- **False Content Mismatch on `xl/_rels/workbook.xml.rels` for unchanged Excel files** — Excel's calculation-chain cache (`xl/calcChain.xml`) was already excluded from comparison, but the relationship entry referencing it inside `workbook.xml.rels` is added or removed whenever a recalculation happens on only one side — a real structural change to that file even though it's the same already-ignored volatile artifact. Relationship (`.rels`) parts are now compared as an order-independent set of relationship type/target pairs with calcChain entries filtered out, so this no longer flags a genuinely identical file while still catching an actually-changed relationship (e.g. a removed worksheet).
- **False Content Mismatch from a SharePoint-added `customXml` relationship (and generalized against future cases)** — SharePoint stamps a Document ID into a new `customXml/itemN.xml` part on upload; that part's content was already excluded from comparison, but the relationship entry referencing it from `xl/_rels/workbook.xml.rels` wasn't — the same class of issue as the calcChain fix above. Rather than add another one-off exclusion, relationship targets are now resolved against their owning part's folder and filtered out whenever they point at a part already treated as volatile (customXml, calcChain, etc.), so any future relationship pointing into an already-excluded part is handled automatically instead of needing its own fix.
- **False Content Mismatch from a missing vs. explicit `TargetMode="Internal"` in relationship files** — `TargetMode` is optional in the OOXML spec and defaults to Internal when omitted; a source `.rels` file that omits it and a re-saved target `.rels` that writes it explicitly were being treated as two different relationships for an otherwise-identical entry. This was the actual cause of the `xl/_rels/workbook.xml.rels` mismatches surviving the calcChain fix above, since it affects every relationship-bearing `.rels` part, not just ones with a calcChain entry. Also: a `.rels` mismatch's Note now lists the specific differing relationship entries instead of just the part name, so any future false positive is diagnosable directly from the report.
- **False Content Mismatch from Excel's `calcId` recalculation stamp** — `xl/workbook.xml` contains a `<calcPr calcId="..."/>` attribute Excel updates on every recalculation (including one silently triggered by SharePoint's own preview/co-authoring pipeline), regardless of whether any actual value changed — the same class of noise as the `calcChain.xml` issue above, just inside `workbook.xml` itself. This attribute (and the similarly-volatile `fullCalcOnLoad`) is now ignored when comparing.
- **Verification status message no longer gets cut off** — it previously shared a narrow, fixed-width column with the buttons in both the main window and the History dialog, so a long message (deep-verify progress, throttle notices) was either truncated with an ellipsis or pushed the Cancel button out of view depending on which fix was in place. The status message now gets its own full-width row below the buttons and wraps instead of truncating, so the complete message is always visible.
- **Verification status no longer looks frozen/stalled during a Deep Verify pass** — the status line only ever showed the *scan's* file-count progress; once scanning finished and Deep Verify started downloading and comparing files (which can easily take longer than a few seconds per file), the line stayed stuck on the last scan text with no visible sign anything was still happening. Deep Verify's own progress messages ("N need deep verification", "Deep-verifying: N / total", "complete") now update the status line directly; only genuine transient notices (throttle waits, scan errors) still auto-clear after a few seconds rather than sticking around.
- **False Content Mismatch for legacy VML drawings** (`.vml` files used for cell comments and form-control drawings in `.xlsx`) — these are XML but weren't recognized as such, so they fell through to a strict raw-byte comparison instead of the same structural comparison every other XML part gets, flagging insignificant formatting differences from a resave as real content changes.
- **False Content Mismatch from Excel's saved window/view state** — `<bookViews>` (in `xl/workbook.xml`) and `<sheetViews>` (in each worksheet) store window position/size, the active tab, and selected cell as of the last save — not calculated content — and are regenerated by whatever process or machine last touched the file. Both are now ignored when comparing.
- **Verification report no longer shows a confusing blank cell for a file compared by size** — when Microsoft Graph doesn't return a content hash for a file (a known, already-handled occurrence), the comparison already correctly falls back to comparing file size — but the report gave no indication of this, showing an empty Source/Target Value next to a Content Mismatch verdict that looked like a broken comparison. The report now shows the actual size compared for both sides (e.g. "12,345 bytes (by size — hash unavailable on at least one side)") whenever either side lacks a hash, rather than showing a hash for whichever side happens to have one next to a size for the side that doesn't.

### Fixed

- **Folder Created By / Modified By now actually preserved in Migration API mode** — SharePoint's Migration API silently never honored `Author`/`ModifiedBy` on `<Folder>` manifest elements, even on a brand-new folder creation; every migrated folder showed the importing account regardless of the source's real author. A post-import correction pass now sets each folder's Author/Editor directly against the target via SharePoint's CSOM `UpdateOverwriteVersion()` mechanism — the only SharePoint update mode that can set these fields to a user other than the one running the copy. Dates were already being corrected by the existing metadata pass; this closes the remaining gap for the person fields.
- **Ancestor-only folders no longer show a placeholder 1999/2000 date** — a folder containing only subfolders (nothing loose directly inside it) previously never received a metadata entry, so it fell back to a hardcoded placeholder date rather than its real created/modified date. Metadata is now resolved for every ancestor folder in the selection, borrowing a descendant file's metadata and walking up the folder chain when a folder has no files of its own.
- **Already-completed migrations can now have folder metadata corrected without re-copying files** — re-running a finished Migration API copy in "If newer" mode previously skipped the entire folder metadata pass whenever every file was already up to date (nothing to import). The correction pass now runs on every re-run regardless, so a 100,000+ file migration with wrong folder dates/authors can be fixed with a fast, all-skip re-run instead of a full re-transfer. To keep that capability from slowing down every future routine re-run once metadata is already correct, each folder's current Author/Editor/Created/Modified is now checked with a single lightweight read first — a folder that's already right is confirmed in one call instead of paying the full correction cost every time.
- **Folder metadata correction no longer looks like a stall on a large tree** — this pass reported nothing between its start message and its final result, so a run with thousands of folders under sustained throttling could show 30+ minutes of nothing but "Graph throttled" lines with zero indication anything was still progressing (observed on a real run: 39 minutes silent for 3,813 folders). It now reports "Correcting folder metadata: N / M" every 100 folders, matching the progress reporting the folder metadata *fetch* phase already had.
- **Folder Modified By no longer intermittently attributed to the importing account instead of the real source author** — `UserGroup.xml` (which registers every user referenced in the import package) was being built before folder authorship was registered into the manifest, so a user who appeared only as a folder's Author/ModifiedBy — never as any file's Author/Editor in the same batch — had no corresponding entry SharePoint could resolve. Most likely to appear on a small or single-file re-copy; large, author-diverse batches happened to register the same users via their files anyway.
- **Verification's Cancel button no longer hidden at the History dialog's default window size** — it shared a stretchy layout column with an unbounded-length status line (file counts, throttle notices), which could push the button out past the visible area while the button toolbar next to it kept its full fixed width. Cancel is now always positioned first in that column, and the status text is capped with a tooltip for the full text when truncated.
- **Concurrent sign-in no longer spawns multiple browser prompts** — dozens of in-flight requests hitting an expired token at the same moment used to each independently attempt a silent refresh, and if that failed (Conditional Access policy, revoked session) each one separately fell through to an interactive prompt, triggering MSAL's concurrent-interactive-request exception. Token refresh is now single-flighted — one request refreshes, the rest wait and reuse its result.
- **A corrupted `settings.json` is now backed up instead of silently discarded** — previously fell back to defaults with no trace of the original file, and the very next Save() overwrote it for good, losing the user's Azure AD app registrations with no way to recover them by hand.
- **Migration API Skip-mode batch-abort retry now correctly marks conflicting files as Skipped** rather than retrying or failing them — Skip means "already exists," which is exactly what triggered the abort. Also: only files that genuinely failed go back to `Copying` for the retry pass; files already marked `Skipped` by an If-Newer "up to date" decision now survive the retry instead of being reset and needlessly re-copied.
- **Migration API import polling no longer reports a fabricated Success** when polling ends without ever seeing SharePoint's completion event (cancellation, a hung poll endpoint, a server-evicted job) — the outcome for any file still in flight at that point is genuinely unknown, and is now surfaced as an error advising a Copy-If-Newer re-run to reconcile, instead of silently promoted to Success.
- **SPMI's `.err` error report is now read completely** — SharePoint splits large error reports into numbered segments (`Import-{jobId}-1.err`, `-2.err`, …); only the first segment was ever read, so errors recorded in later segments went completely unattributed.
- **Import errors that couldn't be matched to a specific file now trigger a target-side confirmation check** instead of silently promoting every still-in-flight file to Success — if SharePoint reports more errors than the app could attribute to individual files, the unconfirmed files are now verified against the actual target before being marked successful.
- **Throttle-notification handler leaks fixed in `MigrationJobService`, `CopyService`, and `VerificationReportService`** — an unsubscribed (or only partially unsubscribed) handler accumulated on the app-lifetime `SharePointService` instance across runs, causing duplicate "Graph throttled" log lines on later runs and, in one case, a closed History dialog's UI being invoked from a run that had already finished.
- **A copy that throws mid-run no longer leaves the wizard stuck on "Updating Metadata" forever** — which also held Windows sleep-prevention active indefinitely. The metadata phase is now correctly reported as incomplete before the exception propagates.
- **Stale cached folder IDs no longer persist for the life of the app** — the folder-path-to-item-ID cache is now cleared at the start of every run (a folder deleted or renamed between runs previously kept resolving to its old, now-invalid ID), and a single transient resolution failure no longer poisons that folder's cache entry permanently.
- **Site-scope copies no longer silently drop libraries on sites with many drives** — the drives listing wasn't paginated, so only the first page of libraries was ever seen.
- **Skip mode's existing-file check no longer risks silently overwriting a file** — only a genuine HTTP 404 is now treated as "file doesn't exist"; other failures (a transient error, a request that got throttled even after retries) were previously read the same way as "missing," and in Skip mode that led straight to an upload that clobbered the real, existing file.
- **Small-file uploads now explicitly fail on a conflict** (`@microsoft.graph.conflictBehavior=fail`) instead of relying on a plain PUT, which always overwrites regardless of any prior existence check — closing a race where a file created between the check and the upload was silently replaced.
- **Large-file upload sessions now use 10 MiB slices instead of 320 KiB** — cuts the number of round trips for a 1 GB file roughly 32×.
- **Navigation link remapping no longer corrupts links to unrelated sibling sites** — a raw substring replace of the source site URL matched sites sharing the same prefix (e.g. replacing `/sites/HR` also matched `/sites/HRArchive`); replacement is now boundary-checked so only the exact site path is affected. Also now rewrites server-relative URL references (used by modern-page web parts and images), not just absolute ones, which previously stayed pointed at the source even with link remapping enabled.
- **Copying navigation now fails loudly if reading the source's existing links fails**, instead of proceeding to clear the target's navigation and rebuild nothing — a transient read failure previously wiped the target's nav with no way to know why it came back empty.
- **List item, bulk custom-field, and navigation reads no longer silently truncate on a failed page** — a failed request partway through pagination previously just stopped and returned whatever had been read so far (sometimes nothing), which could report a "successful" copy of zero items or drop custom column values with no warning.
- **Lookup columns are now recreated as real Lookup/LookupMulti fields** bound to the matching list on the target (resolved by title, since the source column's list GUID is meaningless there) — previously silently degraded to a plain Text column whose copied values then failed to resolve into anything.
- **Custom column creation no longer corrupts the internal name of a renamed column** — SharePoint derives a new field's internal name from its creation Title, not its `StaticName`, so creating with the display name gave the target column a different internal name than the source, and every subsequent value write (which targets the source's internal name) silently failed. Columns are now created with the internal name as the Title and renamed to the display title in a second step.
- **Column-mapping decisions now apply correctly to every library in a multi-library Site copy** — a check meant to detect "the mapping dialog was never opened" instead matched any library after the first, which was silently treated as having no column-creation decisions at all and got zero custom columns created.
- **Lookup-typed custom columns can now actually be read and copied** — both the bulk field reader and the per-item list reader were selecting them directly instead of `$expand`-ing them, which SharePoint rejects with an HTTP 400; the failure silently aborted the read, so lists using lookup columns previously "copied" zero items with no error surfaced.
- **Permission copying no longer risks stripping an item's permissions with nothing to restore them** — inheritance is now only broken after confirming the target's role definitions actually loaded; previously a failed/skipped role-definition fetch let every role assignment fail *after* inheritance was already broken, leaving the item accessible only to the importing account.
- **"Limited Access" role assignments are no longer copied** — this is SharePoint's own internal hierarchy plumbing, rejected when granted directly; copying it previously produced failed-role noise, and an item whose only unique assignment was Limited Access broke inheritance while applying nothing.
- **Permission copying now correctly reports failure when no principals could be resolved on the target** (e.g. cross-tenant users, deleted accounts) — previously reported as Success even though nothing was actually applied.
- **Permission copying now handles target paths containing `#`/`%`/`+`** the same way file copying already did, via `*ByServerRelativePath(decodedurl=...)` — previously these files' permissions were silently skipped.
- **Closing the History dialog during an active verification now cancels it** instead of leaving a scan of up to 100,000 files running headless in the background after the window is gone.
- **Sign-out failures no longer crash the app** — an MSAL cache exception during sign-out is now caught and shown as a message instead of propagating unhandled.
- **Navigating Back after the copy completes, changing options, then Next no longer silently re-shows the old run's results** without actually re-running with the new options — the wizard now correctly starts a fresh copy in this case.
- **Permission result rows are now matched to their file row by full target path first**, falling back to file name only if no path match is found — previously matched by name alone, which could attribute a permission result to the wrong row when the same file name existed in more than one location in the run.

---

## 3.2.0 — 2026-07-03

### Added

- **Verification Report** — Click "Verify" on any run in the History dialog (including runs saved before this feature existed, if their scope is still resolvable) to independently re-scan source and target via fresh Graph calls and produce an `.xlsx` workbook with Overview, Source, Target, Comparison, and Scan Errors sheets. Comparison status is Match / Content Mismatch / Date Mismatch / Only in Source / Only in Target, using a signal chosen per file type: most files get a genuine content-hash comparison (Microsoft Graph's `QuickXorHash`, already returned by the existing scan calls at no extra cost); Office and OLE compound-document formats — both modern (`.docx`/`.xlsx`/`.pptx` and related) and legacy binary (`.doc`/`.xls`/`.ppt`, `.msg`) — get a modified-date comparison instead, since SharePoint routinely re-serializes their internal container (ZIP for modern formats, OLE metadata streams for legacy ones) for indexing, thumbnails, and co-authoring, which changes size and hash for files that are genuinely fine but does not touch the item's official Modified date. The Comparison sheet shows the actual Source Value and Target Value compared for every row (blank on whichever side has no file), so a mismatch can be inspected directly instead of just flagged.
- **Tri-state folder/library selection** — Folders and libraries in the source tree now cycle unchecked → fully selected → indeterminate (parent deselected, its children stay selected) → unchecked, matching the existing "items only" behavior for custom lists, with a dash glyph shown for the indeterminate state.
- **Run settings summary** — The copy-progress screen and activity log now show a one-line summary of the run's configuration (Overwrite mode, Copy Mode, Versions, Parallel Copies, Preserve Metadata, Permissions, Custom Columns), captured at copy start so it always reflects what actually ran.
- **Sleep prevention during long runs** — The app now blocks Windows system sleep (display sleep is still allowed) while a copy, metadata update, or verification is in progress. Added after a 114,000-file run went to sleep mid-run and stalled every in-flight Graph retry for 95 minutes.
- **Automatic recovery when SharePoint cancels an entire import batch** — SharePoint's Migration API aborts a whole job once enough per-item name conflicts accumulate within it (observed threshold: 100), discarding every other valid file in that batch as collateral — a 250-file batch with a cluster of conflicts previously showed as 250 failures needing a full manual re-run. The app now detects this specific abort, clears the conflicting targets, and retries the batch once automatically before giving up on it.

### Changed

- **"If Newer" reuses the modified date captured during the initial scan** — Both the pre-flight skip filter and the per-batch pre-flight now check the source-modified date already captured for free during the folder walk before falling back to a Graph date fetch. On a 114,000-file run under sustained throttling, the old bulk fetch alone took 22 minutes and still left 110,000 dates unresolved — which the "undetermined → treat as needing copy" fallback then misrouted into hours of unnecessary re-copying. That fallback path is now unreachable for any file discovered by the normal folder scan. 
- **Adaptive throttle backoff extended to every Graph-heavy phase, and isolated per phase** — Pre-flight analysis (subfolder creation, existing-file scan), the source scan, the Verification Report scan, and bulk metadata/date fetches each now get their own adaptive concurrency gate, separate from the download-phase gate. Previously a shared gate meant analysis-phase throttling could pre-shrink download concurrency before any transfer started, and several loops resumed at full width immediately after waiting out a `Retry-After`, walking straight back into the same depleted budget — observed as repeated 60–120s waits back to back. Throttling now converges instead of oscillating.
- **Source folder-tree scan parallelized** — The folder walk now fans out across sibling subfolders concurrently instead of issuing one Graph call at a time. A 3,000+-folder library previously took roughly 30 silent minutes to scan before a copy could even start; progress ("Scanning source: N files found so far…") is now reported continuously.
- **Fewer Graph calls during large-run analysis** — HTTP/2 multiplexing is now enabled so concurrent Graph calls spread across multiple connections instead of contending for one; metadata fetches skip the `/versions` sub-request entirely when Copy Versions is off; and Skip/If Newer runs no longer fetch full metadata for files already identified as skippable before the batching pass.
- **SPMI batch preparation and container provisioning made more efficient** — Batch prep (download → encrypt → upload → manifest) now runs up to 3 batches concurrently instead of one at a time. Azure container provisioning is deferred until a batch is confirmed to actually have files to upload, avoiding hundreds of needless provisioning calls on large "copy if newer" runs where most batches are all-skip.
- Cancel is now available during metadata-update and verification phases, not just during the copy itself.
- The "Copying…" status label is now "Processing…", reflecting that the phase includes download/encrypt/upload, not just a network transfer.
- The live Activity Log panel on the Copying screen is shorter, leaving more visible room for the Copy Log results table below it.

### Fixed

- **Overwrite mode could fail entire batches with "already exists" errors, even on a re-copy of files the app itself had just written** — the pre-flight existing-file scan relied on Microsoft Graph, which does not reliably reflect files recently written by a Migration API import; a target folder could report as empty via Graph while genuinely containing files. The existing-file scan now cross-checks the same folder via SharePoint REST as well, and a Graph-only "is the target completely empty" fast path (previously used to skip the scan outright) has been removed — it was the single largest source of missed detections.
- **File deletion during Overwrite/If Newer now verifies removal instead of trusting a misleading success response** — recycling and purging an existing file could report success without the file actually being removed, because the purge step targeted the wrong recycle bin scope (site-collection instead of web). Deletion now operates by file ID, purges the correct bin, and follows up with an existence check and a second delete pass for any file that survives the first attempt.
- **Large files (500 MB+) no longer risk exhausting memory during Migration API copies** — each in-flight download holds the file's content in memory twice (raw, then AES-encrypted for upload); several large files downloading at once could trigger an `OutOfMemoryException`. Such files are now capped to a small number of concurrent downloads, independent of the overall Parallel Copies setting.
- **SPMI import failures on large files (>256 MB) eliminated** — SharePoint's importer requires an MD5 hash on every content blob. Small blobs get one automatically from Azure's single-request upload path, but blobs above the SDK's ~256 MB single-shot threshold upload as blocks and receive no automatic blob-level MD5, so every large file in an affected batch was rejected. The app now computes and attaches the MD5 itself for every blob upload.
- **Per-file import-error attribution fixed** — Errors were previously matched to files by `ListItemId` only, but SharePoint's import errors reference different GUIDs depending on error type ([File]-level errors carry the `FileId`, stream errors the `StreamId`). A batch with 3 MD5 errors could previously attribute only 1, leaving 2 genuinely failed files shown as Success. Every GUID a batch mints is now registered for attribution, and SharePoint's `.err` report is now fetched and cross-referenced on any batch with errors — not just fatal ones — before final status is assigned.
- **Large runs no longer fail outright on a single dropped connection** — A transient transport blip during pre-flight folder creation, existing-file scanning, or the empty-target check previously propagated out of the whole drive-group loop and was caught by a top-level handler that marked every still-pending file as Failed — on a 114,000-file run this meant nearly the entire job. These calls now retry with backoff, matching the download/upload paths.
- **Metadata-update elapsed time no longer shows an absurd duration** (observed as "~2025 years") — Migration API mode reported completion before its start-time field had been set. The start time is now captured before the copy begins.
- **App no longer lingers in Task Manager after closing** — background MSAL/Graph SDK threads outlived WPF's dispatcher shutdown; the process now force-exits on window close.
- **Verification Report scans no longer look hung during throttling** — a Graph throttle wait (observed recurring for over an hour on a busy tenant, up to 120s per wait) was previously silent. The scan status now shows a "⚠ Graph throttled — waiting Ns" notice, matching what copy runs already show.
- **Folder-metadata fetches and modified-date checks no longer go silent for minutes during throttling** — both previously showed no log output for up to ~29 minutes at a time; both now report progress on every retry round.
- **History dialog no longer pauses before opening** — Opening History previously deserialized all 50 saved reports' full per-file results synchronously before the window could even appear; on a tenant with very large (100,000+-file) runs in its history, decoding those in full just to show a one-line summary per run caused a multi-second pause with no indication anything was happening. The report list now loads only the fields needed for display, defers a run's full per-file results until that specific run is actually opened, exported, or verified, loads in the background so the window appears immediately, and reads report files as a stream instead of a fully-buffered string to avoid a redundant re-encoding step. A loading indicator now shows while the list populates.

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
