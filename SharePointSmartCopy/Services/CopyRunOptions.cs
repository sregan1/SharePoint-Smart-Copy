using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Everything CopyService.ExecuteAsync needs to know about HOW to run a copy, as one named object.
//
// ExecuteAsync used to take 26 positional parameters, called from four places in MainViewModel. A
// missing or transposed argument at one call site compiled fine and silently changed behavior (the
// 3.4.0 "missing argument at one of four call sites" bug). With a record every setting is named at
// the call site, and the five settings a run cannot sensibly omit are `required`, so leaving one out
// is a compile error instead of a quiet default.
public sealed record CopyRunOptions
{
    // ── Required: a run is meaningless without a decision on each of these ──
    public required OverwriteMode OverwriteMode { get; init; }
    public required bool CopyVersions           { get; init; }
    public required int MaxParallel             { get; init; }
    // 0 = all versions.
    public required int MaxVersions             { get; init; }
    public required CopyMode CopyMode           { get; init; }

    // ── Metadata and columns ──
    public bool PreserveMetadata                 { get; init; } = true;
    public bool ReapplyFolderMetadata            { get; init; } = true;
    public bool CopyCustomColumns                { get; init; }
    public List<ColumnMapping>? ColumnMappings   { get; init; }
    public Dictionary<string, Dictionary<string, object?>>? BulkFieldCache { get; init; }
    // Source Modified/Created/Editor/Author, bulk-read alongside BulkFieldCache (same
    // "{listId}:{itemId}" keys) — lets the Migration API custom-fields restamp reuse that
    // already-paid-for scan instead of a per-file GetFileMetadataAsync Graph call. Only consulted in
    // Migration API mode; falls back to a live per-file fetch on a cache miss.
    public Dictionary<string, FileMetadata>? SourceMetaCache { get; init; }

    // ── Pages ──
    public bool CopyPages                        { get; init; }
    public bool RemapPageWebPartUrls             { get; init; } = true;

    // ── Permissions ──
    public bool CopyPermissions                  { get; init; }
    public PermissionCopyService? PermissionService { get; init; }
    public Dictionary<string, bool>? PermissionFlags { get; init; }

    // ── Progress reporting ──
    public IProgress<bool>? OnMetadataDone       { get; init; }
    public IProgress<(int, int)>? PreflightProgress { get; init; }
    public IProgress<string>? ActivityLog        { get; init; }
    public IProgress<long>? OnFilePacked         { get; init; }
    public IProgress<(int done, int total)>? OnFolderProgress { get; init; }
}
