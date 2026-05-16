using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Microsoft.Graph.Models;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Builds the SPMI XML manifest package in memory for a batch of files.
// All blobs are AES-256-CBC encrypted with the SP-provided encryption key.
// Format per blob: 16-byte random IV prepended to the ciphertext.
// All XML elements in the migration package use xmlns="urn:deployment-manifest-schema".
public class MigrationPackageBuilder
{
    public record VersionEntry(
        string StreamId,
        string VersionLabel,
        bool IsCurrentVersion,
        DateTimeOffset LastModified,
        DateTimeOffset? Created,
        string? AuthorEmail,
        string? ModifiedByEmail,
        long Size,
        byte[] EncryptedContent);

    public record FileEntry(
        string FileId,
        string FileName,
        string FolderRelativePath,
        string ListItemId,
        DateTimeOffset Created,
        string? CreatedByEmail,
        List<VersionEntry> Versions);

    // Each manifest file uses its own schema namespace per SP Content Deployment format.
    private static readonly XNamespace NsManifest     = "urn:deployment-manifest-schema";
    private static readonly XNamespace NsExport       = "urn:deployment-exportsettings-schema";
    private static readonly XNamespace NsSystemData   = "urn:deployment-systemdata-schema";
    private static readonly XNamespace NsUserGroup    = "urn:deployment-usergroupmap-schema";
    private static readonly XNamespace NsRequirements = "urn:deployment-requirements-schema";
    private static readonly XNamespace NsRootObjMap   = "urn:deployment-rootobjectmap-schema";
    private static readonly XNamespace NsViewForms    = "urn:deployment-viewformslist-schema";

    private readonly byte[] _encryptionKey;
    private readonly List<FileEntry> _files = [];

    // Integer user IDs required by the MS-PRIMEPF schema for Author/ModifiedBy attributes.
    // Claims strings belong only inside <Fields>/<Field> value attributes.
    private readonly Dictionary<string, int> _userIdMap = new(StringComparer.OrdinalIgnoreCase);
    private int _nextUserId = 1;

    public MigrationPackageBuilder(byte[] encryptionKey)
    {
        _encryptionKey = encryptionKey;
    }

    public IReadOnlyList<FileEntry> Files => _files;

    // Adds a file with all its versions to the package.
    // versions must be ordered oldest-first; the last entry is IsCurrentVersion.
    public async Task AddFileAsync(
        string fileName,
        string folderRelativePath,
        FileMetadata fileMetadata,
        List<(DriveItemVersion version, Stream content)> versionStreams)
    {
        var fileId     = Guid.NewGuid().ToString("D").ToUpperInvariant();
        var listItemId = Guid.NewGuid().ToString("D").ToUpperInvariant();
        var entries    = new List<VersionEntry>();

        for (int i = 0; i < versionStreams.Count; i++)
        {
            var (version, content) = versionStreams[i];
            bool isLast = i == versionStreams.Count - 1;

            var streamId              = Guid.NewGuid().ToString("D").ToUpperInvariant();
            var (encrypted, origSize) = await EncryptStreamAsync(content, _encryptionKey);
            var rawId = version.Id ?? (i + 1).ToString();
            var label = rawId.Contains('.') ? rawId : rawId + ".0";
            var modified  = version.LastModifiedDateTime ?? fileMetadata.ModifiedDateTime ?? DateTimeOffset.UtcNow;
            var author    = SharePointService.GetIdentityEmail(version.LastModifiedBy?.User) ?? fileMetadata.CreatedByEmail;
            var editor    = SharePointService.GetIdentityEmail(version.LastModifiedBy?.User) ?? fileMetadata.ModifiedByEmail;

            entries.Add(new VersionEntry(
                StreamId:          streamId,
                VersionLabel:      label,
                IsCurrentVersion:  isLast,
                LastModified:      modified,
                Created:           isLast ? fileMetadata.CreatedDateTime : null,
                AuthorEmail:       author,
                ModifiedByEmail:   editor,
                Size:              origSize,
                EncryptedContent:  encrypted));
        }

        // Pre-register all user emails so UserGroup.xml is fully populated before Manifest.xml is built
        GetUserId(fileMetadata.CreatedByEmail);
        foreach (var entry in entries)
        {
            GetUserId(entry.AuthorEmail);
            GetUserId(entry.ModifiedByEmail);
        }

        _files.Add(new FileEntry(
            FileId:             fileId,
            FileName:           fileName,
            FolderRelativePath: folderRelativePath,
            ListItemId:         listItemId,
            Created:            fileMetadata.CreatedDateTime ?? DateTimeOffset.UtcNow,
            CreatedByEmail:     fileMetadata.CreatedByEmail,
            Versions:           entries));
    }

    private int GetUserId(string? email)
    {
        if (string.IsNullOrEmpty(email)) return 1;
        if (!_userIdMap.TryGetValue(email, out var id))
            _userIdMap[email] = id = _nextUserId++;
        return id;
    }

    // Builds all required XML manifest files and returns them as name→encrypted-bytes pairs
    // ready for upload to the metadata container.
    public Dictionary<string, byte[]> BuildManifestXml(
        string siteId, string webId, string listId,
        string siteUrl, string webServerRelativeUrl, string libraryTitle, string libraryServerRelativeUrl)
    {
        var manifest = new Dictionary<string, byte[]>();

        var debugDir = Path.Combine(Path.GetTempPath(), $"SPMigDebug_{DateTime.Now:yyyyMMdd_HHmmss}");
        Directory.CreateDirectory(debugDir);
        System.Diagnostics.Debug.WriteLine($"[Migration] Debug XMLs → {debugDir}");

        void Add(string name, XDocument doc)
        {
            var xml = doc.ToString(SaveOptions.None);
            System.Diagnostics.Debug.WriteLine($"[XML] {name}:\n{xml}\n---");

            // Write plaintext XML to disk for easy inspection
            File.WriteAllText(Path.Combine(debugDir, name), xml, Encoding.UTF8);

            // Scan every attribute for Int32.Parse failures
            DebugScanAttributes(name, doc);

            manifest[name] = EncryptXml(doc);
        }

        Add("ExportSettings.xml",              BuildExportSettings(siteUrl));
        Add("LookupOrAddUserNamesFromSourceSite.xml", BuildUserNames());
        Add("Requirements.xml",               BuildRequirements());
        Add("RootObjectMap.xml",               BuildRootObjectMap(siteId, webId, webServerRelativeUrl));
        Add("SystemData.xml",                  BuildSystemData());
        Add("UserGroup.xml",                   BuildUserGroupMap());
        Add("ViewFormsList.xml",               BuildViewFormsList());
        Add("Manifest.xml",                    BuildManifest(
            siteId, webId, listId, siteUrl, webServerRelativeUrl, libraryTitle, libraryServerRelativeUrl));

        return manifest;
    }

    // ── XML builders ──────────────────────────────────────────────────────────

    private XDocument BuildExportSettings(string siteUrl) =>
        new(new XDeclaration("1.0", "utf-8", null),
            new XElement(NsExport + "ExportSettings",
                new XAttribute("SiteUrl", siteUrl),
                new XAttribute("FileLocation", ""),
                new XAttribute("IncludeVersions", "All"),
                new XAttribute("ExportMethod", "ExportAll")));

    // LookupOrAddUserNamesFromSourceSite.xml is SPMI-specific with no Content Deployment schema.
    private XDocument BuildUserNames() =>
        new(new XDeclaration("1.0", "utf-8", null),
            new XElement("UserNames"));

    private XDocument BuildRequirements() =>
        new(new XDeclaration("1.0", "utf-8", null),
            new XElement(NsRequirements + "Requirements"));

    private XDocument BuildRootObjectMap(string siteId, string webId, string webRelUrl) =>
        new(new XDeclaration("1.0", "utf-8", null),
            new XElement(NsRootObjMap + "RootObjectMap",
                new XAttribute("Id", webId),
                new XAttribute("Type", "SPWeb"),
                new XAttribute("ParentId", siteId),
                new XAttribute("WebUrl", webRelUrl),
                new XAttribute("Url", webRelUrl),
                new XAttribute("IsDependency", "false")));

    // SystemData.xml: SchemaVersion must be a child element with a Version attribute,
    // not an attribute of the root. SP throws "Invalid schema version XML" if it's an attribute.
    private XDocument BuildSystemData() =>
        new(new XDeclaration("1.0", "utf-8", null),
            new XElement(NsSystemData + "SystemData",
                new XElement(NsSystemData + "SchemaVersion",
                    new XAttribute("Version", "15.0.0.0"),
                    new XAttribute("Build", "16.0.3111.1200"),
                    new XAttribute("DatabaseVersion", "11552"),
                    new XAttribute("SiteVersion", "15")),
                new XElement(NsSystemData + "ManifestFiles",
                    new XElement(NsSystemData + "ManifestFile",
                        new XAttribute("Name", "Manifest.xml"))),
                new XElement(NsSystemData + "SystemObjects"),
                new XElement(NsSystemData + "RootWebOnlyLists")));

    private XDocument BuildUserGroupMap()
    {
        var usersEl = new XElement(NsUserGroup + "Users");
        foreach (var (email, id) in _userIdMap.OrderBy(kv => kv.Value))
        {
            usersEl.Add(new XElement(NsUserGroup + "User",
                new XAttribute("Id", id.ToString()),
                new XAttribute("Name", email),
                new XAttribute("Login", $"i:0#.f|membership|{email}"),
                new XAttribute("Email", email),
                new XAttribute("IsDomainGroup", "false"),
                new XAttribute("IsSiteAdmin", "false"),
                new XAttribute("SystemId", Convert.ToBase64String(Guid.NewGuid().ToByteArray())),
                new XAttribute("IsDeleted", "false")));
        }
        return new XDocument(
            new XDeclaration("1.0", "utf-8", null),
            new XElement(NsUserGroup + "UserGroupMap",
                usersEl,
                new XElement(NsUserGroup + "Groups")));
    }

    private XDocument BuildViewFormsList() =>
        new(new XDeclaration("1.0", "utf-8", null),
            new XElement(NsViewForms + "ViewFormsList"));

    private XDocument BuildManifest(
        string siteId, string webId, string listId,
        string siteUrl, string webRelUrl, string libraryTitle, string libraryRelUrl)
    {
        XElement E(string name, params object[] content) => new XElement(NsManifest + name, content);

        var objects = E("SPObjects");

        objects.Add(E("SPObject",
            new XAttribute("Id", listId),
            new XAttribute("ObjectType", "SPDocumentLibrary"),
            new XAttribute("ParentId", webId),
            new XAttribute("ParentWebId", webId),
            E("DocumentLibrary",
                new XAttribute("Id", listId),
                new XAttribute("ParentWebId", webId),
                new XAttribute("Title", libraryTitle),
                new XAttribute("BaseTemplate", "101"),
                new XAttribute("BaseType", "DocumentLibrary"),
                new XAttribute("RootFolderUrl", libraryRelUrl))));

        // Url attributes must be web-relative (no leading site path).
        // SP prepends the target web's server-relative URL when computing the destination path,
        // so including the full server-relative URL causes doubling (e.g. /sites/x/sites/x/...).
        var webPrefix = webRelUrl.TrimEnd('/');
        var libraryWebRelUrl = libraryRelUrl.StartsWith(webPrefix + "/")
            ? libraryRelUrl[(webPrefix.Length + 1)..]
            : libraryRelUrl.TrimStart('/');

        foreach (var file in _files)
        {
            var currentVersion = file.Versions[^1];
            var fileUrl = string.IsNullOrEmpty(file.FolderRelativePath)
                ? $"{libraryWebRelUrl}/{file.FileName}"
                : $"{libraryWebRelUrl}/{file.FolderRelativePath}/{file.FileName}";

            var fileEl = E("File",
                new XAttribute("Url", fileUrl),
                new XAttribute("Id", file.FileId),
                new XAttribute("Name", file.FileName),
                new XAttribute("Version", currentVersion.VersionLabel),
                new XAttribute("ParentId", listId),
                new XAttribute("ParentWebId", webId),
                new XAttribute("StreamId", currentVersion.StreamId),
                new XAttribute("FileValue", currentVersion.StreamId),
                new XAttribute("Level", "1"),
                new XAttribute("IsCurrentVersion", "1"),
                new XAttribute("HasWebParts", "0"),
                new XAttribute("CheckOutType", "2"),
                new XAttribute("CheckOutUserId", "0"),
                new XAttribute("VirusStatus", "0"),
                new XAttribute("VirusVendorID", "0"),
                new XAttribute("DocFlags", "0"),
                new XAttribute("SetupPathVersion", "15"),
                new XAttribute("UIVersion", UiVersion(currentVersion.VersionLabel)),
                new XAttribute("MajorVersion", MajorVersion(currentVersion.VersionLabel)),
                new XAttribute("MinorVersion", MinorVersion(currentVersion.VersionLabel)),
                new XAttribute("TemplateFileType", "0"),
                new XAttribute("MetaInfo", ""),
                new XAttribute("MetaInfoSize", "0"),
                new XAttribute("InternalVersion", "0"),
                new XAttribute("BumpVersion", "0"),
                new XAttribute("ContentVersion", "0"),
                new XAttribute("CharSet", "0"),
                new XAttribute("AuditFlags", "0"),
                new XAttribute("DraftOwnerId", "0"),
                new XAttribute("Size", currentVersion.Size.ToString()),
                new XAttribute("FileSize", currentVersion.Size.ToString()),
                new XAttribute("ListItemIntId", "0"),
                new XAttribute("TimeCreated", FormatDate(file.Created)),
                new XAttribute("TimeLastModified", FormatDate(currentVersion.LastModified)),
                new XAttribute("Author", GetUserId(file.CreatedByEmail)),
                new XAttribute("ModifiedBy", GetUserId(currentVersion.ModifiedByEmail)));

            if (file.Versions.Count > 1)
            {
                var versionsEl = E("Versions");
                foreach (var v in file.Versions)
                {
                    bool isCurr = v == currentVersion;
                    versionsEl.Add(E("File",
                        new XAttribute("Url", fileUrl),
                        new XAttribute("Id", Guid.NewGuid().ToString("D").ToUpperInvariant()),
                        new XAttribute("Version", v.VersionLabel),
                        new XAttribute("IsCurrentVersion", isCurr ? "1" : "0"),
                        new XAttribute("HasWebParts", "0"),
                        new XAttribute("StreamId", v.StreamId),
                        new XAttribute("FileValue", v.StreamId),
                        new XAttribute("Level", "1"),
                        new XAttribute("CheckOutType", "2"),
                        new XAttribute("CheckOutUserId", "0"),
                        new XAttribute("VirusStatus", "0"),
                        new XAttribute("VirusVendorID", "0"),
                        new XAttribute("DocFlags", "0"),
                        new XAttribute("SetupPathVersion", "15"),
                        new XAttribute("UIVersion", UiVersion(v.VersionLabel)),
                        new XAttribute("MajorVersion", MajorVersion(v.VersionLabel)),
                        new XAttribute("MinorVersion", MinorVersion(v.VersionLabel)),
                        new XAttribute("TemplateFileType", "0"),
                        new XAttribute("MetaInfo", ""),
                        new XAttribute("MetaInfoSize", "0"),
                        new XAttribute("InternalVersion", "0"),
                        new XAttribute("BumpVersion", "0"),
                        new XAttribute("ContentVersion", "0"),
                        new XAttribute("CharSet", "0"),
                        new XAttribute("AuditFlags", "0"),
                        new XAttribute("DraftOwnerId", "0"),
                        new XAttribute("Size", v.Size.ToString()),
                        new XAttribute("FileSize", v.Size.ToString()),
                        new XAttribute("ListItemIntId", "0"),
                        new XAttribute("TimeLastModified", FormatDate(v.LastModified)),
                        new XAttribute("Author", GetUserId(v.AuthorEmail)),
                        new XAttribute("ModifiedBy", GetUserId(v.ModifiedByEmail))));
                }
                fileEl.Add(versionsEl);
            }

            objects.Add(E("SPObject",
                new XAttribute("Id", file.FileId),
                new XAttribute("ObjectType", "SPFile"),
                new XAttribute("ParentId", listId),
                new XAttribute("ParentWebId", webId),
                fileEl));

            objects.Add(E("SPObject",
                new XAttribute("Id", file.ListItemId),
                new XAttribute("ObjectType", "SPListItem"),
                new XAttribute("ParentId", listId),
                new XAttribute("ParentWebId", webId),
                E("ListItem",
                    new XAttribute("Id", file.ListItemId),
                    new XAttribute("IntId", "0"),
                    new XAttribute("ObjectType", "0"),
                    new XAttribute("FSObjType", "0"),
                    new XAttribute("ContentTypeId", "0x0101"),
                    new XAttribute("DocType", "File"),
                    new XAttribute("HasAttachments", "0"),
                    new XAttribute("ParentFolderId", listId),
                    new XAttribute("Order", "100"),
                    new XAttribute("Version", "1"),
                    new XAttribute("UIVersion", UiVersion(currentVersion.VersionLabel)),
                    new XAttribute("MajorVersion", MajorVersion(currentVersion.VersionLabel)),
                    new XAttribute("MinorVersion", MinorVersion(currentVersion.VersionLabel)),
                    new XAttribute("Level", "1"),
                    new XAttribute("WorkflowVersion", "1"),
                    new XAttribute("ThreadIndex", "0"),
                    new XAttribute("ModerationStatus", "Approved"),
                    new XAttribute("Name", file.FileName),
                    new XAttribute("DocId", file.FileId),
                    new XAttribute("ParentWebId", webId),
                    new XAttribute("ParentListId", listId),
                    new XAttribute("FileUrl", fileUrl),
                    new XAttribute("FileId", file.FileId),
                    E("Fields",
                        E("Field", new XAttribute("Name", "Author"),             new XAttribute("Value", Claims(file.CreatedByEmail))),
                        E("Field", new XAttribute("Name", "Editor"),             new XAttribute("Value", Claims(currentVersion.ModifiedByEmail))),
                        E("Field", new XAttribute("Name", "Created_x0020_Date"), new XAttribute("Value", FormatDate(file.Created))),
                        E("Field", new XAttribute("Name", "Last_x0020_Modified"),new XAttribute("Value", FormatDate(currentVersion.LastModified)))))));
        }

        return new XDocument(
            new XDeclaration("1.0", "utf-8", null),
            objects);
    }

    // ── Debug helpers ─────────────────────────────────────────────────────────

    // Integer attributes SP reads via GetInt32() on <File> elements.
    private static readonly HashSet<string> _fileIntAttrs = new(StringComparer.OrdinalIgnoreCase)
    {
        "Level", "IsCurrentVersion", "HasWebParts", "CheckOutType", "CheckOutUserId",
        "VirusStatus", "VirusVendorID", "DocFlags", "SetupPathVersion", "UIVersion",
        "TemplateFileType", "MetaInfoSize", "InternalVersion", "BumpVersion",
        "ContentVersion", "CharSet", "AuditFlags", "DraftOwnerId", "Size", "FileSize",
        "ListItemIntId", "Version", "MajorVersion", "MinorVersion",
        "Author", "ModifiedBy",
    };

    // Integer attributes SP reads via GetInt32() on <ListItem> elements.
    // ModerationStatus is SPModerationStatusType: 0=Approved, 1=Denied, 2=Pending.
    // DocType is ListItemDocType: 0=File, 1=Folder — XSD requires the string name so NOT listed here.
    private static readonly HashSet<string> _listItemIntAttrs = new(StringComparer.OrdinalIgnoreCase)
    {
        "IntId", "ObjectType", "FSObjType", "Order", "Version", "UIVersion",
        "MajorVersion", "MinorVersion", "Level", "WorkflowVersion", "ThreadIndex",
        "HasAttachments", "ModerationStatus",
    };

    private static void DebugScanAttributes(string docName, XDocument doc)
    {
        bool anyIntFail = false;
        bool anyEmpty   = false;

        foreach (var elem in doc.Descendants())
        {
            var localElem = elem.Name.LocalName;

            // Per-element int-attr set — only check File/ListItem to avoid false positives
            // from <SPObject ObjectType="SPFile"> or <SchemaVersion Version="15.0.0.0">.
            HashSet<string>? intAttrs = localElem switch
            {
                "File"     => _fileIntAttrs,
                "ListItem" => _listItemIntAttrs,
                _          => null,
            };

            // Log every attribute on ListItem for full visibility
            if (localElem == "ListItem")
            {
                foreach (var a in elem.Attributes())
                    System.Diagnostics.Debug.WriteLine(
                        $"[LISTITEM-ATTR] @{a.Name.LocalName}=\"{a.Value}\"  int-ok={int.TryParse(a.Value, out _)}");
            }

            foreach (var attr in elem.Attributes())
            {
                var localAttr = attr.Name.LocalName;

                if (intAttrs != null && intAttrs.Contains(localAttr) && !int.TryParse(attr.Value, out _))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[INT-PARSE-FAIL] {docName} <{localElem}> @{localAttr}=\"{attr.Value}\"");
                    anyIntFail = true;
                }

                // Flag empty-string attributes (GetString("x") returns "" → ParseInt32("") throws)
                if (attr.Value.Length == 0 && localAttr != "MetaInfo" && localAttr != "FileLocation")
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[EMPTY-ATTR] {docName} <{localElem}> @{localAttr}=\"\"");
                    anyEmpty = true;
                }
            }
        }

        if (!anyIntFail)
            System.Diagnostics.Debug.WriteLine($"[INT-SCAN-OK] {docName}: all known-int attributes are valid");
        if (!anyEmpty)
            System.Diagnostics.Debug.WriteLine($"[EMPTY-SCAN-OK] {docName}: no unexpected empty attributes");
    }

    // ── Encryption ────────────────────────────────────────────────────────────

    private static async Task<(byte[] Encrypted, long Size)> EncryptStreamAsync(Stream plaintext, byte[] key)
    {
        using var ms = new MemoryStream();
        await plaintext.CopyToAsync(ms);
        var bytes = ms.ToArray();
        return (AesEncrypt(bytes, key), bytes.Length);
    }

    private byte[] EncryptXml(XDocument doc)
    {
        using var ms = new MemoryStream();
        var settings = new System.Xml.XmlWriterSettings
        {
            Encoding        = new System.Text.UTF8Encoding(false),
            Indent          = false,
            OmitXmlDeclaration = false,
        };
        using (var writer = System.Xml.XmlWriter.Create(ms, settings))
            doc.Save(writer);
        var bytes = ms.ToArray();
        System.Diagnostics.Debug.WriteLine($"[XML-serialized] {System.Text.Encoding.UTF8.GetString(bytes)}");
        System.Diagnostics.Debug.WriteLine($"[XML-bytes] first4={BitConverter.ToString(bytes, 0, Math.Min(4, bytes.Length))} len={bytes.Length}");
        return AesEncrypt(bytes, _encryptionKey);
    }

    // AES-256-CBC with a random per-blob IV. Output format: [16-byte IV][ciphertext].
    // SP reads the first 16 bytes as the IV when decrypting.
    private static byte[] AesEncrypt(byte[] plaintext, byte[] key)
    {
        var iv = System.Security.Cryptography.RandomNumberGenerator.GetBytes(16);

        using var aes = Aes.Create();
        aes.Key     = key;
        aes.IV      = iv;
        aes.Mode    = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;

        using var output = new MemoryStream();
        output.Write(iv, 0, iv.Length);
        using (var encryptor = aes.CreateEncryptor())
        using (var cs = new CryptoStream(output, encryptor, CryptoStreamMode.Write))
            cs.Write(plaintext, 0, plaintext.Length);
        return output.ToArray();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    // UIVersion = MajorVersion * 512 + MinorVersion. Graph returns labels like "1", "2", "1.0", "2.0".
    private static (int Major, int Minor) ParseVersionLabel(string versionLabel)
    {
        var parts = versionLabel.Split('.');
        var major = int.TryParse(parts[0], out var maj) ? maj : 1;
        var minor = parts.Length > 1 && int.TryParse(parts[1], out var min) ? min : 0;
        return (major, minor);
    }

    private static string UiVersion(string versionLabel)
    {
        var (major, minor) = ParseVersionLabel(versionLabel);
        return (major * 512 + minor).ToString();
    }

    private static string MajorVersion(string versionLabel) => ParseVersionLabel(versionLabel).Major.ToString();
    private static string MinorVersion(string versionLabel) => ParseVersionLabel(versionLabel).Minor.ToString();

    private static string FormatDate(DateTimeOffset dt) =>
        dt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");

    private static string Claims(string? email) =>
        string.IsNullOrEmpty(email) ? "" : $"i:0#.f|membership|{email}";

    private static XElement Field(string name, string value) =>
        new XElement("Field",
            new XAttribute("Name", name),
            new XAttribute("Value", value));
}
