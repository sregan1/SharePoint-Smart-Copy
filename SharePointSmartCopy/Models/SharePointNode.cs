using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

namespace SharePointSmartCopy.Models;

public enum NodeType { Library, Folder, File, ListItem }

public partial class SharePointNode : ObservableObject
{
    [ObservableProperty] private string _name = string.Empty;
    [ObservableProperty] private string _id = string.Empty;
    [ObservableProperty] private string _driveId = string.Empty;
    [ObservableProperty] private string _siteId = string.Empty;
    [ObservableProperty] private string _siteUrl = string.Empty;
    [ObservableProperty] private NodeType _type;
    [ObservableProperty] private long? _size;
    [ObservableProperty] private string? _webUrl;
    [ObservableProperty] private bool _isLoading;
    [ObservableProperty] private bool _isExpanded;
    [ObservableProperty] private bool? _isChecked = false;
    [ObservableProperty] private ObservableCollection<SharePointNode> _children = [];

    [ObservableProperty] private string? _overrideName;

    // Item-level copies (a list in items-only/partial mode): where THIS list's selected items go.
    // Chosen per source list on the Target step — either an existing target list or a new one.
    public const string NewDestinationSentinelId = "__new__";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsCreatingNewDestination))]
    [NotifyPropertyChangedFor(nameof(EffectiveDestinationName))]
    [NotifyPropertyChangedFor(nameof(HasValidDestination))]
    private ListPickerItem? _destinationList;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(EffectiveDestinationName))]
    [NotifyPropertyChangedFor(nameof(HasValidDestination))]
    private string _newDestinationName = string.Empty;

    public bool IsCreatingNewDestination => DestinationList?.Id == NewDestinationSentinelId;
    public string EffectiveDestinationName =>
        IsCreatingNewDestination ? NewDestinationName.Trim() : (DestinationList?.Title ?? string.Empty);
    public bool HasValidDestination =>
        DestinationList != null && (!IsCreatingNewDestination || !string.IsNullOrWhiteSpace(NewDestinationName));

    public bool HasChildren { get; set; }
    public bool IsPage { get; set; }
    public bool IsCustomList { get; set; }
    public int ListBaseTemplate { get; set; }
    public string? SourceListId { get; set; }
    public SharePointNode? Parent { get; set; }
    public string? ServerRelativePath { get; set; }

    public bool IsPlaceholder => Name == "__placeholder__";

    public string SizeDisplay => SizeFormatter.FormatBytes(Size);

    // Segoe Fluent Icons / Segoe MDL2 Assets glyph codepoints — rendered with the
    // shared GlyphIcon style (Themes/AppStyles.xaml), not as emoji text.
    public string TypeIcon => IsCustomList ? "" : Type switch
    {
        NodeType.Library  => "", // Library
        NodeType.Folder   => "", // Folder
        NodeType.ListItem => "", // BulletedList
        _                 => GetFileIcon(Name)
    };

    partial void OnIsCheckedChanged(bool? value)
    {
        // Folder null (indeterminate): parent deselected but children retain their state — no cascade.
        // Custom list null (items-only): cascade true to children so their content is included.
        // true: cascade true. false: cascade false.
        if (value == null && !IsCustomList) return;
        var childValue = value == false ? (bool?)false : (bool?)true;
        foreach (var child in Children)
        {
            if (!child.IsPlaceholder)
                child.IsChecked = childValue;
        }
    }

    public IEnumerable<SharePointNode> GetCheckedNodes()
    {
        if (IsChecked == true && !IsPlaceholder)
        {
            yield return this;
            yield break;
        }
        foreach (var child in Children.Where(c => !c.IsPlaceholder))
            foreach (var n in child.GetCheckedNodes())
                yield return n;
    }

    private static string GetFileIcon(string name)
    {
        var ext = System.IO.Path.GetExtension(name).ToLowerInvariant();
        return ext switch
        {
            ".docx" or ".doc"  => "", // Document
            ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" => "", // Picture
            ".mp4" or ".avi" or ".mov" => "", // Video
            ".mp3" or ".wav"   => "", // MusicInfo
            _                  => ""  // Page
        };
    }

}
