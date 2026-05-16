using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

namespace SharePointSmartCopy.Models;

public enum NodeType { Library, Folder, File }

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
    [ObservableProperty] private bool _isChecked;
    [ObservableProperty] private ObservableCollection<SharePointNode> _children = [];

    public bool HasChildren { get; set; }
    public SharePointNode? Parent { get; set; }
    public string? ServerRelativePath { get; set; }

    public bool IsPlaceholder => Name == "__placeholder__";

    public string SizeDisplay => Size.HasValue ? FormatSize(Size.Value) : string.Empty;

    public string TypeIcon => Type switch
    {
        NodeType.Library => "📚",
        NodeType.Folder  => "📁",
        _                => GetFileIcon(Name)
    };

    partial void OnIsCheckedChanged(bool value)
    {
        foreach (var child in Children)
        {
            if (!child.IsPlaceholder)
                child.IsChecked = value;
        }
    }

    public IEnumerable<SharePointNode> GetCheckedNodes()
    {
        if (IsChecked && !IsPlaceholder)
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
            ".docx" or ".doc"  => "📝",
            ".xlsx" or ".xls"  => "📊",
            ".pptx" or ".ppt"  => "📊",
            ".pdf"             => "📄",
            ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" => "🖼️",
            ".mp4" or ".avi" or ".mov" => "🎬",
            ".mp3" or ".wav"   => "🎵",
            ".zip" or ".rar"   => "🗜️",
            _                  => "📄"
        };
    }

    private static string FormatSize(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024 * 1024) return $"{bytes / 1024.0:N1} KB";
        if (bytes < 1024L * 1024 * 1024) return $"{bytes / (1024.0 * 1024):N1} MB";
        return $"{bytes / (1024.0 * 1024 * 1024):N2} GB";
    }
}
