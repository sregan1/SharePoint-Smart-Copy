using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Services;

// Lets the user copy text from anywhere in the app. Grid cells and labels (TextBlocks) are not selectable in WPF, so
// right-clicking one opens a small menu with Copy. One handler on every window, so no dialog needs its own markup.
//
// Left alone, so they keep their own behavior: text boxes (they have a Copy menu and real selection), buttons and other
// controls, combo boxes, scroll bars, and anything that sits under an element with a context menu of its own.
public static class TextCopySupport
{
    private static bool _registered;

    public static void Register()
    {
        if (_registered) return;
        _registered = true;
        EventManager.RegisterClassHandler(typeof(Window), UIElement.PreviewMouseRightButtonUpEvent, new MouseButtonEventHandler(OnRightClick));
    }

    private static void OnRightClick(object sender, MouseButtonEventArgs e)
    {
        if (e.Handled || e.OriginalSource is not DependencyObject source) return;
        var text = FindText(source);
        if (string.IsNullOrWhiteSpace(text)) return;

        var copy = new MenuItem { Header = Loc.T("Ui_Copy") };
        copy.Click += (_, _) => SetClipboard(text);
        new ContextMenu { Items = { copy }, Placement = PlacementMode.MousePoint, IsOpen = true };
        e.Handled = true;
    }

    // The text under the click: the whole cell for a grid cell, otherwise the label that was hit. Null when the click was on
    // something this should not take over.
    private static string? FindText(DependencyObject start)
    {
        string? label = null;
        for (var d = start; d is not null; d = ParentOf(d))
        {
            if (d is TextBoxBase or PasswordBox or ComboBox or ComboBoxItem or ButtonBase or MenuItem or ScrollBar or Slider or ProgressBar)
                return null;
            if (d is FrameworkElement { ContextMenu: not null }) return null;
            if (d is DataGridCell cell) return CellText(cell);
            if (d is TextBlock tb && label is null) label = TextOf(tb);
        }
        return label;
    }

    private static DependencyObject? ParentOf(DependencyObject d) =>
        d is Visual or System.Windows.Media.Media3D.Visual3D ? VisualTreeHelper.GetParent(d) ?? LogicalTreeHelper.GetParent(d) : LogicalTreeHelper.GetParent(d);

    private static string TextOf(TextBlock tb) =>
        tb.Text.Length > 0 ? tb.Text : string.Concat(tb.Inlines.OfType<Run>().Select(r => r.Text));

    private static string? CellText(DataGridCell cell)
    {
        if (cell.Content is TextBlock tb) return TextOf(tb);
        if (cell.Content is TextBox box) return box.Text;
        // a templated cell (chips, icons plus text): the texts in it, in order
        var parts = new List<string>();
        Collect(cell, parts);
        return parts.Count > 0 ? string.Join(" ", parts) : cell.Content?.ToString();
    }

    private static void Collect(DependencyObject root, List<string> parts)
    {
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(root); i++)
        {
            var child = VisualTreeHelper.GetChild(root, i);
            if (child is TextBlock t && TextOf(t) is { Length: > 0 } s) parts.Add(s);
            else Collect(child, parts);
        }
    }

    // The clipboard can be held open by another program for a moment; a copy that silently does nothing is worse than a retry.
    private static void SetClipboard(string text)
    {
        for (var attempt = 0; attempt < 5; attempt++)
        {
            try { Clipboard.SetDataObject(text, copy: true); return; }
            catch (System.Runtime.InteropServices.COMException) { Thread.Sleep(50); }
        }
    }
}
