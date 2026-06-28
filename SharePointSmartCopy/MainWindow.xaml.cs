using System.Collections.Specialized;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using SharePointSmartCopy.Dialogs;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.ViewModels;

namespace SharePointSmartCopy;

public partial class MainWindow : Window
{
    private MainViewModel VM => (MainViewModel)DataContext;

    public MainWindow() : this(null) { }

    internal MainWindow(MainViewModel? vm)
    {
        InitializeComponent();
        DataContext = vm ?? new MainViewModel();
        ((System.Collections.ObjectModel.ObservableCollection<CopyResult>)VM.CopyResults)
            .CollectionChanged += ProgressList_CollectionChanged;
    }

    // ── Window caption buttons (custom chrome) ────────────────────────────────

    private void Minimize_Click(object sender, RoutedEventArgs e)
        => WindowState = WindowState.Minimized;

    private void MaximizeRestore_Click(object sender, RoutedEventArgs e)
        => WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;

    // A borderless (WindowStyle=None) window maximizes to the full monitor bounds, overflowing the
    // work area so its bottom edge — the navigation footer with the View Report / Next button — is
    // hidden under the taskbar/off-screen. Constrain the maximized size to the monitor work area.
    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        var handle = new WindowInteropHelper(this).Handle;
        HwndSource.FromHwnd(handle)?.AddHook(WindowProc);
    }

    private IntPtr WindowProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        const int WM_GETMINMAXINFO = 0x0024;
        if (msg == WM_GETMINMAXINFO)
        {
            var mmi = Marshal.PtrToStructure<MINMAXINFO>(lParam);
            const int MONITOR_DEFAULTTONEAREST = 0x00000002;
            var monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
            if (monitor != IntPtr.Zero)
            {
                var info = new MONITORINFO { cbSize = Marshal.SizeOf<MONITORINFO>() };
                GetMonitorInfo(monitor, ref info);
                var work = info.rcWork;
                var screen = info.rcMonitor;
                // Position/size relative to the monitor so the maximized window exactly fills the
                // work area (excludes the taskbar) instead of the whole screen.
                mmi.ptMaxPosition.X = work.Left - screen.Left;
                mmi.ptMaxPosition.Y = work.Top - screen.Top;
                mmi.ptMaxSize.X     = work.Right - work.Left;
                mmi.ptMaxSize.Y     = work.Bottom - work.Top;
                // Preserve the window's own minimum-size constraints while maximized.
                mmi.ptMinTrackSize.X = (int)MinWidth;
                mmi.ptMinTrackSize.Y = (int)MinHeight;
                Marshal.StructureToPtr(mmi, lParam, true);
            }
            handled = true;
        }
        return IntPtr.Zero;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, int flags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left, Top, Right, Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MINMAXINFO
    {
        public POINT ptReserved;
        public POINT ptMaxSize;
        public POINT ptMaxPosition;
        public POINT ptMinTrackSize;
        public POINT ptMaxTrackSize;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MONITORINFO
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public int dwFlags;
    }

    private void CloseWindow_Click(object sender, RoutedEventArgs e)
        => Close();

    private void HistoryButton_Click(object sender, RoutedEventArgs e)
        => new HistoryDialog { Owner = this }.ShowDialog();

    private void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SettingsDialog(VM.Settings, VM.AuthService) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            VM.ApplySettings(dlg.Result);
            ShowToast("Settings saved. Connect again if you were already signed in.");
        }
    }

    // Shows a transient bottom-center notification for ~3 seconds.
    private System.Windows.Threading.DispatcherTimer? _toastTimer;
    private void ShowToast(string message)
    {
        ToastText.Text = message;
        ToastHost.Visibility = Visibility.Visible;
        _toastTimer?.Stop();
        _toastTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(3)
        };
        _toastTimer.Tick += (_, _) =>
        {
            _toastTimer.Stop();
            ToastHost.Visibility = Visibility.Collapsed;
        };
        _toastTimer.Start();
    }

    private async void SourceTreeItem_Expanded(object sender, RoutedEventArgs e)
    {
        if (sender is TreeViewItem tvi && tvi.DataContext is SharePointNode node)
        { e.Handled = true; await VM.LoadNodeChildrenAsync(node); }
    }

    private void SourceNode_CheckChanged(object sender, RoutedEventArgs e)
        => VM.NotifySelectionChanged();

    // Cycles custom list nodes: blank → checked (structure+items) → blank (items only) → blank (nothing).
    // The null state is visually identical to unchecked via the ItemsOnlyCheckBox style.
    // All other nodes are a plain two-state toggle — without this, WPF's native tri-state
    // cycle sends "checked → indeterminate" on uncheck, leaving stray null states the app
    // would misread as items-only mode (phantom entries in the Copy Preview).
    private void SourceCheckBox_PreviewClick(object sender, MouseButtonEventArgs e)
    {
        if ((sender as CheckBox)?.DataContext is not SharePointNode node)
            return;
        node.IsChecked = node.IsCustomList
            ? node.IsChecked switch
              {
                  false => true,   // blank → checked
                  true  => null,   // checked → blank (items only, no structure)
                  _     => false   // null/blank → blank (deselect all)
              }
            : node.IsChecked != true;
        e.Handled = true;
        VM.NotifySelectionChanged();
    }

    private void SelectAll_Click(object sender, RoutedEventArgs e)
    { VM.SelectAllSource(true); VM.NotifySelectionChanged(); }

    private void DeselectAll_Click(object sender, RoutedEventArgs e)
    { VM.SelectAllSource(false); VM.NotifySelectionChanged(); }

    private async void TargetTreeItem_Expanded(object sender, RoutedEventArgs e)
    {
        if (sender is TreeViewItem tvi && tvi.DataContext is SharePointNode node)
        { e.Handled = true; await VM.LoadTargetNodeChildrenAsync(node); }
    }

    private void TargetTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is SharePointNode node) VM.SelectedTargetFolder = node;
    }

    private void InfoIcon_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.ToolTip is ToolTip tt)
        {
            if (tt.IsOpen)
            {
                tt.StaysOpen = false;
                tt.IsOpen = false;
            }
            else
            {
                tt.PlacementTarget = fe;
                tt.StaysOpen = true;
                tt.IsOpen = true;
            }
            e.Handled = true;
        }
    }

    private bool _autoScrollQueued;
    private void ProgressList_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        // Coalesce auto-scroll: scrolling to the bottom on EVERY add forces a layout pass per item,
        // which saturates the UI thread on huge copies (tens of thousands of rows). Queue at most one
        // pending scroll at Background priority; rapid bursts of adds collapse to a single scroll once
        // the UI thread goes idle.
        if (_autoScrollQueued) return;
        _autoScrollQueued = true;
        Dispatcher.BeginInvoke(() =>
        {
            _autoScrollQueued = false;
            if (ProgressList.Items.Count > 0)
                ProgressList.ScrollIntoView(ProgressList.Items[^1]);
        }, System.Windows.Threading.DispatcherPriority.Background);
    }

    private void StartNewCopy_Click(object sender, RoutedEventArgs e)
    {
        DataContext = new MainViewModel(VM.AuthService);
        ((System.Collections.ObjectModel.ObservableCollection<CopyResult>)VM.CopyResults)
            .CollectionChanged += ProgressList_CollectionChanged;
    }

    // ── Mode tile click handlers ───────────────────────────────────────────────

    private void ModeFiles_Click(object sender, RoutedEventArgs e)
    {
        VM.ColumnMappings.Clear();
        VM.CopyScope = CopyScope.Files;
        _ = VM.LoadLibrariesAsync();
    }

    private void ModeLibrary_Click(object sender, RoutedEventArgs e)
    {
        VM.ColumnMappings.Clear();
        VM.CopyScope = CopyScope.Library;
        _ = VM.LoadLibrariesAsync();
    }

    private void ModeSite_Click(object sender, RoutedEventArgs e)
    {
        VM.ColumnMappings.Clear();
        VM.CopyScope = CopyScope.Site;
        _ = VM.LoadLibrariesAsync();
    }

    private void ModePages_Click(object sender, RoutedEventArgs e)
    {
        VM.ColumnMappings.Clear();
        VM.CopyScope = CopyScope.Pages;
        _ = VM.LoadPageLibraryAsync();
    }

    // ── Step 4 Copy log column sorting ───────────────────────────────────────────

    static readonly Dictionary<string, string> _headerSortMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Item"]                = "FileName",
        ["Status"]              = "Status",
        ["Source"]              = "SourcePath",
        ["Target"]              = "TargetPath",
        ["Versions"]            = "VersionsCopied",
        ["Details"]             = "ErrorMessage",
        ["Permissions Status"]  = "PermissionStatus",
        ["Permissions Details"] = "PermissionDetails",
    };

    private GridViewColumnHeader? _sortHeader;
    private SortAdorner?          _sortAdorner;

    private void ProgressList_ColumnHeaderClick(object sender, RoutedEventArgs e)
    {
        if (e.OriginalSource is not GridViewColumnHeader header ||
            header.Role == GridViewColumnHeaderRole.Padding ||
            header.Content is not string headerText ||
            !_headerSortMap.TryGetValue(headerText, out var sortBy))
            return;

        var dir = (_sortHeader == header && _sortAdorner?.Direction == ListSortDirection.Ascending)
            ? ListSortDirection.Descending
            : ListSortDirection.Ascending;

        if (_sortAdorner != null && _sortHeader != null)
            AdornerLayer.GetAdornerLayer(_sortHeader)?.Remove(_sortAdorner);

        _sortAdorner = new SortAdorner(header, dir);
        AdornerLayer.GetAdornerLayer(header)?.Add(_sortAdorner);
        _sortHeader  = header;

        VM.CopyResultsView.SortDescriptions.Clear();
        VM.CopyResultsView.SortDescriptions.Add(new SortDescription(sortBy, dir));
    }

    sealed class SortAdorner(UIElement element, ListSortDirection direction) : Adorner(element)
    {
        static readonly Geometry AscGeometry  = Geometry.Parse("M 0 4 L 3.5 0 L 7 4 Z");
        static readonly Geometry DescGeometry = Geometry.Parse("M 0 0 L 3.5 4 L 7 0 Z");

        public ListSortDirection Direction { get; } = direction;

        protected override void OnRender(DrawingContext dc)
        {
            base.OnRender(dc);
            if (AdornedElement.RenderSize.Width < 20) return;
            var brush = Application.Current.TryFindResource("AccentBrush") as Brush ?? Brushes.SteelBlue;
            dc.PushTransform(new TranslateTransform(
                AdornedElement.RenderSize.Width - 15,
                (AdornedElement.RenderSize.Height - 5) / 2));
            dc.DrawGeometry(brush, null,
                Direction == ListSortDirection.Ascending ? AscGeometry : DescGeometry);
            dc.Pop();
        }
    }

    private void CopyCustomColumns_Changed(object sender, RoutedEventArgs e)
    {
        VM.OnPropertyChanged(nameof(VM.MappingButtonLabel));
    }

    private async void ConfigureMappings_Click(object sender, RoutedEventArgs e)
    {
        // Columns are loaded in the background when step 2 → 3 advances.
        // Wait here so the dialog always has data, even if the user clicks immediately.
        if (VM._columnLoadTask != null)
            await VM._columnLoadTask;
        var dlg = new ColumnMappingDialog(VM) { Owner = this };
        dlg.ShowDialog();
        VM.OnPropertyChanged(nameof(VM.MappingButtonLabel));
    }
}
