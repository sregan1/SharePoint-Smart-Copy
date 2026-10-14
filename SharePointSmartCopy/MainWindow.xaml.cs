using SharePointSmartCopy.Localization;
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
        // Real launches only (demo mode passes its own view model).
        if (vm == null) _ = RecoverInterruptedRunsAsync();

        AttachViewModel();
        Loaded += (_, _) => FocusStepInput();
        // The ⓘ help icons are focusable now (Tab reaches them); show their text when they get focus.
        AddHandler(GotKeyboardFocusEvent, new KeyboardFocusChangedEventHandler(HelpIcon_GotKeyboardFocus), true);
        AddHandler(LostKeyboardFocusEvent, new KeyboardFocusChangedEventHandler(HelpIcon_LostKeyboardFocus), true);
    }

    // ── View model hookup ─────────────────────────────────────────────────────

    private MainViewModel? _attachedVm;

    // Re-run whenever the window gets a NEW view model (Start New Copy replaces it).
    private void AttachViewModel()
    {
        if (_attachedVm != null) _attachedVm.PropertyChanged -= Vm_PropertyChanged;
        _attachedVm = VM;
        // The view model never shows UI itself — this is how it asks "are you sure?".
        VM.ConfirmAction = (title, message) =>
            MessageBox.Show(this, message, title, MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No)
                == MessageBoxResult.Yes;
        VM.PropertyChanged += Vm_PropertyChanged;
    }

    private void Vm_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        switch (e.PropertyName)
        {
            // A long run finished while the user was elsewhere: flash the taskbar button.
            case nameof(MainViewModel.IsReadyForReport) when VM.IsReadyForReport && !IsActive:
                FlashTaskbar();
                break;
            case nameof(MainViewModel.CurrentStep):
                FocusStepInput();
                break;
        }
    }

    // Puts the cursor where the user will type next: the URL box on the Source and Target steps.
    private void FocusStepInput()
    {
        Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input, () =>
        {
            TextBox? box = VM.CurrentStep switch { 0 => SourceUrlBox, 2 => TargetUrlBox, _ => null };
            if (box is { IsVisible: true, IsEnabled: true }) box.Focus();
        });
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct FLASHWINFO
    {
        public uint cbSize;
        public IntPtr hwnd;
        public uint dwFlags;
        public uint uCount;
        public uint dwTimeout;
    }

    [DllImport("user32.dll")]
    private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

    private void FlashTaskbar()
    {
        var handle = new WindowInteropHelper(this).Handle;
        if (handle == IntPtr.Zero) return;
        // FLASHW_ALL | FLASHW_TIMERNOFG: flash caption and taskbar button until the window is brought forward.
        var info = new FLASHWINFO
        {
            cbSize    = (uint)Marshal.SizeOf<FLASHWINFO>(),
            hwnd      = handle,
            dwFlags   = 0x00000003 | 0x0000000C,
            uCount    = uint.MaxValue,
            dwTimeout = 0,
        };
        FlashWindowEx(ref info);
    }

    private static void HelpIcon_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (e.NewFocus is TextBlock { ToolTip: ToolTip tip } icon)
        {
            tip.PlacementTarget = icon;
            tip.IsOpen = true;
        }
    }

    private static void HelpIcon_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (e.OldFocus is TextBlock { ToolTip: ToolTip tip })
            tip.IsOpen = false;
    }

    private void OpenSettings_Click(object sender, RoutedEventArgs e) => SettingsButton_Click(sender, e);

    // Clicking the Failed / Skipped tile on the report filters the list to those items.
    private void FailedTile_Click(object sender, MouseButtonEventArgs e)
        => VM.FileResultFilter = ResultFilterKind.Failed;

    private void SkippedTile_Click(object sender, MouseButtonEventArgs e)
        => VM.FileResultFilter = ResultFilterKind.Skipped;

    // Drag handle under the live activity feed: make it taller or shorter.
    private void ActivityResize_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
    {
        var current = double.IsInfinity(ActivityFeedBorder.MaxHeight)
            ? ActivityFeedBorder.ActualHeight
            : ActivityFeedBorder.MaxHeight;
        ActivityFeedBorder.MaxHeight = Math.Clamp(current + e.VerticalChange, 60, 600);
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

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        // One stray click on X used to end a multi-hour copy with no warning.
        if (!VM.ConfirmCloseDuringRun())
        {
            e.Cancel = true;
            return;
        }
        base.OnClosing(e);
        VM.CancelAllOperations();
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        // Force-kill the process — MSAL/Graph SDK create threads that outlive WPF's
        // Dispatcher shutdown, causing the process to linger in Task Manager.
        Environment.Exit(0);
    }

    private void CloseWindow_Click(object sender, RoutedEventArgs e)
        => Close();

    private async void HistoryButton_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new HistoryDialog(VM.SpService,
            r => VM.PrepareReportAccessAsync(r.SourceUrl, r.TargetUrl, r.Roots),
            (owner, users) => OpenIdentityMapping(owner, users)) { Owner = this };
        dlg.ShowDialog();
        if (dlg.ResumeReport is not { } report) return;

        // async void: nothing may escape here or it would crash the app.
        try
        {
            if (VM.IsCopying) { ShowToast(Loc.T("Dlg_ResumeBusy")); return; }
            await VM.ResumeFromReportAsync(report);
        }
        catch (Exception ex)
        {
            ShowToast(Loc.T("Dlg_ResumeFailed", ex.Message));
        }
    }

    // Turns any run whose journal was never completed (the app closed or crashed mid-copy) into an
    // "Interrupted" History entry — see RunJournal.
    private async Task RecoverInterruptedRunsAsync()
    {
        try
        {
            var recovered = await Task.Run(Services.RunJournal.RecoverInterrupted);
            if (recovered > 0) ShowToast(Loc.T("Dlg_InterruptedRecovered", recovered));
        }
        catch { /* recovery is a convenience; never disturb startup */ }
    }

    private void CompareButton_Click(object sender, RoutedEventArgs e)
        => new CompareDialog(VM.SpService, (url, isSource) => VM.PrepareSiteAccessAsync(url, isSource)) { Owner = this }.ShowDialog();

    private void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SettingsDialog(VM.Settings, VM.AuthService, VM.TargetAuthService) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            VM.ApplySettings(dlg.Result);
            ShowToast(Loc.T("Dlg_SettingsSaved"));
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

    // Cycles folder/library and custom-list nodes through three states:
    //   false → true  (select this node + all children)
    //   true  → null  (deselect this node only; children stay checked)
    //   null  → false (deselect this node + all children)
    // For custom lists: null = items-only (no structure). For folders: null = indeterminate
    // (parent header unchecked, sub-items still checked). Files/list items: plain toggle.
    private void SourceCheckBox_PreviewClick(object sender, MouseButtonEventArgs e)
    {
        if ((sender as CheckBox)?.DataContext is not SharePointNode node)
            return;
        bool isFolder = node.Type is NodeType.Folder or NodeType.Library;
        node.IsChecked = (node.IsCustomList || isFolder)
            ? node.IsChecked switch
              {
                  false => true,   // unchecked → all selected
                  true  => null,   // all selected → parent only deselected
                  _     => false   // indeterminate → all deselected
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

    private void IdentityMapping_Click(object sender, RoutedEventArgs e) => OpenIdentityMapping(this, null);

    // "Map these users…" on the report: the people the last run could not find, ready to be mapped.
    private void MapUnresolved_Click(object sender, RoutedEventArgs e) => OpenIdentityMapping(this, VM.UnresolvedUsers);

    // One entry point for every way into the mapping table (Source step, Target step, report, History).
    // prefill = people to add as rows; owner = whichever window it was opened from.
    private void OpenIdentityMapping(Window owner, IReadOnlyList<UnresolvedUser>? prefill)
    {
        var dlg = new IdentityMappingDialog(VM.Settings, prefill,
            // "Check in target" works once there is a target URL (multi-tenant mode).
            VM.CanCheckTargetUsers ? (email => VM.CheckTargetUserAsync(email)) : null) { Owner = owner };
        if (dlg.ShowDialog() != true) return;

        VM.ApplyIdentityMapping();
        // Coming from a run: say plainly what saving does NOT do. Copy-if-newer skips files that are up to
        // date, so their authors are not re-stamped — only a re-run with Overwrite (or new files) picks the
        // mapping up. A toast is too brief for that.
        if (prefill is { Count: > 0 })
            MessageBox.Show(owner, Loc.T("Dlg_MappingSavedRerun"), Loc.T("Ui_IdentityMapping"),
                MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void StartNewCopy_Click(object sender, RoutedEventArgs e)
    {
        DataContext = new MainViewModel(VM.AuthService, existingTargetAuthService: VM.TargetAuthService);
        ((System.Collections.ObjectModel.ObservableCollection<CopyResult>)VM.CopyResults)
            .CollectionChanged += ProgressList_CollectionChanged;
        AttachViewModel();
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

    // Sort member for each copy-log column, in the column order of the ListView's GridView. Matching on
    // the header TEXT (the old approach) only worked in English — the headers are localized — and it
    // never covered the custom-field columns.
    static readonly string[] _columnSortMembers =
    [
        "FileName", "Status", "SourcePath", "TargetPath", "VersionsCopied", "ErrorMessage",
        "PermissionStatus", "PermissionDetails", "CustomFieldStatus", "CustomFieldDetails",
    ];

    private GridViewColumnHeader? _sortHeader;
    private SortAdorner?          _sortAdorner;

    private void ProgressList_ColumnHeaderClick(object sender, RoutedEventArgs e)
    {
        if (e.OriginalSource is not GridViewColumnHeader header ||
            header.Role == GridViewColumnHeaderRole.Padding ||
            header.Column == null ||
            ProgressList.View is not GridView gridView)
            return;
        int columnIndex = gridView.Columns.IndexOf(header.Column);
        if (columnIndex < 0 || columnIndex >= _columnSortMembers.Length) return;
        var sortBy = _columnSortMembers[columnIndex];

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
