using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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

    private void ProgressList_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        Dispatcher.BeginInvoke(() =>
        {
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
