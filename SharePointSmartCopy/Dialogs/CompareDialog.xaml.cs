using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using CommunityToolkit.Mvvm.ComponentModel;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.Dialogs;

public partial class CompareDialog : Window
{
    private readonly SharePointService _spService;
    private readonly VerificationReportService _verificationReportService;
    private CancellationTokenSource? _compareCts;
    private SharePointNode? _selectedSourceNode;
    private SharePointNode? _selectedTargetNode;

    public CompareDialog(SharePointService spService)
    {
        _spService = spService;
        InitializeComponent();

        _verificationReportService = new VerificationReportService(spService);
        var settings = AppSettings.Load();
        DataContext = new CompareViewModel
        {
            SourceUrl = settings.SourceUrl,
            TargetUrl = settings.TargetUrl,
        };
        // Seeds the checkbox from the user's last choice; the live IsChecked value (read in
        // CompareButton_Click, not this settings snapshot) is what actually gates a given run.
        DeepVerifyCheckBox.IsChecked = settings.DeepVerifyOfficeFiles;
    }

    private CompareViewModel VM => (CompareViewModel)DataContext;

    // Persisted immediately, same as HistoryDialog's checkbox — this dialog has no other commit
    // point (no "Next" step) to save on.
    private void DeepVerifyCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        var settings = AppSettings.Load();
        settings.DeepVerifyOfficeFiles = DeepVerifyCheckBox.IsChecked == true;
        settings.Save();
    }

    // Closing the window must stop an in-flight compare — without this the scan of a potentially
    // large pair kept running headless after the dialog was gone (same reasoning as HistoryDialog).
    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        _compareCts?.Cancel();
        base.OnClosing(e);
    }

    private async void SourceConnect_Click(object sender, RoutedEventArgs e)
    {
        var url = VM.SourceUrl.Trim();
        if (string.IsNullOrWhiteSpace(url)) return;

        _selectedSourceNode = null;
        UpdateCompareEnabled();
        VM.SourceStatus = "Connecting…";
        VM.IsConnectingSource = true;
        try
        {
            VM.SourceSiteId = await _spService.GetSiteIdAsync(url);
            var libs = await _spService.GetLibrariesAsync(VM.SourceSiteId, url);
            VM.SourceLibraries.Clear();
            foreach (var lib in libs) VM.SourceLibraries.Add(lib);
            VM.SourceStatus = "✅ Connected";
        }
        catch (Exception ex)
        {
            VM.SourceStatus = $"❌ {ex.Message}";
        }
        finally
        {
            VM.IsConnectingSource = false;
        }
    }

    private async void TargetConnect_Click(object sender, RoutedEventArgs e)
    {
        var url = VM.TargetUrl.Trim();
        if (string.IsNullOrWhiteSpace(url)) return;

        _selectedTargetNode = null;
        UpdateCompareEnabled();
        VM.TargetStatus = "Connecting…";
        VM.IsConnectingTarget = true;
        try
        {
            VM.TargetSiteId = await _spService.GetSiteIdAsync(url);
            var libs = await _spService.GetLibrariesAsync(VM.TargetSiteId, url);
            VM.TargetLibraries.Clear();
            foreach (var lib in libs) VM.TargetLibraries.Add(lib);
            VM.TargetStatus = "✅ Connected";
        }
        catch (Exception ex)
        {
            VM.TargetStatus = $"❌ {ex.Message}";
        }
        finally
        {
            VM.IsConnectingTarget = false;
        }
    }

    // Both trees are folders-only (see LoadChildrenAsync) — a compare root is always a library or
    // folder, never an individual file, matching TargetNodeTemplate's existing folder-picker semantics.
    private async void SourceTreeItem_Expanded(object sender, RoutedEventArgs e)
    {
        if (sender is TreeViewItem tvi && tvi.DataContext is SharePointNode node)
        { e.Handled = true; await LoadChildrenAsync(node); }
    }

    private async void TargetTreeItem_Expanded(object sender, RoutedEventArgs e)
    {
        if (sender is TreeViewItem tvi && tvi.DataContext is SharePointNode node)
        { e.Handled = true; await LoadChildrenAsync(node); }
    }

    private async Task LoadChildrenAsync(SharePointNode node)
    {
        if (!node.HasChildren) return;
        if (!node.Children.Any(c => c.IsPlaceholder)) return;

        node.IsLoading = true;
        try
        {
            var children = await _spService.GetChildrenAsync(node.DriveId, node.Id, node.SiteId, node.SiteUrl, foldersOnly: true);
            node.Children.Clear();
            foreach (var child in children)
            {
                child.Parent = node;
                node.Children.Add(child);
            }
        }
        catch { node.Children.Clear(); }
        finally { node.IsLoading = false; }
    }

    private void SourceTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        _selectedSourceNode = e.NewValue as SharePointNode;
        UpdateCompareEnabled();
    }

    private void TargetTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        _selectedTargetNode = e.NewValue as SharePointNode;
        UpdateCompareEnabled();
    }

    private void UpdateCompareEnabled()
        => CompareButton.IsEnabled = _selectedSourceNode != null && _selectedTargetNode != null;

    private async void CompareButton_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedSourceNode == null || _selectedTargetNode == null) return;

        var settings = AppSettings.Load();
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Filter   = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = $"{SiteUrlHelper.ReportFilenamePrefix(VM.SourceUrl, VM.TargetUrl, settings.PrefixReportFilenamesWithSiteNames)}{(DeepVerifyCheckBox.IsChecked == true ? "Deep" : "")}CompareReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
        };
        if (dlg.ShowDialog() != true) return;

        _compareCts?.Dispose();
        _compareCts = new CancellationTokenSource();

        SourceUrlBox.IsEnabled        = false;
        TargetUrlBox.IsEnabled        = false;
        SourceConnectButton.IsEnabled = false;
        TargetConnectButton.IsEnabled = false;
        SourceTreeView.IsEnabled      = false;
        TargetTreeView.IsEnabled      = false;
        CompareButton.IsEnabled       = false;
        DeepVerifyCheckBox.IsEnabled  = false;
        CompareCancelButton.Visibility = Visibility.Visible;
        CompareStatus.Visibility       = Visibility.Visible;
        CompareStatus.Text             = "Scanning…";

        // Live checkbox state, not a re-read of AppSettings — see HistoryDialog's identical field
        // comment for why this must come from the live UI at run time, not a settings snapshot.
        bool deepVerify   = DeepVerifyCheckBox.IsChecked == true;
        string compareName = deepVerify ? "Deep comparison" : "Comparison";

        try
        {
            // IsFolder=true, IsLibrary=true, TargetSubFolderPath="": forces the scan to start
            // exactly at the two picked nodes with no path-wrapping by name — VerificationReportService
            // otherwise assumes the target is "the destination container a copy landed in" and nests
            // an extra SourceName segment under it (see its own Merge/target-root comments). Here the
            // user picked both sides directly, so no such nesting should happen.
            var root = new VerificationRoot(
                _selectedSourceNode.DriveId, _selectedSourceNode.Id, _selectedSourceNode.Name,
                IsFolder: true, IsLibrary: true,
                _selectedTargetNode.DriveId, _selectedTargetNode.Id, TargetSubFolderPath: "");
            var roots = new List<VerificationRoot> { root };

            // Combine the persistent phase-status line with the most recent throttle/error notice —
            // identical pattern to HistoryDialog's VerifyButton_Click.
            string baseText = "Scanning…";
            string noticeText = "";
            void UpdateStatus() =>
                CompareStatus.Text = string.IsNullOrEmpty(noticeText) ? baseText : $"{baseText}  {noticeText}";
            var onScanned = new Progress<VerificationReportService.ScanProgress>(p =>
            {
                baseText = $"Scanning… found {p.SourceFilesFound:N0} source file(s), {p.TargetFilesFound:N0} target file(s)";
                UpdateStatus();
            });
            var onNotice = new Progress<string>(msg =>
            {
                if (msg.StartsWith('⚠'))
                {
                    noticeText = msg;
                    UpdateStatus();
                    var shown = msg;
                    _ = Task.Delay(TimeSpan.FromSeconds(10)).ContinueWith(_ =>
                        Dispatcher.BeginInvoke(() =>
                        {
                            if (noticeText == shown) { noticeText = ""; UpdateStatus(); }
                        }));
                }
                else
                {
                    baseText = msg;
                    UpdateStatus();
                }
            });

            var result = await _verificationReportService.RunAsync(
                roots, settings.MaxParallelCopies, activityLog: onNotice, progress: onScanned, _compareCts.Token, deepVerify);
            CompareStatus.Text = "Writing workbook…";
            await Task.Run(() => ExcelReportWriter.Write(dlg.FileName, result));
            if (result.ScanErrors.Count > 0)
                MessageBox.Show(
                    $"{result.ScanErrors.Count} root(s) could not be scanned — see the Scan Errors tab in the workbook.",
                    compareName, MessageBoxButton.OK, MessageBoxImage.Warning);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
        }
        catch (OperationCanceledException) { /* user cancelled — no message needed */ }
        catch (Exception ex)
        {
            MessageBox.Show($"{compareName} failed: {ex.Message}", compareName, MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SourceUrlBox.IsEnabled        = true;
            TargetUrlBox.IsEnabled        = true;
            SourceConnectButton.IsEnabled = true;
            TargetConnectButton.IsEnabled = true;
            SourceTreeView.IsEnabled      = true;
            TargetTreeView.IsEnabled      = true;
            UpdateCompareEnabled();
            DeepVerifyCheckBox.IsEnabled  = true;
            CompareCancelButton.Visibility = Visibility.Collapsed;
            CompareStatus.Visibility       = Visibility.Collapsed;
        }
    }

    private void CompareCancelButton_Click(object sender, RoutedEventArgs e) => _compareCts?.Cancel();

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}

// Dialog-local view model — a deliberately trimmed-down slice of MainViewModel's connect/browse
// state (no CopyScope, CurrentStep, versions, overwrite mode, or permissions concepts at all).
public partial class CompareViewModel : ObservableObject
{
    [ObservableProperty] private string _sourceUrl = string.Empty;
    [ObservableProperty] private string _sourceStatus = string.Empty;
    [ObservableProperty] private string _sourceSiteId = string.Empty;
    [ObservableProperty] private ObservableCollection<SharePointNode> _sourceLibraries = [];
    [ObservableProperty] private bool _isConnectingSource;

    [ObservableProperty] private string _targetUrl = string.Empty;
    [ObservableProperty] private string _targetStatus = string.Empty;
    [ObservableProperty] private string _targetSiteId = string.Empty;
    [ObservableProperty] private ObservableCollection<SharePointNode> _targetLibraries = [];
    [ObservableProperty] private bool _isConnectingTarget;
}
