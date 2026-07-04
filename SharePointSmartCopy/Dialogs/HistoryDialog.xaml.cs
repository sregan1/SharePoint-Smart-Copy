using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.Dialogs;

public partial class HistoryDialog : Window
{
    private List<SavedReportSummary> _reports = [];
    private readonly VerificationReportService _verificationReportService;
    private CancellationTokenSource? _verifyCts;

    public HistoryDialog(SharePointService spService)
    {
        InitializeComponent();
        _verificationReportService = new VerificationReportService(spService);
        DetailHeader.Text = "Loading previous runs…";
        _ = LoadReportsAsync();
    }

    // Off the UI thread: on a tenant with several very large (100,000+-file) runs in its history,
    // reading and parsing all 50 saved report files synchronously in the constructor is what made
    // the window itself pause before appearing. LoadSummaries (not LoadAll) additionally skips
    // materializing each report's Items array — see SavedReportSummary — so this is fast even before
    // accounting for the background thread.
    private async Task LoadReportsAsync()
    {
        _reports = await Task.Run(ReportHistoryService.LoadSummaries);
        ReportList.ItemsSource = null;
        ReportList.ItemsSource = _reports;

        ReportListLoading.Visibility = Visibility.Collapsed;
        ReportList.Visibility        = Visibility.Visible;
        DetailHeader.Text = _reports.Count == 0 ? "No previous runs found." : "Select a run to view details";
    }

    private void ReportList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        var report     = ReportList.SelectedItem as SavedReportSummary;
        bool hasSelection = report != null;

        DeleteButton.IsEnabled  = hasSelection;
        ExportButton.IsEnabled  = hasSelection;
        SummaryCards.Visibility = hasSelection ? Visibility.Visible : Visibility.Collapsed;
        EmptyHint.Visibility    = hasSelection ? Visibility.Collapsed : Visibility.Visible;

        bool canVerify = hasSelection && report!.Roots.Count > 0;
        VerifyButton.IsEnabled = canVerify;
        VerifyButton.ToolTip = hasSelection && !canVerify
            ? "This run predates verification support — re-run the copy to enable this."
            : null;

        if (report == null)
        {
            DetailHeader.Text      = "Select a run to view details";
            DetailGrid.ItemsSource = null;
            return;
        }

        DetailHeader.Text = $"{report.DisplayDate}  —  {report.TotalCount} files  —  {report.DurationDisplay}";
        SuccessCard.Text  = report.SuccessCount.ToString();
        FailedCard.Text   = report.FailedCount.ToString();
        SkippedCard.Text  = report.SkippedCount.ToString();

        // Items lives only in the full report on disk (see SavedReportSummary) — load it now that
        // this specific run's detail is actually being viewed, not for every row up front.
        DetailGrid.ItemsSource = ReportHistoryService.LoadFull(report.Id)?.Items ?? [];
    }

    private void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (ReportList.SelectedItem is not SavedReportSummary report) return;

        var result = MessageBox.Show(
            $"Delete the report from {report.DisplayDate}?",
            "Delete Report", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes) return;

        ReportHistoryService.Delete(report.Id);
        _ = LoadReportsAsync();
    }

    private void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        if (ReportList.SelectedItem is not SavedReportSummary summary) return;
        var report = ReportHistoryService.LoadFull(summary.Id);
        if (report == null) return;

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Filter   = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt",
            FileName = $"CopyReport_{report.Id}.csv"
        };
        if (dlg.ShowDialog() != true) return;

        static string Csv(string? s) => $"\"{(s ?? "").Replace("\"", "\"\"")}\"";
        var sb = new StringBuilder();
        sb.AppendLine("File Name,Source Path,Target Path,Status,Versions Copied,Error,Permissions Status,Permissions Details");
        foreach (var item in report.Items)
        {
            sb.AppendLine(
                $"{Csv(item.FileName)}," +
                $"{Csv(item.SourcePath)}," +
                $"{Csv(item.TargetPath)}," +
                $"{item.Status}," +
                $"{item.VersionsCopied}," +
                $"{Csv(item.ErrorMessage)}," +
                $"{item.PermissionStatus?.ToString() ?? ""}," +
                $"{Csv(item.PermissionDetails)}");
        }
        System.IO.File.WriteAllText(dlg.FileName, sb.ToString());
        System.Diagnostics.Process.Start(
            new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
    }

    private async void VerifyButton_Click(object sender, RoutedEventArgs e)
    {
        if (ReportList.SelectedItem is not SavedReportSummary report || report.Roots.Count == 0) return;

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Filter   = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = $"VerificationReport_{report.Id}.xlsx"
        };
        if (dlg.ShowDialog() != true) return;

        _verifyCts?.Dispose();
        _verifyCts = new CancellationTokenSource();

        ReportList.IsEnabled   = false;
        DeleteButton.IsEnabled = false;
        ExportButton.IsEnabled = false;
        VerifyButton.IsEnabled = false;
        VerifyCancelButton.Visibility = Visibility.Visible;
        VerifyStatus.Visibility = Visibility.Visible;
        VerifyStatus.Text = "Scanning…";

        try
        {
            var roots       = VerificationRoot.FromSavedReport(report.Roots);
            var maxParallel = AppSettings.Load().MaxParallelCopies;

            // Combine the scan-progress line with the most recent throttle/error notice (if any) so
            // a long Retry-After wait — previously invisible here — shows up instead of just leaving
            // the file counts frozen, which looked indistinguishable from a hang.
            string scanText = "Scanning…";
            string noticeText = "";
            void UpdateStatus() =>
                VerifyStatus.Text = string.IsNullOrEmpty(noticeText) ? scanText : $"{scanText}  {noticeText}";
            var onScanned = new Progress<VerificationReportService.ScanProgress>(p =>
            {
                scanText = $"Scanning… found {p.SourceFilesFound:N0} source file(s), {p.TargetFilesFound:N0} target file(s)";
                UpdateStatus();
            });
            var onNotice = new Progress<string>(msg =>
            {
                noticeText = msg;
                UpdateStatus();
            });
            var result = await _verificationReportService.RunAsync(
                roots, maxParallel, activityLog: onNotice, progress: onScanned, _verifyCts.Token);
            ExcelReportWriter.Write(dlg.FileName, result);
            if (result.ScanErrors.Count > 0)
                MessageBox.Show(
                    $"{result.ScanErrors.Count} root(s) could not be scanned — see the Scan Errors tab in the workbook.",
                    "Verification", MessageBoxButton.OK, MessageBoxImage.Warning);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
        }
        catch (OperationCanceledException) { /* user cancelled — no message needed */ }
        catch (Exception ex)
        {
            MessageBox.Show($"Verification failed: {ex.Message}", "Verification", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            ReportList.IsEnabled   = true;
            DeleteButton.IsEnabled = ReportList.SelectedItem != null;
            ExportButton.IsEnabled = ReportList.SelectedItem != null;
            VerifyButton.IsEnabled = ReportList.SelectedItem is SavedReportSummary r && r.Roots.Count > 0;
            VerifyCancelButton.Visibility = Visibility.Collapsed;
            VerifyStatus.Visibility = Visibility.Collapsed;
        }
    }

    private void VerifyCancelButton_Click(object sender, RoutedEventArgs e) => _verifyCts?.Cancel();

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
