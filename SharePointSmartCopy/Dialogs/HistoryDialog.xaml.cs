using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.Dialogs;

public partial class HistoryDialog : Window
{
    private List<SavedReport> _reports = [];

    public HistoryDialog()
    {
        InitializeComponent();
        LoadReports();
    }

    private void LoadReports()
    {
        _reports = ReportHistoryService.LoadAll();
        ReportList.ItemsSource = null;
        ReportList.ItemsSource = _reports;

        if (_reports.Count == 0)
            DetailHeader.Text = "No previous runs found.";
    }

    private void ReportList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        var report     = ReportList.SelectedItem as SavedReport;
        bool hasSelection = report != null;

        DeleteButton.IsEnabled  = hasSelection;
        ExportButton.IsEnabled  = hasSelection;
        SummaryCards.Visibility = hasSelection ? Visibility.Visible : Visibility.Collapsed;
        EmptyHint.Visibility    = hasSelection ? Visibility.Collapsed : Visibility.Visible;

        if (report == null)
        {
            DetailHeader.Text      = "Select a run to view details";
            DetailGrid.ItemsSource = null;
            return;
        }

        DetailHeader.Text      = $"{report.DisplayDate}  —  {report.TotalCount} files  —  {report.DurationDisplay}";
        SuccessCard.Text       = report.SuccessCount.ToString();
        FailedCard.Text        = report.FailedCount.ToString();
        SkippedCard.Text       = report.SkippedCount.ToString();
        DetailGrid.ItemsSource = report.Items;
    }

    private void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (ReportList.SelectedItem is not SavedReport report) return;

        var result = MessageBox.Show(
            $"Delete the report from {report.DisplayDate}?",
            "Delete Report", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes) return;

        ReportHistoryService.Delete(report);
        LoadReports();
    }

    private void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        if (ReportList.SelectedItem is not SavedReport report) return;

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

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
