using System.Windows;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.Dialogs;

// Shown after the source scan when the pre-copy check flagged something, before anything is copied.
// ShowDialog() == true means "copy everything else"; anything else (Cancel, Esc, the close button) cancels the run.
public partial class PrecheckDialog : Window
{
    private readonly IReadOnlyList<PrecheckIssue> _issues;

    public PrecheckDialog(IReadOnlyList<PrecheckIssue> issues)
    {
        InitializeComponent();
        _issues = issues;
        IssueGrid.ItemsSource = issues;
        IntroText.Text = Loc.T("Ui_PrecheckIntro", issues.Count);
    }

    private void SaveList_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Filter   = Loc.T("VM_FilterCsvTxt"),
            FileName = $"PrecheckIssues_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog(this) != true) return;
        try { CsvReportExport.SavePrecheckList(dlg.FileName, _issues); }
        catch (Exception ex)
        {
            MessageBox.Show(this, Loc.T("Dlg_IdentitySaveFailed", ex.Message), Title,
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void CancelCopy_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    private void CopyRest_Click(object sender, RoutedEventArgs e)   => DialogResult = true;
}
