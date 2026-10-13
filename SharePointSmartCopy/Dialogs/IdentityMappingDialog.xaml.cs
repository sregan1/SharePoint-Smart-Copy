using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.Dialogs;

// Edits the cross-tenant identity mapping (domain swap + explicit overrides) stored in AppSettings.
public partial class IdentityMappingDialog : Window
{
    private readonly AppSettings _settings;

    public IdentityMappingDialog(AppSettings settings)
    {
        InitializeComponent();
        _settings = settings;
        var m = settings.IdentityMapping;
        SourceDomainBox.Text = m.SourceDomain;
        TargetDomainBox.Text = m.TargetDomain;
        OverridesBox.Text    = m.Overrides;
        UpdateSummary();
    }

    private void Input_Changed(object sender, TextChangedEventArgs e) => UpdateSummary();

    // Tells the user what the mapping will actually do — a malformed line used to be skipped silently,
    // leaving a person unmapped with no sign anything was wrong.
    private void UpdateSummary()
    {
        if (SummaryText == null) return; // fires during InitializeComponent before the field exists

        var valid   = IdentityMapper.ParseOverrides(OverridesBox.Text).Count;
        var invalid = IdentityMapper.FindInvalidLines(OverridesBox.Text);
        var parts   = new List<string> { Loc.T("Ui_IdentityMappingsFound", valid) };
        bool warn   = false;

        if (invalid.Count > 0)
        {
            parts.Add(Loc.T("Ui_IdentityLinesInvalid", invalid.Count, string.Join(", ", invalid.Take(8)) + (invalid.Count > 8 ? "…" : "")));
            warn = true;
        }

        var src = SourceDomainBox.Text.Trim().TrimStart('@');
        var tgt = TargetDomainBox.Text.Trim().TrimStart('@');
        if (src.Length > 0 && src.Equals(tgt, StringComparison.OrdinalIgnoreCase))
        {
            parts.Add(Loc.T("Ui_IdentityDomainsSame"));
            warn = true;
        }

        SummaryText.Text = string.Join("  ", parts);
        SummaryText.Foreground = (Brush)FindResource(warn ? "WarningTextBrush" : "TextTertiaryBrush");
    }

    // Appends a CSV/text file of "source,target" lines to the overrides.
    private void ImportCsv_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "CSV / text (*.csv;*.txt)|*.csv;*.txt|All files (*.*)|*.*",
        };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            var text = System.IO.File.ReadAllText(dlg.FileName).Trim();
            if (text.Length == 0) return;
            var existing = OverridesBox.Text.TrimEnd();
            OverridesBox.Text = existing.Length == 0 ? text : existing + Environment.NewLine + text;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, Loc.T("Dlg_IdentityImportFailed", ex.Message), Title,
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        // Don't save a mapping that silently drops lines: say which ones are unreadable and stay open.
        var invalid = IdentityMapper.FindInvalidLines(OverridesBox.Text);
        if (invalid.Count > 0)
        {
            MessageBox.Show(this, Loc.T("Dlg_IdentityFixLines", string.Join(", ", invalid.Take(8)) + (invalid.Count > 8 ? "…" : "")),
                Title, MessageBoxButton.OK, MessageBoxImage.Warning);
            OverridesBox.Focus();
            return;
        }

        var previous = _settings.IdentityMapping;
        _settings.IdentityMapping = new IdentityMappingSettings
        {
            SourceDomain = SourceDomainBox.Text.Trim(),
            TargetDomain = TargetDomainBox.Text.Trim(),
            Overrides    = OverridesBox.Text,
        };
        try
        {
            _settings.Save();
        }
        catch (Exception ex)
        {
            // Saving failed: put the old mapping back and keep the dialog open so nothing is lost
            // (it used to close as though the save had worked).
            _settings.IdentityMapping = previous;
            MessageBox.Show(this, Loc.T("Dlg_IdentitySaveFailed", ex.Message), Title,
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
