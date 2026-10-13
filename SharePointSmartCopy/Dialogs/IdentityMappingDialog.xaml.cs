using System.Windows;
using SharePointSmartCopy.Models;

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
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        _settings.IdentityMapping = new IdentityMappingSettings
        {
            SourceDomain = SourceDomainBox.Text.Trim(),
            TargetDomain = TargetDomainBox.Text.Trim(),
            Overrides    = OverridesBox.Text,
        };
        try { _settings.Save(); }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, Title, MessageBoxButton.OK, MessageBoxImage.Warning);
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
