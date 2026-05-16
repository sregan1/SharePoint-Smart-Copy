using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.Dialogs;

public partial class SettingsDialog : Window
{
    private readonly AuthService _authService;
    private readonly AppSettings _original;
    private readonly ObservableCollection<AzureRegistration> _registrations = [];

    public AppSettings Result { get; private set; }

    public SettingsDialog(AppSettings current, AuthService authService)
    {
        InitializeComponent();
        _authService = authService;
        _original    = current;
        Result       = current;

        foreach (var r in current.Registrations)
            _registrations.Add(new AzureRegistration { Name = r.Name, ClientId = r.ClientId, TenantId = r.TenantId });

        RegistrationList.ItemsSource = _registrations;

        if (_registrations.Count > 0)
        {
            var idx = Math.Clamp(current.ActiveRegistrationIndex, 0, _registrations.Count - 1);
            RegistrationList.SelectedIndex = idx;
        }
    }

    private void RegistrationList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        RemoveButton.IsEnabled = RegistrationList.SelectedItem != null;
    }

    private void AddRegistration_Click(object sender, RoutedEventArgs e)
    {
        var reg = new AzureRegistration { Name = "New Registration" };
        _registrations.Add(reg);
        RegistrationList.SelectedItem = reg;
    }

    private void RemoveRegistration_Click(object sender, RoutedEventArgs e)
    {
        if (RegistrationList.SelectedItem is not AzureRegistration reg) return;
        var idx = _registrations.IndexOf(reg);
        _registrations.Remove(reg);
        if (_registrations.Count > 0)
            RegistrationList.SelectedIndex = Math.Min(idx, _registrations.Count - 1);
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (_registrations.Count == 0)
        {
            MessageBox.Show("Add at least one app registration.", "Validation",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var activeIdx = Math.Max(0, RegistrationList.SelectedIndex);
        var active    = _registrations[activeIdx];

        if (string.IsNullOrWhiteSpace(active.ClientId))
        {
            MessageBox.Show("Client ID is required for the selected registration.", "Validation",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Result = new AppSettings
        {
            Registrations           = _registrations.ToList(),
            ActiveRegistrationIndex = activeIdx,
            SourceUrl               = _original.SourceUrl,
            TargetUrl               = _original.TargetUrl,
            PreferredCopyMode       = _original.PreferredCopyMode,
        };

        Result.Save();
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private async void SignOutButton_Click(object sender, RoutedEventArgs e)
    {
        await _authService.SignOutAsync();
        MessageBox.Show("Signed out. You will be prompted to sign in on the next connection.",
            "Signed Out", MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
