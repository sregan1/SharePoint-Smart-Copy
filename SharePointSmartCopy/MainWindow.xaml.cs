using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
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

    // ── History dialog ────────────────────────────────────────────────────────

    private void HistoryButton_Click(object sender, RoutedEventArgs e)
        => new Dialogs.HistoryDialog { Owner = this }.ShowDialog();

    // ── Settings dialog ───────────────────────────────────────────────────────

    private void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SettingsDialog(VM.Settings, VM.AuthService) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            VM.ApplySettings(dlg.Result);
            MessageBox.Show("Settings saved. Connect again if you were already signed in.",
                "Settings Updated", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    // ── Source tree lazy loading ──────────────────────────────────────────────

    private async void SourceTreeItem_Expanded(object sender, RoutedEventArgs e)
    {
        if (sender is TreeViewItem tvi && tvi.DataContext is SharePointNode node)
        {
            e.Handled = true;
            await VM.LoadNodeChildrenAsync(node);
        }
    }

    private void SourceNode_CheckChanged(object sender, RoutedEventArgs e)
        => VM.NextCommand.NotifyCanExecuteChanged();

    private void SelectAll_Click(object sender, RoutedEventArgs e)
    {
        VM.SelectAllSource(true);
        VM.NextCommand.NotifyCanExecuteChanged();
    }

    private void DeselectAll_Click(object sender, RoutedEventArgs e)
    {
        VM.SelectAllSource(false);
        VM.NextCommand.NotifyCanExecuteChanged();
    }

    // ── Target tree lazy loading + selection ─────────────────────────────────

    private async void TargetTreeItem_Expanded(object sender, RoutedEventArgs e)
    {
        if (sender is TreeViewItem tvi && tvi.DataContext is SharePointNode node)
        {
            e.Handled = true;
            await VM.LoadTargetNodeChildrenAsync(node);
        }
    }

    private void TargetTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is SharePointNode node)
            VM.SelectedTargetFolder = node;
    }

    // ── Progress list auto-scroll ─────────────────────────────────────────────

    private void ProgressList_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        // Defer the scroll so it runs after WPF finishes processing the CollectionChanged
        // notification — calling ScrollIntoView synchronously inside CollectionChanged can
        // cause the ItemContainerGenerator to see an inconsistent state.
        Dispatcher.BeginInvoke(() =>
        {
            if (ProgressList.Items.Count > 0)
                ProgressList.ScrollIntoView(ProgressList.Items[^1]);
        }, System.Windows.Threading.DispatcherPriority.Background);
    }

    // ── Start new copy ────────────────────────────────────────────────────────

    private void StartNewCopy_Click(object sender, RoutedEventArgs e)
    {
        // Reuse the existing AuthService so cached credentials survive into the new copy
        DataContext = new MainViewModel(VM.AuthService);
        ((System.Collections.ObjectModel.ObservableCollection<CopyResult>)VM.CopyResults)
            .CollectionChanged += ProgressList_CollectionChanged;
    }
}
