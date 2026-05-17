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

    private void HistoryButton_Click(object sender, RoutedEventArgs e)
        => new HistoryDialog { Owner = this }.ShowDialog();

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

    private async void SourceTreeItem_Expanded(object sender, RoutedEventArgs e)
    {
        if (sender is TreeViewItem tvi && tvi.DataContext is SharePointNode node)
        { e.Handled = true; await VM.LoadNodeChildrenAsync(node); }
    }

    private void SourceNode_CheckChanged(object sender, RoutedEventArgs e)
        => VM.NextCommand.NotifyCanExecuteChanged();

    private void SelectAll_Click(object sender, RoutedEventArgs e)
    { VM.SelectAllSource(true); VM.NextCommand.NotifyCanExecuteChanged(); }

    private void DeselectAll_Click(object sender, RoutedEventArgs e)
    { VM.SelectAllSource(false); VM.NextCommand.NotifyCanExecuteChanged(); }

    private async void TargetTreeItem_Expanded(object sender, RoutedEventArgs e)
    {
        if (sender is TreeViewItem tvi && tvi.DataContext is SharePointNode node)
        { e.Handled = true; await VM.LoadTargetNodeChildrenAsync(node); }
    }

    private void TargetTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is SharePointNode node) VM.SelectedTargetFolder = node;
    }

    private void MigrationApiMode_Click(object sender, RoutedEventArgs e)
        => VM.CopyMode = CopyMode.MigrationApi;

    private void EnhancedRestMode_Click(object sender, RoutedEventArgs e)
        => VM.CopyMode = CopyMode.EnhancedRest;

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
}
