using System.Windows;

namespace SharePointSmartCopy;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        Services.ThemeManager.Apply(Models.AppSettings.Load().Theme);
        DispatcherUnhandledException += (_, args) =>
        {
            var ex  = args.Exception;
            var msg = ex.Message;
            if (ex.InnerException != null)
                msg += $"\n\nInner: {ex.InnerException.Message}";
            MessageBox.Show($"Unexpected error:\n\n{msg}",
                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };

        InitDemoMode(e);
        if (_demoStarted) return;

        new MainWindow().Show();
    }

    private bool _demoStarted;
    partial void InitDemoMode(StartupEventArgs e);  // implemented in App.Demo.cs when present
}
