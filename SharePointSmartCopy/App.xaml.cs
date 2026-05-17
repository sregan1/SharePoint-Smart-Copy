using System.Threading.Tasks;
using System.Windows;

namespace SharePointSmartCopy;

public partial class App : Application
{
    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
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

        var splash = new SplashWindow();
        splash.Show();
        await Task.Delay(250);
        new MainWindow().Show();
        splash.Close();
    }

    private bool _demoStarted;
    partial void InitDemoMode(StartupEventArgs e);  // implemented in App.Demo.cs when present
}
