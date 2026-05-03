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
            var ex = args.Exception;
            var msg = ex.Message;
            if (ex.InnerException != null)
                msg += $"\n\nInner: {ex.InnerException.Message}";
            MessageBox.Show($"Unexpected error:\n\n{msg}",
                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };

        if (e.Args.Contains("--screenshot"))
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            var runner = new Screenshots.ScreenshotRunner();
            try
            {
                await runner.RunAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Screenshot error:\n\n{ex.InnerException?.Message ?? ex.Message}",
                    "Screenshot Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown(1);
            }
        }
        else
        {
            var splash = new SplashWindow();
            splash.Show();
            await Task.Delay(2000);
            new MainWindow().Show();
            splash.Close();
        }
    }
}
