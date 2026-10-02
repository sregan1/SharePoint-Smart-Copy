using SharePointSmartCopy.Localization;
using System.Windows;

namespace SharePointSmartCopy;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        var settings = Models.AppSettings.Load();
        Localization.Loc.Apply(settings.Language);
        if (Localization.Loc.IsRtl)
            EventManager.RegisterClassHandler(typeof(Window), FrameworkElement.LoadedEvent,
                new EventHandler((s, _) => ((Window)s!).FlowDirection = Localization.Loc.FlowDirection));
        Services.ThemeManager.Apply(settings.Theme);
        DispatcherUnhandledException += (_, args) =>
        {
            var ex  = args.Exception;
            var msg = ex.Message;
            if (ex.InnerException != null)
                msg += "\n\n" + Loc.T("Dlg_InnerError", ex.InnerException.Message);
            MessageBox.Show(Loc.T("Dlg_UnexpectedError", msg),
                Loc.T("Dlg_ErrorTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };

        InitDemoMode(e);
        if (_demoStarted) return;

        new MainWindow().Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        base.OnExit(e);
        // Force-terminate the process so MSAL/Graph SDK background threads don't
        // keep the process alive after the main window closes.
        Environment.Exit(0);
    }

    private bool _demoStarted;
    partial void InitDemoMode(StartupEventArgs e);  // implemented in App.Demo.cs when present
}
