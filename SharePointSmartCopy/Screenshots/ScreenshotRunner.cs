using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using SharePointSmartCopy.Dialogs;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.ViewModels;

namespace SharePointSmartCopy.Screenshots;

internal class ScreenshotRunner
{
    private readonly string _outputDir;
    private readonly string _reportsDir;

    internal ScreenshotRunner()
    {
        // Walk up from bin/Debug/net8.0-windows to the project root (contains Docs folder)
        var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
        while (dir != null && !Directory.Exists(Path.Combine(dir.FullName, "Docs")))
            dir = dir.Parent;

        _outputDir = dir != null
            ? Path.Combine(dir.FullName, "Docs", "screenshots")
            : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "screenshots");

        Directory.CreateDirectory(_outputDir);

        _reportsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SharePointSmartCopy", "Reports");
    }

    internal async Task RunAsync()
    {
        Console.WriteLine($"Saving screenshots to: {_outputDir}");

        var vm = new MainViewModel();
        var window = new MainWindow(vm) { Width = 1280, Height = 820 };
        window.Show();
        await WaitForRenderAsync();

        // ── 01: Source connected ───────────────────────────────────────────────
        vm.SourceUrl = "https://company.sharepoint.com/sites/hr-documents";
        vm.SignedInUser = "John Smith";
        vm.SourceStatus = "✅ Connected as John Smith";
        vm.SourceConnected = true;
        vm.CurrentStep = 0;
        await CaptureAsync(window, "01_source_connected.png");

        // ── 02: Browse & Select ────────────────────────────────────────────────
        vm.CurrentStep = 1;
        Application.Current.Dispatcher.Invoke(() =>
        {
            vm.SourceLibraries.Clear();
            foreach (var lib in BuildMockSourceLibraries())
                vm.SourceLibraries.Add(lib);
        });
        await CaptureAsync(window, "02_browse.png");

        // ── 03: Target connected ───────────────────────────────────────────────
        vm.CurrentStep = 2;
        vm.TargetUrl = "https://company.sharepoint.com/sites/hr-archive";
        vm.TargetStatus = "✅ Connected";
        vm.TargetConnected = true;
        var targetLib = BuildMockTargetLibrary();
        Application.Current.Dispatcher.Invoke(() =>
        {
            vm.TargetLibraries.Clear();
            vm.TargetLibraries.Add(targetLib);
        });
        vm.SelectedTargetFolder = targetLib.Children.FirstOrDefault(c => c.Name == "HR Department");
        await CaptureAsync(window, "03_target_connected.png");

        // ── 04: Options (standard) ─────────────────────────────────────────────
        vm.CurrentStep = 3;
        vm.OverwriteFiles = true;
        vm.CopyVersions = false;
        vm.MaxParallelCopies = 4;
        Application.Current.Dispatcher.Invoke(() =>
        {
            vm.CopyJobs.Clear();
            foreach (var job in BuildMockCopyJobs())
                vm.CopyJobs.Add(job);
        });
        await CaptureAsync(window, "04_options.png");

        // ── 04b: Options with version copying ────────────────────────────────
        vm.CopyVersions = true;
        vm.CopyAllVersions = false;
        vm.MaxVersions = 5;
        await CaptureAsync(window, "04_options_versions.png");

        // ── 05: Copying in progress ────────────────────────────────────────────
        vm.CurrentStep = 4;
        vm.CopyVersions = false;
        vm.IsCopying = true;
        vm.IsCopyComplete = false;
        var inProgressResults = BuildMockCopyResultsInProgress().ToList();
        Application.Current.Dispatcher.Invoke(() =>
        {
            vm.CopyResults.Clear();
            foreach (var r in inProgressResults)
                vm.CopyResults.Add(r);
        });
        vm.TotalCount = 20;
        vm.CompletedCount = inProgressResults.Count(r =>
            r.Status is CopyStatus.Success or CopyStatus.Failed or CopyStatus.Skipped);
        vm.TotalProgress = vm.TotalCount > 0 ? vm.CompletedCount * 100.0 / vm.TotalCount : 0;
        vm.ElapsedTime = "5s";
        await CaptureAsync(window, "05_copying.png");

        // ── 06: Report ─────────────────────────────────────────────────────────
        vm.CurrentStep = 5;
        vm.IsCopying = false;
        vm.IsCopyComplete = true;
        var completedResults = BuildMockCopyResultsComplete().ToList();
        Application.Current.Dispatcher.Invoke(() =>
        {
            vm.CopyResults.Clear();
            foreach (var r in completedResults)
                vm.CopyResults.Add(r);
        });
        vm.TotalCount = completedResults.Count;
        vm.CompletedCount = vm.TotalCount;
        vm.TotalProgress = 100;
        vm.CopyDuration = "12s";
        vm.RefreshCopyStats();
        await CaptureAsync(window, "06_report.png");

        // ── 07: Settings dialog ────────────────────────────────────────────────
        var settingsDlg = new SettingsDialog(BuildMockSettings(), vm.AuthService) { Owner = window };
        settingsDlg.Show();
        await WaitForRenderAsync();
        await CaptureAsync(settingsDlg, "07_settings.png");
        settingsDlg.Close();

        // ── 08 & 08b: History dialog ───────────────────────────────────────────
        var mockReports = BuildMockReports();
        WriteMockReports(mockReports);
        try
        {
            var historyDlg = new HistoryDialog { Owner = window };
            historyDlg.Show();
            await WaitForRenderAsync();
            await CaptureAsync(historyDlg, "08_history.png");

            historyDlg.ReportList.SelectedIndex = 0;
            await WaitForRenderAsync();
            await CaptureAsync(historyDlg, "08_history_detail.png");

            historyDlg.Close();
        }
        finally
        {
            DeleteMockReports(mockReports);
        }

        Console.WriteLine("All screenshots saved.");
        window.Close();
        Application.Current.Shutdown();
    }

    // ── Rendering ──────────────────────────────────────────────────────────────

    private static async Task WaitForRenderAsync()
    {
        for (int i = 0; i < 5; i++)
            await Application.Current.Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
    }

    private async Task CaptureAsync(Window window, string filename)
    {
        await WaitForRenderAsync();

        var w = (int)window.ActualWidth;
        var h = (int)window.ActualHeight;

        var rtb = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
        rtb.Render(window);

        var enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(rtb));

        var path = Path.Combine(_outputDir, filename);
        using var fs = File.Create(path);
        enc.Save(fs);

        Console.WriteLine($"  ✓ {filename}");
    }

    // ── Mock data ──────────────────────────────────────────────────────────────

    private static IEnumerable<SharePointNode> BuildMockSourceLibraries()
    {
        var lib1 = new SharePointNode
        {
            Name = "Documents", Type = NodeType.Library,
            Id = "lib1", DriveId = "drv1", SiteId = "site1",
            HasChildren = true, IsExpanded = true
        };

        var hrFolder = new SharePointNode
        {
            Name = "HR Materials", Type = NodeType.Folder,
            Id = "f1", DriveId = "drv1", SiteId = "site1",
            IsChecked = true, HasChildren = true, Parent = lib1
        };
        hrFolder.Children.Add(new SharePointNode { Name = "__placeholder__", Type = NodeType.Folder });

        var policyFolder = new SharePointNode
        {
            Name = "Policies 2024", Type = NodeType.Folder,
            Id = "f2", DriveId = "drv1", SiteId = "site1",
            IsChecked = true, HasChildren = true, Parent = lib1
        };
        policyFolder.Children.Add(new SharePointNode { Name = "__placeholder__", Type = NodeType.Folder });

        lib1.Children.Add(hrFolder);
        lib1.Children.Add(policyFolder);
        lib1.Children.Add(new SharePointNode
        {
            Name = "Annual Report 2024.pdf", Type = NodeType.File,
            Id = "f3", DriveId = "drv1", SiteId = "site1",
            IsChecked = true, Size = 2_400_000, Parent = lib1
        });
        lib1.Children.Add(new SharePointNode
        {
            Name = "Budget Overview.xlsx", Type = NodeType.File,
            Id = "f4", DriveId = "drv1", SiteId = "site1",
            IsChecked = true, Size = 1_100_000, Parent = lib1
        });
        lib1.Children.Add(new SharePointNode
        {
            Name = "Q3 Summary.pptx", Type = NodeType.File,
            Id = "f5", DriveId = "drv1", SiteId = "site1",
            IsChecked = false, Size = 3_200_000, Parent = lib1
        });

        var lib2 = new SharePointNode
        {
            Name = "Shared Documents", Type = NodeType.Library,
            Id = "lib2", DriveId = "drv2", SiteId = "site1",
            HasChildren = true
        };
        lib2.Children.Add(new SharePointNode { Name = "__placeholder__", Type = NodeType.Folder });

        return [lib1, lib2];
    }

    private static SharePointNode BuildMockTargetLibrary()
    {
        var lib = new SharePointNode
        {
            Name = "Documents", Type = NodeType.Library,
            Id = "tlib1", DriveId = "tdrv1", SiteId = "tsite1",
            HasChildren = true, IsExpanded = true
        };

        var hrDept = new SharePointNode
        {
            Name = "HR Department", Type = NodeType.Folder,
            Id = "tf1", DriveId = "tdrv1", SiteId = "tsite1",
            HasChildren = true, Parent = lib
        };
        hrDept.Children.Add(new SharePointNode { Name = "__placeholder__", Type = NodeType.Folder });

        lib.Children.Add(hrDept);
        lib.Children.Add(new SharePointNode
        {
            Name = "Finance", Type = NodeType.Folder,
            Id = "tf2", DriveId = "tdrv1", SiteId = "tsite1",
            Parent = lib
        });
        lib.Children.Add(new SharePointNode
        {
            Name = "Marketing", Type = NodeType.Folder,
            Id = "tf3", DriveId = "tdrv1", SiteId = "tsite1",
            Parent = lib
        });

        return lib;
    }

    private static IEnumerable<CopyJob> BuildMockCopyJobs() =>
    [
        new CopyJob { SourceName = "HR Materials", SourceDisplayPath = "Documents/HR Materials", TargetDisplayPath = "HR Materials", IsFolder = true },
        new CopyJob { SourceName = "Policies 2024", SourceDisplayPath = "Documents/Policies 2024", TargetDisplayPath = "Policies 2024", IsFolder = true },
        new CopyJob { SourceName = "Annual Report 2024.pdf", SourceDisplayPath = "Documents/Annual Report 2024.pdf", TargetDisplayPath = "Annual Report 2024.pdf" },
        new CopyJob { SourceName = "Budget Overview.xlsx", SourceDisplayPath = "Documents/Budget Overview.xlsx", TargetDisplayPath = "Budget Overview.xlsx" },
    ];

    private static IEnumerable<CopyResult> BuildMockCopyResultsInProgress()
    {
        string[] done = [
            "Onboarding Guide.docx", "Benefits Summary.pdf", "Code of Conduct.pdf",
            "Leave Policy.docx", "Travel Policy.docx", "IT Security Policy.pdf",
            "Remote Work Guidelines.docx", "Performance Review Template.xlsx",
            "Org Chart 2024.pptx", "Q1 Headcount.xlsx", "Training Schedule.xlsx",
            "Recruitment Tracker.xlsx", "New Hire Checklist.docx"
        ];
        foreach (var f in done)
            yield return new CopyResult { FileName = f, SourcePath = "Documents/" + f, TargetPath = f, Status = CopyStatus.Success, VersionsCopied = 1 };

        yield return new CopyResult { FileName = "Exit Interview Form.docx", SourcePath = "Documents/Exit Interview Form.docx", TargetPath = "Exit Interview Form.docx", Status = CopyStatus.Copying };

        string[] pending = [
            "Employee Handbook.docx", "Salary Bands 2024.xlsx", "Holiday Calendar.pdf",
            "Sick Leave Policy.pdf", "Diversity Report.pdf", "Wellness Program.pptx"
        ];
        foreach (var f in pending)
            yield return new CopyResult { FileName = f, SourcePath = "Documents/" + f, TargetPath = f, Status = CopyStatus.Pending };
    }

    private static IEnumerable<CopyResult> BuildMockCopyResultsComplete()
    {
        string[] success = [
            "Onboarding Guide.docx", "Benefits Summary.pdf", "Code of Conduct.pdf",
            "Leave Policy.docx", "Travel Policy.docx", "IT Security Policy.pdf",
            "Remote Work Guidelines.docx", "Performance Review Template.xlsx",
            "Org Chart 2024.pptx", "Q1 Headcount.xlsx", "Training Schedule.xlsx",
            "Recruitment Tracker.xlsx", "New Hire Checklist.docx",
            "Exit Interview Form.docx", "Employee Handbook.docx",
            "Salary Bands 2024.xlsx", "Wellness Program.pptx"
        ];
        foreach (var f in success)
            yield return new CopyResult { FileName = f, SourcePath = "Documents/" + f, TargetPath = f, Status = CopyStatus.Success, VersionsCopied = 3, VersionsTotal = 3 };

        yield return new CopyResult { FileName = "Legacy Archive.zip", SourcePath = "Documents/Legacy Archive.zip", TargetPath = "Legacy Archive.zip", Status = CopyStatus.Failed, ErrorMessage = "File size exceeds limit" };
        yield return new CopyResult { FileName = "Holiday Calendar.pdf", SourcePath = "Documents/Holiday Calendar.pdf", TargetPath = "Holiday Calendar.pdf", Status = CopyStatus.Skipped };
        yield return new CopyResult { FileName = "Sick Leave Policy.pdf", SourcePath = "Documents/Sick Leave Policy.pdf", TargetPath = "Sick Leave Policy.pdf", Status = CopyStatus.Skipped };
    }

    private static AppSettings BuildMockSettings() => new()
    {
        Registrations =
        [
            new AzureRegistration { Name = "Contoso Production", ClientId = "a1b2c3d4-e5f6-7890-abcd-ef1234567890", TenantId = "contoso.onmicrosoft.com" },
            new AzureRegistration { Name = "Contoso Dev / Sandbox", ClientId = "b2c3d4e5-f6a7-8901-bcde-f01234567891", TenantId = "common" }
        ],
        ActiveRegistrationIndex = 0
    };

    private List<SavedReport> BuildMockReports()
    {
        var opts = new JsonSerializerOptions { WriteIndented = true, Converters = { new JsonStringEnumConverter() } };

        var items = new List<SavedReportItem>();
        string[] files = ["Annual Report 2024.pdf", "Budget Overview.xlsx", "Q3 Summary.pptx", "Org Chart 2024.pptx", "Onboarding Guide.docx", "Leave Policy.docx", "HR Policy Manual.pdf", "Training Materials.pptx"];
        foreach (var f in files)
            items.Add(new SavedReportItem { FileName = f, SourcePath = "Documents/" + f, TargetPath = f, Status = CopyStatus.Success, VersionsCopied = 2, VersionsTotal = 2 });
        items.Add(new SavedReportItem { FileName = "Legacy Backup.zip", SourcePath = "Documents/Legacy Backup.zip", TargetPath = "Legacy Backup.zip", Status = CopyStatus.Failed, ErrorMessage = "File too large" });

        return
        [
            new SavedReport
            {
                Id = "20260501_143022",
                Timestamp = DateTimeOffset.Now.AddHours(-2),
                SourceUrl = "https://company.sharepoint.com/sites/hr-documents",
                TargetUrl = "https://company.sharepoint.com/sites/hr-archive",
                SuccessCount = 8, FailedCount = 1, SkippedCount = 0, TotalCount = 9,
                Duration = TimeSpan.FromSeconds(27),
                Items = items
            },
            new SavedReport
            {
                Id = "20260430_091500",
                Timestamp = DateTimeOffset.Now.AddDays(-1),
                SourceUrl = "https://company.sharepoint.com/sites/finance-docs",
                TargetUrl = "https://company.sharepoint.com/sites/finance-archive",
                SuccessCount = 23, FailedCount = 0, SkippedCount = 2, TotalCount = 25,
                Duration = TimeSpan.FromMinutes(1) + TimeSpan.FromSeconds(44),
                Items = []
            },
            new SavedReport
            {
                Id = "20260429_160045",
                Timestamp = DateTimeOffset.Now.AddDays(-2),
                SourceUrl = "https://company.sharepoint.com/sites/projects",
                TargetUrl = "https://company.sharepoint.com/sites/archive",
                SuccessCount = 142, FailedCount = 3, SkippedCount = 5, TotalCount = 150,
                Duration = TimeSpan.FromMinutes(8) + TimeSpan.FromSeconds(12),
                Items = []
            }
        ];
    }

    private void WriteMockReports(List<SavedReport> reports)
    {
        Directory.CreateDirectory(_reportsDir);
        var opts = new JsonSerializerOptions { WriteIndented = true, Converters = { new JsonStringEnumConverter() } };
        foreach (var r in reports)
            File.WriteAllText(
                Path.Combine(_reportsDir, $"report_{r.Id}.json"),
                JsonSerializer.Serialize(r, opts));
    }

    private void DeleteMockReports(List<SavedReport> reports)
    {
        foreach (var r in reports)
        {
            var path = Path.Combine(_reportsDir, $"report_{r.Id}.json");
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
