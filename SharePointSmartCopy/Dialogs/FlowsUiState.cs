using System.IO;
using System.Text.Json;

namespace SharePointSmartCopy.Dialogs;

// The little the Flows dialog remembers between sessions: the last output folder. Lives next to settings.json in
// %APPDATA%\SharePointSmartCopy. A damaged or missing file simply means the defaults.
public static class FlowsUiState
{
    public sealed record State(string? OutputPath);

    private static string Path_ { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SharePointSmartCopy", "flows-ui.json");

    public static State Load()
    {
        try
        {
            return File.Exists(Path_) ? JsonSerializer.Deserialize<State>(File.ReadAllText(Path_)) ?? new State(null) : new State(null);
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            return new State(null);
        }
    }

    public static void SaveOutputPath(string outputPath)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path_)!);
            File.WriteAllText(Path_, JsonSerializer.Serialize(new State(outputPath)));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // remembering the folder is a convenience
        }
    }
}