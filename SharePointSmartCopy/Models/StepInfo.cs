namespace SharePointSmartCopy.Models;

// One entry in the wizard step indicator. Instantiated from XAML (x:Array in
// MainWindow resources), so it needs a parameterless constructor and settable properties.
public class StepInfo
{
    public int    Index   { get; set; }
    public string Display { get; set; } = string.Empty; // bubble text: "1".."5" or "✓"
    public string Label   { get; set; } = string.Empty; // caption under the bubble
    public bool   IsFinal { get; set; }                 // final (Report) step turns green when active
}
