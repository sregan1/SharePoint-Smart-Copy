using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Converters;

public class BoolToVisibilityConverter : IValueConverter
{
    public bool Invert { get; set; }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        bool b = value is bool bv && bv;
        if (Invert) b = !b;
        return b ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => value is Visibility v && v == Visibility.Visible;
}

public class IntToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is int i && i > 0 ? Visibility.Visible : Visibility.Collapsed;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

public class EqualToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value == null || parameter == null) return Visibility.Collapsed;
        return value.ToString() == parameter.ToString() ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

// String-compares value to parameter (same semantics as EqualToVisibilityConverter)
// but returns bool — used by DataTriggers and two-way RadioButton.IsChecked bindings.
// ConvertBack writes the parameter into the source when checked (enum-aware), so a
// group of radio buttons can bind one enum property without code-behind handlers.
public class EqualToBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value != null && parameter != null && value.ToString() == parameter.ToString();

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => value is true && parameter != null
            ? (targetType.IsEnum ? Enum.Parse(targetType, parameter.ToString()!) : parameter)
            : Binding.DoNothing;
}

// String-compares two bound values — used where the comparison target is itself a
// binding (e.g. step indicator: CurrentStep == StepInfo.Index).
public class MultiEqualToBoolConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        => values.Length == 2 && values[0]?.ToString() == values[1]?.ToString();

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

// Numerically compares two bound ints: values[0] > values[1].
// Used by the step indicator to mark completed steps (CurrentStep > StepInfo.Index).
public class MultiGreaterThanToBoolConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        => values.Length == 2
           && int.TryParse(values[0]?.ToString(), out var a)
           && int.TryParse(values[1]?.ToString(), out var b)
           && a > b;

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

public class NotEqualToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value == null || parameter == null) return Visibility.Visible;
        return value.ToString() != parameter.ToString() ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

public class StringToColorBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string hex)
        {
            try { return new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex)); }
            catch { }
        }
        return Brushes.Black;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

public class CopyStatusToColorConverter : IValueConverter
{
    // Resolves theme palette brushes so log/report colors follow Light/Dark mode;
    // the RGB fallbacks match the light palette for safety outside an app context.
    private static Brush Themed(string key, Color fallback)
        => Application.Current?.TryFindResource(key) as Brush ?? new SolidColorBrush(fallback);

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is CopyStatus s)
        {
            return s switch
            {
                CopyStatus.Success => Themed("SuccessBrush", Color.FromRgb(16, 124, 16)),
                CopyStatus.Failed  => Themed("DangerBrush", Color.FromRgb(164, 38, 44)),
                CopyStatus.Copying => Themed("AccentBrush", Color.FromRgb(0, 120, 212)),
                CopyStatus.Skipped   => Themed("TextTertiaryBrush", Color.FromRgb(121, 119, 117)),
                CopyStatus.Cancelled => Themed("TextTertiaryBrush", Color.FromRgb(121, 119, 117)),
                _                    => Themed("TextSecondaryBrush", Color.FromRgb(50, 49, 48))
            };
        }
        return Themed("TextPrimaryBrush", Colors.Black);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

public class InvertBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => !(value is bool b && b);

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => !(value is bool b && b);
}

public class NullToVisibilityConverter : IValueConverter
{
    public bool Invert { get; set; }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        bool isNull = value == null || (value is string s && string.IsNullOrEmpty(s));
        bool show   = Invert ? isNull : !isNull;
        return show ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}
