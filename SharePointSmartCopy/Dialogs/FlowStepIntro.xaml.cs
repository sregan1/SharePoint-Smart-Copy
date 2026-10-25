using System.Windows;
using System.Windows.Controls;

namespace SharePointSmartCopy.Dialogs;

// Title, one-line "what happens" and an expandable "why" for a step of the Flows dialog. See FlowStepIntro.xaml.
public partial class FlowStepIntro : UserControl
{
    public static readonly DependencyProperty TitleProperty = Register(nameof(Title));
    public static readonly DependencyProperty WhatProperty = Register(nameof(What));
    public static readonly DependencyProperty WhyProperty = Register(nameof(Why));

    public FlowStepIntro() => InitializeComponent();

    public string Title { get => (string)GetValue(TitleProperty); set => SetValue(TitleProperty, value); }
    public string What { get => (string)GetValue(WhatProperty); set => SetValue(WhatProperty, value); }
    public string Why { get => (string)GetValue(WhyProperty); set => SetValue(WhyProperty, value); }

    private static DependencyProperty Register(string name) =>
        DependencyProperty.Register(name, typeof(string), typeof(FlowStepIntro), new PropertyMetadata(string.Empty));
}
