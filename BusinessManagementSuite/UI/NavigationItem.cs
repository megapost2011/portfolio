using System.Drawing;

namespace BusinessManagementSuite.UI;

public sealed class NavigationItem
{
    public string Text { get; }

    public string Icon { get; }

    public string PageKey { get; }

    public NavigationItem(
        string text,
        string icon,
        string pageKey)
    {
        Text = text;
        Icon = icon;
        PageKey = pageKey;
    }

    public override string ToString()
    {
        return $"{Icon}  {Text}";
    }
}