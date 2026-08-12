using System.Drawing;

namespace BusinessManagementSuite.UI;

/// <summary>
/// Business Management Suite 共通テーマ
/// </summary>
public static class Theme
{
    //========================
    // Main Colors
    //========================

    public static readonly Color Background =
        Color.FromArgb(30, 30, 30);

    public static readonly Color Surface =
        Color.FromArgb(37, 37, 38);

    public static readonly Color Sidebar =
        Color.FromArgb(31, 41, 55);

    public static readonly Color Header =
        Color.FromArgb(45, 45, 48);

    public static readonly Color Card =
        Color.FromArgb(51, 51, 55);

    //========================
    // Accent
    //========================

    public static readonly Color Accent =
        Color.FromArgb(0, 122, 204);

    public static readonly Color Success =
        Color.FromArgb(16, 185, 129);

    public static readonly Color Warning =
        Color.FromArgb(245, 158, 11);

    public static readonly Color Danger =
        Color.FromArgb(220, 53, 69);

    //========================
    // Text
    //========================

    public static readonly Color Text =
        Color.Gainsboro;

    public static readonly Color TextSecondary =
        Color.Silver;

    public static readonly Color TextLight =
        Color.White;

    //========================
    // Border
    //========================

    public static readonly Color Border =
        Color.FromArgb(62, 62, 64);

    //========================
    // Grid
    //========================

    public static readonly Color GridBackground =
        Surface;

    public static readonly Color GridAlternate =
        Color.FromArgb(45, 45, 48);

    public static readonly Color GridHeader =
        Accent;

    public static readonly Color GridSelection =
        Color.FromArgb(28, 151, 234);

    //========================
    // Button
    //========================

    public static readonly Color Button =
        Color.FromArgb(63, 63, 70);

    public static readonly Color ButtonHover =
        Accent;

    public static readonly Color ButtonPressed =
        Color.FromArgb(0, 90, 158);

    //========================
    // Fonts
    //========================

    public static readonly Font DefaultFont =
        new("Segoe UI", 9F);

    public static readonly Font TitleFont =
        new("Segoe UI", 16F, FontStyle.Bold);

    public static readonly Font CardTitleFont =
        new("Segoe UI", 10F, FontStyle.Bold);

    public static readonly Font CardValueFont =
        new("Segoe UI", 18F, FontStyle.Bold);

    //========================
    // Sizes
    //========================

    public const int SidebarWidth = 230;

    public const int HeaderHeight = 70;

    public const int ToolbarHeight = 42;

    public const int StatusbarHeight = 24;

    public const int CardRadius = 10;

    public const int DefaultMargin = 12;
}