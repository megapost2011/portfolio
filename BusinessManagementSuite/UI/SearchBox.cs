using System;
using System.Drawing;
using System.Windows.Forms;

namespace BusinessManagementSuite.UI;

public class SearchBox : UserControl
{
    private readonly Label icon;
    private readonly TextBox textBox;

    public event EventHandler? SearchTextChanged;

    public string SearchText
    {
        get => textBox.Text;
        set => textBox.Text = value;
    }

    public SearchBox()
    {
        Height = 40;
        Width = 320;

        BackColor = Theme.Card;

        BorderStyle = BorderStyle.FixedSingle;

        icon = new Label
        {
            Text = "🔍",
            Width = 40,
            Dock = DockStyle.Left,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI Emoji", 12F),
            ForeColor = Theme.TextSecondary
        };

        textBox = new TextBox
        {
            BorderStyle = BorderStyle.None,
            Dock = DockStyle.Fill,
            Margin = new Padding(0),
            BackColor = Theme.Card,
            ForeColor = Theme.Text,
            Font = Theme.DefaultFont
        };

        textBox.TextChanged += (_, _) =>
        {
            SearchTextChanged?.Invoke(this, EventArgs.Empty);
        };

        Controls.Add(textBox);
        Controls.Add(icon);

        Padding = new Padding(5, 10, 5, 5);
    }

    public void Clear()
    {
        textBox.Clear();
    }

    public void FocusSearch()
    {
        textBox.Focus();
    }
}