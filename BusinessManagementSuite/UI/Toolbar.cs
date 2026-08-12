using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace BusinessManagementSuite.UI;

public sealed class Toolbar : UserControl
{
    private readonly ToolStrip toolStrip;

    private readonly Dictionary<string, ToolStripButton> buttons =
        new();

    public event EventHandler? NewClicked;
    public event EventHandler? EditClicked;
    public event EventHandler? DeleteClicked;
    public event EventHandler? RefreshClicked;
    public event EventHandler? CsvClicked;
    public event EventHandler? ExcelClicked;
    public event EventHandler? PrintClicked;

    public Toolbar()
    {
        Height = Theme.ToolbarHeight;

        Dock = DockStyle.Top;

        BackColor = Theme.Header;

        toolStrip = new ToolStrip
        {
            Dock = DockStyle.Fill,
            GripStyle = ToolStripGripStyle.Hidden,
            BackColor = Theme.Header,
            ForeColor = Theme.Text,
            RenderMode = ToolStripRenderMode.System
        };

        Controls.Add(toolStrip);

        AddButton("NEW", "＋ 新規");
        AddButton("EDIT", "✎ 編集");
        AddButton("DELETE", "✖ 削除");

        toolStrip.Items.Add(new ToolStripSeparator());

        AddButton("REFRESH", "↻ 更新");

        toolStrip.Items.Add(new ToolStripSeparator());

        AddButton("CSV", "CSV");
        AddButton("EXCEL", "Excel");

        toolStrip.Items.Add(new ToolStripSeparator());

        AddButton("PRINT", "🖨 印刷");
    }

    private void AddButton(
        string key,
        string text)
    {
        ToolStripButton button = new()
        {
            Text = text,
            Tag = key,
            DisplayStyle =
                ToolStripItemDisplayStyle.Text,

            AutoSize = false,

            Width = 85,

            Height = 32,

            ForeColor = Theme.Text,

            BackColor = Theme.Button
        };

        button.Click += Button_Click;

        buttons.Add(key, button);

        toolStrip.Items.Add(button);
    }

    public ToolStripButton? GetButton(string key)
    {
        if (buttons.TryGetValue(key, out var button))
            return button;

        return null;
    }

    public void EnableButton(
        string key,
        bool enabled)
    {
        if (buttons.TryGetValue(key, out var button))
        {
            button.Enabled = enabled;
        }
    }

    private void Button_Click(object? sender, EventArgs e)
    {
        if (sender is not ToolStripButton button)
            return;

        string key = button.Tag?.ToString() ?? "";

        switch (key)
        {
            case "NEW":
                NewClicked?.Invoke(this, EventArgs.Empty);
                break;

            case "EDIT":
                EditClicked?.Invoke(this, EventArgs.Empty);
                break;

            case "DELETE":
                DeleteClicked?.Invoke(this, EventArgs.Empty);
                break;

            case "REFRESH":
                RefreshClicked?.Invoke(this, EventArgs.Empty);
                break;

            case "CSV":
                CsvClicked?.Invoke(this, EventArgs.Empty);
                break;

            case "EXCEL":
                ExcelClicked?.Invoke(this, EventArgs.Empty);
                break;

            case "PRINT":
                PrintClicked?.Invoke(this, EventArgs.Empty);
                break;
        }
    }

    public void EnableAll()
    {
        foreach (ToolStripButton button in buttons.Values)
        {
            button.Enabled = true;
        }
    }

    public void DisableAll()
    {
        foreach (ToolStripButton button in buttons.Values)
        {
            button.Enabled = false;
        }
    }

    public void SetReadOnlyMode()
    {
        EnableButton("NEW", false);
        EnableButton("EDIT", false);
        EnableButton("DELETE", false);

        EnableButton("REFRESH", true);
        EnableButton("CSV", true);
        EnableButton("EXCEL", true);
        EnableButton("PRINT", true);
    }

    public void SetEditMode()
    {
        EnableButton("NEW", true);
        EnableButton("EDIT", true);
        EnableButton("DELETE", true);

        EnableButton("REFRESH", true);
        EnableButton("CSV", true);
        EnableButton("EXCEL", true);
        EnableButton("PRINT", true);
    }

    public void SetButtonText(
        string key,
        string text)
    {
        if (buttons.TryGetValue(key, out ToolStripButton? button))
        {
            button.Text = text;
        }
    }

    public void SetButtonVisible(
        string key,
        bool visible)
    {
        if (buttons.TryGetValue(key, out ToolStripButton? button))
        {
            button.Visible = visible;
        }
    }

    public IReadOnlyDictionary<string, ToolStripButton> Buttons
    {
        get
        {
            return buttons;
        }
    }
}