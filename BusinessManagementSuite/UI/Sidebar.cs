using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace BusinessManagementSuite.UI;

public sealed class Sidebar : UserControl
{
    private readonly FlowLayoutPanel menuPanel;

    private readonly Label titleLabel;

    private readonly List<Button> buttons = new();

    private readonly List<NavigationItem> items = new();

    private Button? selectedButton;

    public event EventHandler<NavigationItem>? PageChanged;

    public Sidebar()
    {
        Width = Theme.SidebarWidth;

        Dock = DockStyle.Left;

        BackColor = Theme.Sidebar;

        titleLabel = new Label
        {
            Dock = DockStyle.Top,
            Height = 70,
            Text = "Business Management Suite",
            Font = new Font("Segoe UI", 16F, FontStyle.Bold),
            ForeColor = Theme.TextLight,
            TextAlign = ContentAlignment.MiddleCenter
        };

        menuPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true,
            Padding = new Padding(8)
        };

        Controls.Add(menuPanel);
        Controls.Add(titleLabel);
    }

    public void SetItems(IEnumerable<NavigationItem> navigationItems)
    {
        items.Clear();

        items.AddRange(navigationItems);

        menuPanel.Controls.Clear();

        buttons.Clear();

        foreach (NavigationItem item in items)
        {
            Button button = CreateButton(item);

            buttons.Add(button);

            menuPanel.Controls.Add(button);
        }

        if (buttons.Count > 0)
        {
            SelectButton(buttons[0]);
        }
    }

    private Button CreateButton(NavigationItem item)
    {
        Button button = new Button
        {
            Width = Width - 24,
            Height = 46,

            Margin = new Padding(0, 0, 0, 6),

            FlatStyle = FlatStyle.Flat,

            TextAlign = ContentAlignment.MiddleLeft,

            Text = item.ToString(),

            Tag = item,

            BackColor = Theme.Sidebar,

            ForeColor = Theme.TextLight,

            Font = Theme.DefaultFont,

            Cursor = Cursors.Hand
        };

        button.FlatAppearance.BorderSize = 0;

        button.Click += Button_Click;

        button.MouseEnter += Button_MouseEnter;

        button.MouseLeave += Button_MouseLeave;

        return button;
    }

    private void Button_Click(object? sender, EventArgs e)
    {
        if (sender is not Button button)
            return;

        SelectButton(button);

        if (button.Tag is NavigationItem item)
        {
            PageChanged?.Invoke(this, item);
        }
    }

    private void Button_MouseEnter(object? sender, EventArgs e)
    {
        if (sender is not Button button)
            return;

        if (button == selectedButton)
            return;

        button.BackColor = Theme.ButtonHover;
    }

    private void Button_MouseLeave(object? sender, EventArgs e)
    {
        if (sender is not Button button)
            return;

        if (button == selectedButton)
            return;

        button.BackColor = Theme.Sidebar;
    }

    private void SelectButton(Button button)
    {
        foreach (Button b in buttons)
        {
            b.BackColor = Theme.Sidebar;
            b.ForeColor = Theme.TextLight;
        }

        selectedButton = button;

        button.BackColor = Theme.Accent;
        button.ForeColor = Color.White;
    }

    public void SelectPage(string pageKey)
    {
        foreach (Button button in buttons)
        {
            if (button.Tag is NavigationItem item &&
                item.PageKey == pageKey)
            {
                SelectButton(button);
                PageChanged?.Invoke(this, item);
                return;
            }
        }
    }

    public string? SelectedPage
    {
        get
        {
            if (selectedButton?.Tag is NavigationItem item)
                return item.PageKey;

            return null;
        }
    }

    public IReadOnlyList<NavigationItem> Items
    {
        get
        {
            return items;
        }
    }
}