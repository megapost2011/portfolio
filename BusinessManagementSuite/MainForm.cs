using System;
using System.Drawing;
using System.Windows.Forms;

using BusinessManagementSuite.Pages;
using BusinessManagementSuite.UI;

namespace BusinessManagementSuite;

public sealed class MainForm : Form
{
    //--------------------------------------------------------
    // UI Components
    //--------------------------------------------------------

    private readonly Sidebar   sidebar;
    private readonly Toolbar   toolbar;
    private readonly PageHost  pageHost;
    private readonly Panel     workspace;
    private readonly StatusStrip           status;
    private readonly ToolStripStatusLabel  statusLabel;

    //--------------------------------------------------------
    // Constructor
    //--------------------------------------------------------

    public MainForm()
    {
        sidebar     = new Sidebar();
        toolbar     = new Toolbar();
        pageHost    = new PageHost();
        workspace   = new Panel();
        status      = new StatusStrip();
        statusLabel = new ToolStripStatusLabel();

        InitializeComponent();
        ThemeManager.Apply(this);
    }

    //--------------------------------------------------------
    // Initialize
    //--------------------------------------------------------

    private void InitializeComponent()
    {
        SuspendLayout();

        Text            = "Business Management Suite";
        StartPosition   = FormStartPosition.CenterScreen;
        WindowState     = FormWindowState.Maximized;
        MinimumSize     = new Size(1400, 850);
        Font            = Theme.DefaultFont;
        BackColor       = Theme.Background;

        BuildSidebar();
        BuildToolbar();
        BuildWorkspace();
        BuildStatusBar();
        RegisterPages();
        WireEvents();

        Controls.Add(workspace);
        Controls.Add(toolbar);
        Controls.Add(sidebar);
        Controls.Add(status);

        ResumeLayout(true);
    }

    //--------------------------------------------------------
    // Sidebar
    //--------------------------------------------------------

    private void BuildSidebar()
    {
        sidebar.SetItems(new[]
        {
            new NavigationItem("Dashboard",  "🏠", "Dashboard"),
            new NavigationItem("Customers",  "👥", "Customers"),
            new NavigationItem("Products",   "📦", "Products"),
            new NavigationItem("Inventory",  "📋", "Inventory"),
            new NavigationItem("Purchasing", "🚚", "Purchasing"),
            new NavigationItem("Sales",      "🛒", "Sales"),
            new NavigationItem("Analytics",  "📈", "Analytics"),
            new NavigationItem("Users",      "👤", "Users"),
            new NavigationItem("Settings",   "⚙",  "Settings")
        });
    }

    //--------------------------------------------------------
    // Toolbar
    //--------------------------------------------------------

    private void BuildToolbar()
    {
        toolbar.Dock = DockStyle.Top;
    }

    //--------------------------------------------------------
    // Workspace
    //--------------------------------------------------------

    private void BuildWorkspace()
    {
        workspace.Dock      = DockStyle.Fill;
        workspace.BackColor = Theme.Surface;
        workspace.Padding   = new Padding(0);
        workspace.Controls.Add(pageHost);
    }

    //--------------------------------------------------------
    // StatusBar
    //--------------------------------------------------------

    private void BuildStatusBar()
    {
        status.Dock       = DockStyle.Bottom;
        statusLabel.Text  = "Ready";
        status.Items.Add(statusLabel);
    }

    //--------------------------------------------------------
    // Register Pages
    //--------------------------------------------------------

    private void RegisterPages()
    {
        pageHost.RegisterPage("Dashboard",  new DashboardPage());
        pageHost.RegisterPage("Customers",  new CustomersPage());
        pageHost.RegisterPage("Products",   new ProductsPage());
        pageHost.RegisterPage("Inventory",  new InventoryPage());
        pageHost.RegisterPage("Purchasing", new PurchasingPage());
        pageHost.RegisterPage("Sales",      new SalesPage());
        pageHost.RegisterPage("Analytics",  new AnalyticsPage());
        pageHost.RegisterPage("Users",      new UsersPage());
        pageHost.RegisterPage("Settings",   new SettingsPage());

        pageHost.ShowPage("Dashboard");
    }

    //--------------------------------------------------------
    // Events
    //--------------------------------------------------------

    private void WireEvents()
    {
        sidebar.PageChanged += Sidebar_PageChanged;

        toolbar.NewClicked     += (_, _) => CurrentPage()?.CreateNew();
        toolbar.EditClicked    += (_, _) => CurrentPage()?.EditSelected();
        toolbar.DeleteClicked  += (_, _) => CurrentPage()?.DeleteSelected();
        toolbar.RefreshClicked += (_, _) => CurrentPage()?.RefreshData();
        toolbar.CsvClicked     += (_, _) => CurrentPage()?.ExportCsv();
        toolbar.ExcelClicked   += (_, _) => CurrentPage()?.ExportExcel();
        toolbar.PrintClicked   += (_, _) => CurrentPage()?.Print();
    }

    private void Sidebar_PageChanged(
        object? sender,
        NavigationItem item)
    {
        pageHost.ShowPage(item.PageKey);
        SetStatus(item.Text);
    }

    //--------------------------------------------------------
    // Helpers
    //--------------------------------------------------------

    private PageBase? CurrentPage()
    {
        string? key = pageHost.CurrentPageKey;
        if (key == null) return null;
        return pageHost.Pages.TryGetValue(key, out var page)
            ? page as PageBase
            : null;
    }

    private void SetStatus(string message)
        => statusLabel.Text = message;

    private void EnableEditing(bool enabled)
    {
        if (enabled)
            toolbar.SetEditMode();
        else
            toolbar.SetReadOnlyMode();
    }

    //--------------------------------------------------------
    // Form events
    //--------------------------------------------------------

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        sidebar.SelectPage("Dashboard");
        EnableEditing(true);
        SetStatus("Ready");
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        statusLabel.Text = $"Ready    {Width} × {Height}";
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            status.Dispose();
            workspace.Dispose();
            sidebar.Dispose();
            toolbar.Dispose();
            pageHost.Dispose();
        }

        base.Dispose(disposing);
    }
}
