using System;
using System.Drawing;
using System.Windows.Forms;

using BusinessManagementSuite.UI;

namespace BusinessManagementSuite.Pages;

public sealed class DashboardPage : UserControl
{
    //----------------------------------------------------------
    // Controls
    //----------------------------------------------------------

    private readonly FlowLayoutPanel kpiPanel;

    private readonly SearchBox searchBox;

    private readonly SplitContainer split;

    private readonly BMSGrid orderGrid;

    private readonly BMSGrid warningGrid;

    //----------------------------------------------------------
    // Constructor
    //----------------------------------------------------------

    public DashboardPage()
    {
        Dock = DockStyle.Fill;

        BackColor = Theme.Surface;

        kpiPanel = new FlowLayoutPanel();

        searchBox = new SearchBox();

        split = new SplitContainer();

        orderGrid = new BMSGrid();

        warningGrid = new BMSGrid();

        InitializePage();
    }

    //----------------------------------------------------------
    // Initialize
    //----------------------------------------------------------

    private void InitializePage()
    {
        BuildKPI();

        BuildSearch();

        BuildGrids();

        Controls.Add(split);

        Controls.Add(searchBox);

        Controls.Add(kpiPanel);
    }

    //----------------------------------------------------------
    // KPI
    //----------------------------------------------------------

    private void BuildKPI()
    {
        kpiPanel.Dock = DockStyle.Top;

        kpiPanel.Height = 140;

        kpiPanel.Padding = new Padding(15);

        kpiPanel.WrapContents = false;

        KPICard sales = new()
        {
            Title = "Today's Sales"
        };

        sales.SetCurrency(128000);

        KPICard orders = new()
        {
            Title = "Orders"
        };

        orders.SetValue(52);

        KPICard inventory = new()
        {
            Title = "Inventory Alert"
        };

        inventory.SetValue(3);

        KPICard profit = new()
        {
            Title = "Gross Profit"
        };

        profit.SetCurrency(68400);

        kpiPanel.Controls.Add(sales);

        kpiPanel.Controls.Add(orders);

        kpiPanel.Controls.Add(inventory);

        kpiPanel.Controls.Add(profit);
    }

    //----------------------------------------------------------
    // Search
    //----------------------------------------------------------

    private void BuildSearch()
    {
        searchBox.Dock = DockStyle.Top;
    }

    //----------------------------------------------------------
    // Grids
    //----------------------------------------------------------

    private void BuildGrids()
    {
        split.Dock = DockStyle.Fill;

        split.Orientation = Orientation.Horizontal;

        split.SplitterDistance = 320;

        split.Panel1.Controls.Add(orderGrid);

        split.Panel2.Controls.Add(warningGrid);

        BuildOrderGrid();

        BuildWarningGrid();
    }
    //----------------------------------------------------------
    // Order Grid
    //----------------------------------------------------------

    private void BuildOrderGrid()
    {
        orderGrid.Columns.Clear();

        orderGrid.Columns.Add("OrderNo", "Order No");
        orderGrid.Columns.Add("Customer", "Customer");
        orderGrid.Columns.Add("Product", "Product");
        orderGrid.Columns.Add("Qty", "Qty");
        orderGrid.Columns.Add("Amount", "Amount");
        orderGrid.Columns.Add("Status", "Status");

        orderGrid.Rows.Add(
            "SO0001",
            "ABC Foods",
            "Hamburger",
            12,
            "¥11,760",
            "Completed");

        orderGrid.Rows.Add(
            "SO0002",
            "XYZ Office",
            "Coffee",
            25,
            "¥11,250",
            "Processing");

        orderGrid.Rows.Add(
            "SO0003",
            "Sample Store",
            "Cake",
            8,
            "¥4,160",
            "Waiting");

        orderGrid.Rows.Add(
            "SO0004",
            "Tokyo Shop",
            "Pizza",
            16,
            "¥24,000",
            "Completed");
    }

    //----------------------------------------------------------
    // Warning Grid
    //----------------------------------------------------------

    private void BuildWarningGrid()
    {
        warningGrid.Columns.Clear();

        warningGrid.Columns.Add("Code", "Code");
        warningGrid.Columns.Add("Product", "Product");
        warningGrid.Columns.Add("Stock", "Stock");
        warningGrid.Columns.Add("Safety", "Safety");
        warningGrid.Columns.Add("Location", "Location");

        warningGrid.Rows.Add(
            "P004",
            "Cake",
            6,
            20,
            "A-01");

        warningGrid.Rows.Add(
            "P010",
            "Orange Juice",
            5,
            25,
            "A-12");

        warningGrid.Rows.Add(
            "P023",
            "French Fries",
            8,
            30,
            "B-05");

        warningGrid.Rows.Add(
            "P031",
            "Milk",
            4,
            15,
            "C-02");
    }

    //----------------------------------------------------------
    // Refresh
    //----------------------------------------------------------

    public void RefreshDashboard()
    {
        // Phase3(SQLite)で実装
    }

    //----------------------------------------------------------
    // Search
    //----------------------------------------------------------

    private void Search(string keyword)
    {
        keyword = keyword.Trim().ToLower();

        foreach (DataGridViewRow row in orderGrid.Rows)
        {
            bool visible = false;

            foreach (DataGridViewCell cell in row.Cells)
            {
                if (cell.Value == null)
                    continue;

                if (cell.Value.ToString()!
                    .ToLower()
                    .Contains(keyword))
                {
                    visible = true;
                    break;
                }
            }

            row.Visible =
                visible ||
                string.IsNullOrWhiteSpace(keyword);
        }
    }

    //----------------------------------------------------------
    // Event
    //----------------------------------------------------------

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);

        searchBox.SearchTextChanged += (_, __) =>
        {
            Search(searchBox.SearchText);
        };
    }

    //----------------------------------------------------------
    // Export
    //----------------------------------------------------------

    public void ExportCsv()
    {
        // Phase4
    }

    public void ExportExcel()
    {
        // Phase4
    }

    public void Print()
    {
        // Phase5
    }
}