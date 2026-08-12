using System;
using System.Windows.Forms;

namespace BusinessManagementSuite.UI;

/// <summary>
/// 全ページ共通のベースクラス
/// </summary>
public abstract class PageBase : UserControl
{
    //---------------------------------------------------------
    // Common Controls
    //---------------------------------------------------------

    protected readonly SearchBox Search;

    protected readonly BMSGrid Grid;

    protected readonly StatusStrip Status;

    protected readonly ToolStripStatusLabel StatusLabel;

    //---------------------------------------------------------
    // Constructor
    //---------------------------------------------------------

    protected PageBase()
    {
        Dock = DockStyle.Fill;

        Search = new SearchBox();

        Grid = new BMSGrid();

        Status = new StatusStrip();

        StatusLabel = new ToolStripStatusLabel();

        InitializePage();

        WireEvents();
    }

    //---------------------------------------------------------
    // Layout
    //---------------------------------------------------------

    private void InitializePage()
    {
        Search.Dock = DockStyle.Top;

        Grid.Dock = DockStyle.Fill;

        Status.Dock = DockStyle.Bottom;

        Status.Items.Add(StatusLabel);

        Controls.Add(Grid);

        Controls.Add(Search);

        Controls.Add(Status);
    }

    //---------------------------------------------------------
    // Events
    //---------------------------------------------------------

    private void WireEvents()
    {
        Search.SearchTextChanged += (_, __) =>
        {
            SearchChanged(Search.SearchText);
        };

        Grid.CellDoubleClick += (_, e) =>
        {
            if (e.RowIndex >= 0)
                EditSelected();
        };
    }

    //---------------------------------------------------------
    // Helpers
    //---------------------------------------------------------

    protected void SetStatus(string text)
    {
        StatusLabel.Text = text;
    }

    protected DataGridViewRow? CurrentRow
    {
        get
        {
            return Grid.CurrentRow;
        }
    }

    protected void ClearGrid()
    {
        Grid.Rows.Clear();
    }

    //---------------------------------------------------------
    // CRUD
    //---------------------------------------------------------

    public virtual void RefreshData()
    {
    }

    public virtual void CreateNew()
    {
    }

    public virtual void EditSelected()
    {
    }

    public virtual void DeleteSelected()
    {
    }

    //---------------------------------------------------------
    // Export
    //---------------------------------------------------------

    public virtual void ExportCsv()
    {
    }

    public virtual void ExportExcel()
    {
    }

    public virtual void Print()
    {
    }

    //---------------------------------------------------------
    // Search
    //---------------------------------------------------------

    protected virtual void SearchChanged(string keyword)
    {
        keyword = keyword.Trim().ToLower();

        foreach (DataGridViewRow row in Grid.Rows)
        {
            bool visible =
                string.IsNullOrWhiteSpace(keyword);

            if (!visible)
            {
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
            }

            row.Visible = visible;
        }

        SetStatus($"{Grid.Rows.Count} records");
    }

    //---------------------------------------------------------
    // Toolbar API
    //---------------------------------------------------------

    public virtual void ToolbarNew()
    {
        CreateNew();
    }

    public virtual void ToolbarEdit()
    {
        EditSelected();
    }

    public virtual void ToolbarDelete()
    {
        DeleteSelected();
    }

    public virtual void ToolbarRefresh()
    {
        RefreshData();
    }

    public virtual void ToolbarCsv()
    {
        ExportCsv();
    }

    public virtual void ToolbarExcel()
    {
        ExportExcel();
    }

    public virtual void ToolbarPrint()
    {
        Print();
    }
}