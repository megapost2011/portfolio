using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace BusinessManagementSuite.UI;

public class BMSGrid : DataGridView
{
    public BMSGrid()
    {
        InitializeGrid();

        EnableDoubleBuffer();
    }

    private void InitializeGrid()
    {
        Dock = DockStyle.Fill;

        BackgroundColor = Theme.GridBackground;

        BorderStyle = BorderStyle.None;

        GridColor = Theme.Border;

        EnableHeadersVisualStyles = false;

        RowHeadersVisible = false;

        AllowUserToAddRows = false;

        AllowUserToDeleteRows = false;

        AllowUserToResizeRows = false;

        AllowUserToResizeColumns = true;

        AllowUserToOrderColumns = true;

        MultiSelect = false;

        ReadOnly = true;

        SelectionMode =
            DataGridViewSelectionMode.FullRowSelect;

        AutoSizeColumnsMode =
            DataGridViewAutoSizeColumnsMode.Fill;

        AutoSizeRowsMode =
            DataGridViewAutoSizeRowsMode.None;

        RowTemplate.Height = 32;

        DefaultCellStyle.BackColor =
            Theme.GridBackground;

        DefaultCellStyle.ForeColor =
            Theme.Text;

        DefaultCellStyle.SelectionBackColor =
            Theme.GridSelection;

        DefaultCellStyle.SelectionForeColor =
            Color.White;

        AlternatingRowsDefaultCellStyle.BackColor =
            Theme.GridAlternate;

        ColumnHeadersDefaultCellStyle.BackColor =
            Theme.GridHeader;

        ColumnHeadersDefaultCellStyle.ForeColor =
            Color.White;

        ColumnHeadersHeight = 36;

        ClipboardCopyMode =
            DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

        KeyDown += GridKeyDown;
    }

    private void GridKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Control && e.KeyCode == Keys.A)
        {
            SelectAll();

            e.Handled = true;
        }

        if (e.Control && e.KeyCode == Keys.C)
        {
            Clipboard.SetDataObject(GetClipboardContent());

            e.Handled = true;
        }
    }

    private void EnableDoubleBuffer()
    {
        typeof(DataGridView)
            .GetProperty(
                "DoubleBuffered",
                BindingFlags.Instance |
                BindingFlags.NonPublic)
            ?.SetValue(this, true);
    }

    public void RefreshTheme()
    {
        BackgroundColor = Theme.GridBackground;

        GridColor = Theme.Border;

        DefaultCellStyle.BackColor =
            Theme.GridBackground;

        DefaultCellStyle.ForeColor =
            Theme.Text;

        DefaultCellStyle.SelectionBackColor =
            Theme.GridSelection;

        AlternatingRowsDefaultCellStyle.BackColor =
            Theme.GridAlternate;

        ColumnHeadersDefaultCellStyle.BackColor =
            Theme.GridHeader;

        ColumnHeadersDefaultCellStyle.ForeColor =
            Color.White;
    }
}