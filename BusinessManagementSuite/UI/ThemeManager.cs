using System.Drawing;
using System.Windows.Forms;

namespace BusinessManagementSuite.UI;

public static class ThemeManager
{
    public static void Apply(Form form)
    {
        form.SuspendLayout();

        form.BackColor = Theme.Background;
        form.ForeColor = Theme.Text;
        form.Font = Theme.DefaultFont;

        ApplyControls(form.Controls);

        form.ResumeLayout(true);
    }

    private static void ApplyControls(Control.ControlCollection controls)
    {
        foreach (Control c in controls)
        {
            switch (c)
            {
                case Panel panel:

                    panel.BackColor = Theme.Surface;
                    panel.ForeColor = Theme.Text;
                    break;

                case Label label:

                    label.ForeColor = Theme.Text;
                    label.BackColor = Color.Transparent;
                    break;

                case Button button:

                    ApplyButton(button);
                    break;

                case TextBox textBox:

                    textBox.BackColor = Theme.Card;
                    textBox.ForeColor = Theme.Text;
                    textBox.BorderStyle = BorderStyle.FixedSingle;
                    break;

                case ComboBox combo:

                    combo.BackColor = Theme.Card;
                    combo.ForeColor = Theme.Text;
                    combo.FlatStyle = FlatStyle.Flat;
                    break;

                case ListBox list:

                    list.BackColor = Theme.Card;
                    list.ForeColor = Theme.Text;
                    break;

                case GroupBox group:

                    group.ForeColor = Theme.Text;
                    group.BackColor = Theme.Surface;
                    break;

                case MenuStrip menu:

                    menu.BackColor = Theme.Header;
                    menu.ForeColor = Theme.Text;
                    break;

                case StatusStrip status:

                    status.BackColor = Theme.Header;
                    status.ForeColor = Theme.Text;
                    break; 

                case ToolStrip tool:

                    tool.BackColor = Theme.Header;
                    tool.ForeColor = Theme.Text;
                    break;

                case SplitContainer split:

                    split.BackColor = Theme.Border;
                    break;

                case DataGridView grid:

                    ApplyGrid(grid);
                    break;
            }

            if (c.HasChildren)
                ApplyControls(c.Controls);
        }
    }

    private static void ApplyButton(Button b)
    {
        b.FlatStyle = FlatStyle.Flat;

        b.FlatAppearance.BorderColor = Theme.Border;

        b.FlatAppearance.BorderSize = 1;

        b.FlatAppearance.MouseOverBackColor = Theme.ButtonHover;

        b.FlatAppearance.MouseDownBackColor = Theme.ButtonPressed;

        b.BackColor = Theme.Button;

        b.ForeColor = Theme.Text;

        b.Cursor = Cursors.Hand;
    }

    private static void ApplyGrid(DataGridView grid)
    {
        grid.EnableHeadersVisualStyles = false;

        grid.BackgroundColor = Theme.GridBackground;

        grid.GridColor = Theme.Border;

        grid.BorderStyle = BorderStyle.None;

        grid.RowHeadersVisible = false;

        grid.AllowUserToAddRows = false;

        grid.AllowUserToDeleteRows = false;

        grid.AllowUserToResizeRows = false;

        grid.SelectionMode =
            DataGridViewSelectionMode.FullRowSelect;

        grid.MultiSelect = false;

        grid.AutoSizeColumnsMode =
            DataGridViewAutoSizeColumnsMode.Fill;

        grid.DefaultCellStyle.BackColor =
            Theme.GridBackground;

        grid.DefaultCellStyle.ForeColor =
            Theme.Text;

        grid.DefaultCellStyle.SelectionBackColor =
            Theme.GridSelection;

        grid.DefaultCellStyle.SelectionForeColor =
            Color.White;

        grid.AlternatingRowsDefaultCellStyle.BackColor =
            Theme.GridAlternate;

        grid.ColumnHeadersDefaultCellStyle.BackColor =
            Theme.GridHeader;

        grid.ColumnHeadersDefaultCellStyle.ForeColor =
            Color.White;

        grid.ColumnHeadersHeight = 36;
    }
}