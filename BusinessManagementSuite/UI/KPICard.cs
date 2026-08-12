using System.Drawing;
using System.Windows.Forms;

namespace BusinessManagementSuite.UI;

public class KPICard : Panel
{
    private readonly Label titleLabel;
    private readonly Label valueLabel;

    public string Title
    {
        get => titleLabel.Text;
        set => titleLabel.Text = value;
    }

    public string Value
    {
        get => valueLabel.Text;
        set => valueLabel.Text = value;
    }

    public KPICard()
    {
        Width = 220;
        Height = 110;

        Margin = new Padding(10);

        BackColor = Theme.Card;

        BorderStyle = BorderStyle.FixedSingle;

        titleLabel = new Label
        {
            Dock = DockStyle.Top,
            Height = 32,
            Font = Theme.CardTitleFont,
            ForeColor = Theme.TextSecondary,
            Padding = new Padding(10, 8, 0, 0)
        };

        valueLabel = new Label
        {
            Dock = DockStyle.Fill,
            Font = Theme.CardValueFont,
            ForeColor = Theme.TextLight,
            TextAlign = ContentAlignment.MiddleCenter
        };

        Controls.Add(valueLabel);
        Controls.Add(titleLabel);
    }

    public void SetValue(decimal value)
    {
        Value = value.ToString("N0");
    }

    public void SetValue(int value)
    {
        Value = value.ToString();
    }

    public void SetCurrency(decimal value)
    {
        Value = value.ToString("C0");
    }
}