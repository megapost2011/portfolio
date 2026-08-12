using System;
using System.Drawing;
using System.Windows.Forms;

using BusinessManagementSuite.Models;
using BusinessManagementSuite.UI;

namespace BusinessManagementSuite.Pages;

public partial class CustomersEditDialog : Form
{
    //---------------------------------------------------------
    // Controls
    //---------------------------------------------------------

    private readonly TextBox txtCode;

    private readonly TextBox txtName;

    private readonly TextBox txtPhone;

    private readonly TextBox txtEmail;

    private readonly TextBox txtAddress;

    private readonly CheckBox chkActive;

    private readonly Button btnOK;

    private readonly Button btnCancel;

    //---------------------------------------------------------
    // Customer
    //---------------------------------------------------------

    public Customer Customer { get; }

    //---------------------------------------------------------
    // Constructor
    //---------------------------------------------------------

    public CustomersEditDialog(Customer? customer = null)
    {
        Customer = customer ?? new Customer();

        txtCode = new TextBox();

        txtName = new TextBox();

        txtPhone = new TextBox();

        txtEmail = new TextBox();

        txtAddress = new TextBox();

        chkActive = new CheckBox();

        btnOK = new Button();

        btnCancel = new Button();

        InitializeComponent();

        LoadCustomer();
    }

    //---------------------------------------------------------
    // Initialize
    //---------------------------------------------------------

    private void InitializeComponent()
    {
        SuspendLayout();

        Text = Customer.Id == 0
            ? "New Customer"
            : "Edit Customer";

        StartPosition =
            FormStartPosition.CenterParent;

        FormBorderStyle =
            FormBorderStyle.FixedDialog;

        MaximizeBox = false;

        MinimizeBox = false;

        ShowInTaskbar = false;

        Width = 520;

        Height = 420;

        Font = Theme.DefaultFont;

        BackColor = Theme.Surface;

        BuildLayout();

        ResumeLayout(false);
    }

    //---------------------------------------------------------
    // Layout
    //---------------------------------------------------------

    private void BuildLayout()
    {
        TableLayoutPanel layout = new()
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(15),
            ColumnCount = 2,
            RowCount = 7
        };

        layout.ColumnStyles.Add(
            new ColumnStyle(SizeType.Absolute,120));

        layout.ColumnStyles.Add(
            new ColumnStyle(SizeType.Percent,100));

        AddRow(layout,0,"Customer Code",txtCode);

        AddRow(layout,1,"Customer Name",txtName);

        AddRow(layout,2,"Phone",txtPhone);

        AddRow(layout,3,"Email",txtEmail);

        AddRow(layout,4,"Address",txtAddress);

        chkActive.Text = "Active";

        layout.Controls.Add(chkActive,1,5);

        FlowLayoutPanel buttons = new()
        {
            Dock = DockStyle.Fill,
            FlowDirection =
                FlowDirection.RightToLeft
        };

        btnOK.Text = "OK";

        btnOK.Width = 100;

        btnOK.Click += BtnOK_Click;

        btnCancel.Text = "Cancel";

        btnCancel.Width = 100;

        btnCancel.DialogResult =
            DialogResult.Cancel;

        buttons.Controls.Add(btnCancel);

        buttons.Controls.Add(btnOK);

        layout.Controls.Add(buttons,1,6);

        Controls.Add(layout);

        AcceptButton = btnOK;

        CancelButton = btnCancel;
    }

    //---------------------------------------------------------
    // Add Row
    //---------------------------------------------------------

    private static void AddRow(
        TableLayoutPanel layout,
        int row,
        string caption,
        Control control)
    {
        Label label = new()
        {
            Text = caption,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };

        control.Dock = DockStyle.Fill;

        layout.Controls.Add(label,0,row);

        layout.Controls.Add(control,1,row);
    }

    //---------------------------------------------------------
    // Load Customer
    //---------------------------------------------------------

    private void LoadCustomer()
    {
        txtCode.Text = Customer.CustomerCode;
        txtName.Text = Customer.CustomerName;
        txtPhone.Text = Customer.Phone;
        txtEmail.Text = Customer.Email;
        txtAddress.Text = Customer.Address;

        chkActive.Checked = Customer.IsActive;
    }

    //---------------------------------------------------------
    // Save Customer
    //---------------------------------------------------------

    private void SaveCustomer()
    {
        Customer.CustomerCode = txtCode.Text.Trim();

        Customer.CustomerName = txtName.Text.Trim();

        Customer.Phone = txtPhone.Text.Trim();

        Customer.Email = txtEmail.Text.Trim();

        Customer.Address = txtAddress.Text.Trim();

        Customer.IsActive = chkActive.Checked;
    }

    //---------------------------------------------------------
    // Validation
    //---------------------------------------------------------

    private bool ValidateInput()
    {
        if (string.IsNullOrWhiteSpace(txtCode.Text))
        {
            MessageBox.Show(
                "Customer Code is required.",
                "Validation",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);

            txtCode.Focus();

            return false;
        }

        if (string.IsNullOrWhiteSpace(txtName.Text))
        {
            MessageBox.Show(
                "Customer Name is required.",
                "Validation",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);

            txtName.Focus();

            return false;
        }

        if (!string.IsNullOrWhiteSpace(txtEmail.Text))
        {
            try
            {
                _ = new System.Net.Mail.MailAddress(txtEmail.Text);
            }
            catch
            {
                MessageBox.Show(
                    "Invalid email address.",
                    "Validation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                txtEmail.Focus();

                return false;
            }
        }

        return true;
    }

    //---------------------------------------------------------
    // OK Button
    //---------------------------------------------------------

    private void BtnOK_Click(
        object? sender,
        EventArgs e)
    {
        if (!ValidateInput())
            return;

        SaveCustomer();

        DialogResult = DialogResult.OK;

        Close();
    }

    //---------------------------------------------------------
    // Clear
    //---------------------------------------------------------

    public void Clear()
    {
        txtCode.Clear();

        txtName.Clear();

        txtPhone.Clear();

        txtEmail.Clear();

        txtAddress.Clear();

        chkActive.Checked = true;
    }

    //---------------------------------------------------------
    // ReadOnly
    //---------------------------------------------------------

    public void SetReadOnly(bool value)
    {
        txtCode.ReadOnly = value;

        txtName.ReadOnly = value;

        txtPhone.ReadOnly = value;

        txtEmail.ReadOnly = value;

        txtAddress.ReadOnly = value;

        chkActive.Enabled = !value;

        btnOK.Enabled = !value;
    }

    //---------------------------------------------------------
    // Focus
    //---------------------------------------------------------

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);

        txtCode.Focus();

        txtCode.SelectAll();
    }

    //---------------------------------------------------------
    // Escape Key
    //---------------------------------------------------------

    protected override bool ProcessCmdKey(
        ref Message msg,
        Keys keyData)
    {
        if (keyData == Keys.Escape)
        {
            DialogResult = DialogResult.Cancel;

            Close();

            return true;
        }

        return base.ProcessCmdKey(
            ref msg,
            keyData);
    }

    //---------------------------------------------------------
    // F5 Clear
    //---------------------------------------------------------

    protected override bool ProcessDialogKey(Keys keyData)
    {
        if (keyData == Keys.F5)
        {
            Clear();

            return true;
        }

        return base.ProcessDialogKey(keyData);
    }
}