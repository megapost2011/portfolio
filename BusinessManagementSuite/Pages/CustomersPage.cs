using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

using BusinessManagementSuite.Models;
using BusinessManagementSuite.Services;
using BusinessManagementSuite.UI;

namespace BusinessManagementSuite.Pages;

public sealed class CustomersPage : PageBase
{
    //---------------------------------------------------------
    // Service
    //---------------------------------------------------------

    private readonly CustomerService _service = new();

    //---------------------------------------------------------
    // Constructor
    //---------------------------------------------------------

    public CustomersPage()
    {
        BuildGrid();
        Reload();
    }

    //---------------------------------------------------------
    // Grid columns
    //---------------------------------------------------------

    private void BuildGrid()
    {
        Grid.AutoGenerateColumns = false;
        Grid.Columns.Clear();

        AddTextColumn("CustomerCode", "Code",     120);
        AddTextColumn("CustomerName", "Customer", 220);
        AddTextColumn("Phone",        "Phone",    150);
        AddTextColumn("Email",        "Email",    220);
        AddCheckColumn("IsActive",    "Active",    80);

        Grid.SelectionChanged += (_, _) => UpdateStatus();
        Grid.CellDoubleClick  += (_, e) =>
        {
            if (e.RowIndex >= 0) EditCurrentCustomer();
        };
    }

    private void AddTextColumn(
        string property, string header, int fillWeight)
    {
        Grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = property,
            HeaderText       = header,
            FillWeight       = fillWeight
        });
    }

    private void AddCheckColumn(
        string property, string header, int width)
    {
        Grid.Columns.Add(new DataGridViewCheckBoxColumn
        {
            DataPropertyName = property,
            HeaderText = header,
            FillWeight = width
        });
    }

    //---------------------------------------------------------
    // Selected customer
    //---------------------------------------------------------

    public Customer? SelectedCustomer
        => Grid.CurrentRow?.DataBoundItem as Customer;

    public long CurrentCustomerId
        => SelectedCustomer?.Id ?? 0;

    public bool HasSelection() => SelectedCustomer != null;

    //---------------------------------------------------------
    // Reload
    //---------------------------------------------------------

    public override void RefreshData() => Reload();

    private void Reload()
    {
        try
        {
            List<Customer> list = _service.GetAll();
            Grid.DataSource = null;
            Grid.DataSource = list;
            UpdateStatus();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.Message, "Customers",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    //---------------------------------------------------------
    // Search override
    //---------------------------------------------------------

    protected override void SearchChanged(string keyword)
    {
        try
        {
            List<Customer> result =
                string.IsNullOrWhiteSpace(keyword)
                    ? _service.GetAll()
                    : _service.Search(keyword);

            Grid.DataSource = null;
            Grid.DataSource = result;
            SetStatus($"{result.Count} customers");
        }
        catch (Exception ex)
        {
            SetStatus($"Search error: {ex.Message}");
        }
    }

    //---------------------------------------------------------
    // Status
    //---------------------------------------------------------

    private void UpdateStatus()
    {
        int total = Grid.Rows.Count;

        if (SelectedCustomer == null)
        {
            SetStatus($"{total} customers");
        }
        else
        {
            SetStatus(
                $"{total} customers    " +
                $"Selected: {SelectedCustomer.CustomerCode} - " +
                $"{SelectedCustomer.CustomerName}");
        }
    }

    //---------------------------------------------------------
    // CRUD
    //---------------------------------------------------------

    public override void CreateNew() => NewCustomer();
    public override void EditSelected() => EditCurrentCustomer();
    public override void DeleteSelected() => DeleteCurrentCustomer();

    public void NewCustomer()
    {
        Customer customer = _service.Create();

        using CustomersEditDialog dlg = new(customer);
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        try
        {
            _service.Save(dlg.Customer);
            Reload();
            SetStatus("Customer created.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.Message, "New Customer",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void EditCurrentCustomer()
    {
        if (SelectedCustomer == null) return;

        // 編集用コピーを作成
        Customer editCopy = new()
        {
            Id           = SelectedCustomer.Id,
            CustomerCode = SelectedCustomer.CustomerCode,
            CustomerName = SelectedCustomer.CustomerName,
            Phone        = SelectedCustomer.Phone,
            Email        = SelectedCustomer.Email,
            Address      = SelectedCustomer.Address,
            IsActive     = SelectedCustomer.IsActive
        };

        using CustomersEditDialog dlg = new(editCopy);
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        try
        {
            _service.Save(dlg.Customer);
            Reload();
            SetStatus("Customer updated.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.Message, "Edit Customer",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void DeleteCurrentCustomer()
    {
        if (SelectedCustomer == null) return;

        if (MessageBox.Show(
                $"Delete '{SelectedCustomer.CustomerName}'?",
                "Delete Customer",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        try
        {
            _service.Delete(CurrentCustomerId);
            Reload();
            SetStatus("Customer deleted.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.Message, "Delete Customer",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    //---------------------------------------------------------
    // Export / Print (stubs)
    //---------------------------------------------------------

    public override void ExportCsv()
    {
        using SaveFileDialog dlg = new()
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = "Customers.csv"
        };

        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            _service.ExportCsv(dlg.FileName);
            SetStatus("CSV exported.");
        }
        catch (NotImplementedException)
        {
            MessageBox.Show(
                "CSV export will be implemented in Phase 5.",
                "Export");
        }
    }

    public override void ExportExcel()
    {
        MessageBox.Show(
            "Excel export will be implemented in Phase 5.",
            "Export");
    }

    public override void Print()
    {
        MessageBox.Show(
            "Print will be implemented in Phase 6.",
            "Print");
    }

    //---------------------------------------------------------
    // Statistics
    //---------------------------------------------------------

    public CustomerStatistics GetStatistics()
        => _service.GetStatistics();
}
