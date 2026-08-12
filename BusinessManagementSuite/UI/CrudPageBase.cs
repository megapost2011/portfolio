using System.Collections.Generic;
using System.Windows.Forms;

namespace BusinessManagementSuite.UI;

/// <summary>
/// CRUD ページ共通ベース
/// </summary>
public abstract class CrudPageBase<T> : PageBase
{
    //---------------------------------------------------------
    // Data Source
    //---------------------------------------------------------

    protected readonly BindingSource BindingSource;
    protected readonly List<T> Items;

    //---------------------------------------------------------
    // Constructor
    //---------------------------------------------------------

    protected CrudPageBase()
    {
        BindingSource = new BindingSource();
        Items = new List<T>();
        Grid.AutoGenerateColumns = false;
        Grid.DataSource = BindingSource;
    }

    //---------------------------------------------------------
    // Load
    //---------------------------------------------------------

    public override void RefreshData()
    {
        Items.Clear();

        foreach (T item in LoadItems())
            Items.Add(item);

        BindingSource.DataSource = null;
        BindingSource.DataSource = Items;
        SetStatus($"{Items.Count} records");
    }

    //---------------------------------------------------------
    // CRUD
    //---------------------------------------------------------

    public override void CreateNew()
    {
        T? item = CreateItem();
        if (item == null) return;
        SaveNew(item);
        RefreshData();
    }

    public override void EditSelected()
    {
        if (CurrentRow == null) return;
        if (CurrentRow.DataBoundItem is not T item) return;
        if (!EditItem(item)) return;
        Save(item);
        RefreshData();
    }

    public override void DeleteSelected()
    {
        if (CurrentRow == null) return;
        if (CurrentRow.DataBoundItem is not T item) return;

        DialogResult result = MessageBox.Show(
            "Delete selected record?",
            "Confirm",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result != DialogResult.Yes) return;

        Delete(item);
        RefreshData();
    }

    //---------------------------------------------------------
    // Abstract
    //---------------------------------------------------------

    protected abstract IEnumerable<T> LoadItems();
    protected abstract T? CreateItem();
    protected abstract bool EditItem(T item);
    protected abstract void Save(T item);
    protected abstract void SaveNew(T item);
    protected abstract void Delete(T item);
}
