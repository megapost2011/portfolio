namespace BusinessManagementSuite.Models;

public sealed class Supplier
{
    public long Id { get; set; }

    public string SupplierCode { get; set; } = "";

    public string SupplierName { get; set; } = "";

    public string ContactPerson { get; set; } = "";

    public string Phone { get; set; } = "";

    public string Email { get; set; } = "";

    public string Address { get; set; } = "";

    public bool IsActive { get; set; } = true;

    public DateTime CreatedAt { get; set; } = DateTime.Now;

    public DateTime UpdatedAt { get; set; } = DateTime.Now;
}