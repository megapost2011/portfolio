namespace BusinessManagementSuite.Models;

public sealed class Product
{
    public long Id { get; set; }

    public string ProductCode { get; set; } = "";

    public string ProductName { get; set; } = "";

    public string Category { get; set; } = "";

    public string Unit { get; set; } = "";

    public decimal UnitPrice { get; set; }

    public decimal CostPrice { get; set; }

    public int StockQuantity { get; set; }

    public int SafetyStock { get; set; }

    public string Location { get; set; } = "";

    public bool IsActive { get; set; } = true;

    public DateTime CreatedAt { get; set; } = DateTime.Now;

    public DateTime UpdatedAt { get; set; } = DateTime.Now;
}