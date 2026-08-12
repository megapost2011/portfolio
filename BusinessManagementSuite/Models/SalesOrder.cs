namespace BusinessManagementSuite.Models;

public sealed class SalesOrder
{
    public long Id { get; set; }

    public string OrderNo { get; set; } = "";

    public long CustomerId { get; set; }

    public DateTime OrderDate { get; set; } = DateTime.Today;

    public decimal Subtotal { get; set; }

    public decimal Tax { get; set; }

    public decimal Total { get; set; }

    public string Status { get; set; } = "Draft";

    public string Memo { get; set; } = "";

    public DateTime CreatedAt { get; set; } = DateTime.Now;
}