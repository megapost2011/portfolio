namespace BusinessManagementSuite.Models;

public sealed class Customer
{
    public long Id { get; set; }

    public string CustomerCode { get; set; } = "";

    public string CustomerName { get; set; } = "";

    public string ContactPerson { get; set; } = "";

    public string Phone { get; set; } = "";

    public string Email { get; set; } = "";

    public string Address { get; set; } = "";

    public string PostalCode { get; set; } = "";

    public string Rank { get; set; } = "";

    public bool IsActive { get; set; } = true;

    public DateTime CreatedAt { get; set; } = DateTime.Now;

    public DateTime UpdatedAt { get; set; } = DateTime.Now;
}