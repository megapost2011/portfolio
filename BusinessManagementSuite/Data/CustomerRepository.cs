using System;
using System.Collections.Generic;
using System.Data;

using BusinessManagementSuite.Models;

namespace BusinessManagementSuite.Data;

public sealed class CustomerRepository : RepositoryBase<Customer>
{
    //---------------------------------------------------------
    // Table
    //---------------------------------------------------------

    protected override string TableName => "Customers";

    //---------------------------------------------------------
    // Map
    //---------------------------------------------------------

    protected override Customer Map(DataRow r)
    {
        return new Customer
        {
            Id           = Convert.ToInt64(r["Id"]),
            CustomerCode = r["CustomerCode"]?.ToString() ?? "",
            CustomerName = r["CustomerName"]?.ToString() ?? "",
            ContactPerson = r["ContactPerson"]?.ToString() ?? "",
            Phone        = r["Phone"]?.ToString() ?? "",
            Email        = r["Email"]?.ToString() ?? "",
            Address      = r["Address"]?.ToString() ?? "",
            PostalCode   = r["PostalCode"]?.ToString() ?? "",
            Rank         = r["Rank"]?.ToString() ?? "",
            IsActive     = Convert.ToBoolean(r["IsActive"]),
            CreatedAt    = Convert.ToDateTime(r["CreatedAt"]),
            UpdatedAt    = Convert.ToDateTime(r["UpdatedAt"])
        };
    }

    //---------------------------------------------------------
    // GetByCode
    //---------------------------------------------------------

    public Customer? GetByCode(string code)
    {
        DataTable table = Database.ExecuteReader(
            "SELECT * FROM Customers WHERE CustomerCode=@Code LIMIT 1;",
            new() { ["@Code"] = code });

        if (table.Rows.Count == 0) return null;
        return Map(table.Rows[0]);
    }

    //---------------------------------------------------------
    // CodeExists
    //---------------------------------------------------------

    public bool CodeExists(string code)
    {
        object? result = Database.ExecuteScalar(
            "SELECT COUNT(*) FROM Customers WHERE CustomerCode=@Code;",
            new() { ["@Code"] = code });

        return Convert.ToInt32(result) > 0;
    }

    //---------------------------------------------------------
    // GetActiveCustomers
    //---------------------------------------------------------

    public List<Customer> GetActiveCustomers()
        => Find("IsActive=1");

    //---------------------------------------------------------
    // Search (parameterized – no SQL injection)
    //---------------------------------------------------------

    public List<Customer> Search(string keyword)
    {
        if (string.IsNullOrWhiteSpace(keyword))
            return GetAll();

        string like = $"%{keyword}%";

        return Find(
            """
            CustomerCode LIKE @Kw
            OR CustomerName LIKE @Kw
            OR Phone        LIKE @Kw
            OR Email        LIKE @Kw
            """,
            new() { ["@Kw"] = like });
    }

    //---------------------------------------------------------
    // Insert
    //---------------------------------------------------------

    public void Insert(Customer c)
    {
        Database.Execute(
            """
            INSERT INTO Customers
            (CustomerCode, CustomerName, ContactPerson, Phone, Email,
             Address, PostalCode, Rank, IsActive, CreatedAt, UpdatedAt)
            VALUES
            (@Code, @Name, @Contact, @Phone, @Email,
             @Address, @Postal, @Rank, @Active,
             datetime('now','localtime'), datetime('now','localtime'));
            """,
            new()
            {
                ["@Code"]    = c.CustomerCode,
                ["@Name"]    = c.CustomerName,
                ["@Contact"] = c.ContactPerson,
                ["@Phone"]   = c.Phone,
                ["@Email"]   = c.Email,
                ["@Address"] = c.Address,
                ["@Postal"]  = c.PostalCode,
                ["@Rank"]    = c.Rank,
                ["@Active"]  = c.IsActive ? 1 : 0
            });
    }

    //---------------------------------------------------------
    // Update
    //---------------------------------------------------------

    public void Update(Customer c)
    {
        Database.Execute(
            """
            UPDATE Customers SET
                CustomerCode  = @Code,
                CustomerName  = @Name,
                ContactPerson = @Contact,
                Phone         = @Phone,
                Email         = @Email,
                Address       = @Address,
                PostalCode    = @Postal,
                Rank          = @Rank,
                IsActive      = @Active,
                UpdatedAt     = datetime('now','localtime')
            WHERE Id = @Id;
            """,
            new()
            {
                ["@Code"]    = c.CustomerCode,
                ["@Name"]    = c.CustomerName,
                ["@Contact"] = c.ContactPerson,
                ["@Phone"]   = c.Phone,
                ["@Email"]   = c.Email,
                ["@Address"] = c.Address,
                ["@Postal"]  = c.PostalCode,
                ["@Rank"]    = c.Rank,
                ["@Active"]  = c.IsActive ? 1 : 0,
                ["@Id"]      = c.Id
            });
    }

    //---------------------------------------------------------
    // Save (insert or update)
    //---------------------------------------------------------

    public void Save(Customer c)
    {
        if (c.Id == 0)
            Insert(c);
        else
            Update(c);
    }

    //---------------------------------------------------------
    // GenerateCustomerCode
    //---------------------------------------------------------

    public string GenerateCustomerCode()
    {
        object? result = Database.ExecuteScalar(
            "SELECT MAX(CustomerCode) FROM Customers;");

        if (result == null || result == DBNull.Value)
            return "C000001";

        string code = result.ToString() ?? "C000000";

        if (code.Length < 2 ||
            !int.TryParse(code.Substring(1), out int number))
            return "C000001";

        return $"C{number + 1:000000}";
    }

    //---------------------------------------------------------
    // Counts
    //---------------------------------------------------------

    public int ActiveCount()
    {
        object? result = Database.ExecuteScalar(
            "SELECT COUNT(*) FROM Customers WHERE IsActive=1;");

        return Convert.ToInt32(result);
    }

    public int InactiveCount()
    {
        object? result = Database.ExecuteScalar(
            "SELECT COUNT(*) FROM Customers WHERE IsActive=0;");

        return Convert.ToInt32(result);
    }
}
