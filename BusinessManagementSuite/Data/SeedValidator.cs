using System;

namespace BusinessManagementSuite.Data;

public static class SeedValidator
{
    public static void Validate()
    {
        Console.WriteLine();
        Console.WriteLine("----------------------------------------");
        Console.WriteLine(" Seed Validation");
        Console.WriteLine("----------------------------------------");

        ValidateTableIfExists("Customers");
        ValidateTableIfExists("Products");
        ValidateTableIfExists("Suppliers");
        ValidateTableIfExists("Users");
        ValidateTableIfExists("Warehouses");
        ValidateTableIfExists("PaymentMethods");

        Console.WriteLine("----------------------------------------");
        Console.WriteLine();
    }

    private static void ValidateTableIfExists(string tableName)
    {
        // テーブルが存在しない場合はスキップ
        object? exists = Database.ExecuteScalar(
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@t;",
            new() { ["@t"] = tableName });

        if (Convert.ToInt32(exists) == 0)
        {
            Console.WriteLine($"{tableName,-24}(table not found)");
            return;
        }

        int count = GetCount(tableName);

        Console.Write($"{tableName,-24}");

        if (count > 0)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"OK ({count} rows)");
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("EMPTY");
        }

        Console.ResetColor();
    }

    public static int GetCount(string table)
    {
        object? result = Database.ExecuteScalar(
            $"SELECT COUNT(*) FROM {table};");
        return Convert.ToInt32(result);
    }

    public static bool Exists(string table) => GetCount(table) > 0;

    public static bool HasPaymentMethods() => Exists("PaymentMethods");
    public static bool HasWarehouses()     => Exists("Warehouses");
    public static bool HasUsers()          => Exists("Users");
    public static bool HasCustomers()      => Exists("Customers");
    public static bool HasProducts()       => Exists("Products");
    public static bool HasSuppliers()      => Exists("Suppliers");

    public static bool IsEmpty()
        => !HasCustomers() && !HasProducts()
        && !HasSuppliers() && !HasUsers();

    public static void PrintSummary()
    {
        Console.WriteLine();
        Console.WriteLine("Database Summary");
        Console.WriteLine("---------------------------");
        Console.WriteLine($"Customers      : {GetCount("Customers")}");
        Console.WriteLine($"Products       : {GetCount("Products")}");
        Console.WriteLine($"Suppliers      : {GetCount("Suppliers")}");
        Console.WriteLine($"Users          : {GetCount("Users")}");
        Console.WriteLine();
    }
}
