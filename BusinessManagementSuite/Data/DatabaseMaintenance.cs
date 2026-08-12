using System;
using System.Data;

namespace BusinessManagementSuite.Data;

public static class DatabaseMaintenance
{
    //---------------------------------------------------------
    // Run All
    //---------------------------------------------------------

    public static void Run()
    {
        Console.WriteLine();
        Console.WriteLine("----------------------------------------");
        Console.WriteLine(" Database Maintenance");
        Console.WriteLine("----------------------------------------");

        Analyze();
        Optimize();
        Reindex();
        Vacuum();

        if (!IntegrityCheck())
            throw new Exception("Database integrity check failed.");

        if (!ForeignKeyCheck())
            throw new Exception("Foreign key check failed.");

        Console.WriteLine("----------------------------------------");
        Console.WriteLine();
    }

    //---------------------------------------------------------
    // Individual operations
    //---------------------------------------------------------

    public static void Analyze()
    {
        Console.Write("ANALYZE ... ");
        Database.Execute("ANALYZE;");
        Console.WriteLine("OK");
    }

    public static void Optimize()
    {
        Console.Write("PRAGMA optimize ... ");
        Database.Execute("PRAGMA optimize;");
        Console.WriteLine("OK");
    }

    public static void Reindex()
    {
        Console.Write("REINDEX ... ");
        Database.Execute("REINDEX;");
        Console.WriteLine("OK");
    }

    public static void Vacuum()
    {
        Console.Write("VACUUM ... ");
        Database.Vacuum();
        Console.WriteLine("OK");
    }

    public static bool IntegrityCheck()
    {
        Console.Write("Integrity Check ... ");
        object? result = Database.ExecuteScalar("PRAGMA integrity_check;");
        bool ok = result?.ToString() == "ok";
        Console.WriteLine(ok ? "OK" : "FAILED");
        return ok;
    }

    /// <summary>
    /// foreign_key_check は行が返ってこなければ OK（DataTable 版）
    /// </summary>
    public static bool ForeignKeyCheck()
    {
        Console.Write("Foreign Key Check ... ");
        DataTable table = Database.ExecuteReader("PRAGMA foreign_key_check;");
        bool ok = table.Rows.Count == 0;
        Console.WriteLine(ok ? "OK" : "FAILED");
        return ok;
    }

    public static bool QuickCheck()
    {
        Console.Write("Quick Check ... ");
        object? result = Database.ExecuteScalar("PRAGMA quick_check;");
        bool ok = result?.ToString() == "ok";
        Console.WriteLine(ok ? "OK" : "FAILED");
        return ok;
    }

    public static void Checkpoint()
    {
        Console.Write("Checkpoint ... ");
        Database.Execute("PRAGMA wal_checkpoint(FULL);");
        Console.WriteLine("OK");
    }

    public static void Shrink()
    {
        Checkpoint();
        Vacuum();
    }

    public static void PrintStatistics()
    {
        Console.WriteLine();
        Console.WriteLine("Database Statistics");
        Console.WriteLine("---------------------------");
        Console.WriteLine($"UserVersion : {Database.GetUserVersion()}");
        Console.WriteLine($"Page Count  : {Database.ExecuteScalar("PRAGMA page_count;")}");
        Console.WriteLine($"Page Size   : {Database.ExecuteScalar("PRAGMA page_size;")}");
        Console.WriteLine($"Cache Size  : {Database.ExecuteScalar("PRAGMA cache_size;")}");
        Console.WriteLine();
    }
}
