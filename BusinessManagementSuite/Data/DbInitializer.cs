using System;

namespace BusinessManagementSuite.Data;

public static class DbInitializer
{
    //---------------------------------------------------------
    // Initialize
    //---------------------------------------------------------

    public static void Initialize()
    {
        Console.WriteLine();
        Console.WriteLine("========================================");
        Console.WriteLine(" Business Management Suite");
        Console.WriteLine(" Database Initializer");
        Console.WriteLine("========================================");
        Console.WriteLine();

        try
        {
            InitializeDatabase();
            ExecuteSqlScripts();
            UpdateSchemaVersion();
            RunMaintenance();
            ValidateSeedData();

            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine(" Database initialization completed.");
            Console.WriteLine("========================================");
            Console.WriteLine();
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine(" Database initialization failed.");
            Console.WriteLine("----------------------------------------");
            Console.WriteLine(ex.Message);
            Console.WriteLine("========================================");
            Console.ResetColor();
            throw;
        }
    }

    private static void InitializeDatabase()
    {
        Console.WriteLine("[Database] Initialize");
        Database.Initialize();
    }

    private static void ExecuteSqlScripts()
    {
        Console.WriteLine("[Database] Execute SQL");
        SqlRunner.ExecuteAll();
    }

    private static void UpdateSchemaVersion()
    {
        Console.WriteLine("[Database] Migration");
        MigrationService.UpdateVersion();
    }

    private static void RunMaintenance()
    {
        Console.WriteLine("[Database] Maintenance");
        DatabaseMaintenance.Run();
    }

    private static void ValidateSeedData()
    {
        Console.WriteLine("[Database] Seed Validation");
        SeedValidator.Validate();
    }

    //---------------------------------------------------------
    // Rebuild
    //---------------------------------------------------------

    public static void Rebuild()
    {
        Console.WriteLine("Rebuilding database...");
        if (Database.Exists())
            Database.DeleteDatabase();
        Initialize();
    }

    public static void Backup(string folder)
        => Database.Backup(folder);

    public static bool DatabaseExists()
        => Database.Exists();

    public static bool CheckIntegrity()
        => Database.IntegrityCheck();
}
