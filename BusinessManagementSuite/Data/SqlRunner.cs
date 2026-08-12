using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace BusinessManagementSuite.Data;

public static class SqlRunner
{
    //---------------------------------------------------------
    // Folders
    //---------------------------------------------------------

    private static readonly string SqlFolder =
        Path.Combine(AppContext.BaseDirectory, "Sql");

    private static readonly string LogFolder =
        Path.Combine(AppContext.BaseDirectory, "Logs");

    //---------------------------------------------------------
    // Execute All
    //---------------------------------------------------------

    public static void ExecuteAll()
    {
        if (!Directory.Exists(SqlFolder))
        {
            Console.WriteLine($"[SqlRunner] SQL folder not found: {SqlFolder}");
            Console.WriteLine("[SqlRunner] Skipping SQL script execution.");
            return;
        }

        string[] files =
            Directory.GetFiles(SqlFolder, "*.sql", SearchOption.TopDirectoryOnly)
            .OrderBy(Path.GetFileName)
            .ToArray();

        Console.WriteLine();
        Console.WriteLine("----------------------------------------");
        Console.WriteLine(" SQL Script Execution");
        Console.WriteLine("----------------------------------------");

        int total = files.Length;
        Stopwatch totalWatch = Stopwatch.StartNew();

        for (int i = 0; i < total; i++)
            Execute(files[i], i + 1, total);

        totalWatch.Stop();
        PrintSummary(total, totalWatch.ElapsedMilliseconds);

        Console.WriteLine("----------------------------------------");
        Console.WriteLine();
    }

    //---------------------------------------------------------
    // Execute One File
    //---------------------------------------------------------

    private static void Execute(
        string path,
        int current,
        int total)
    {
        string fileName = Path.GetFileName(path);
        Console.Write($"[{current:D2}/{total:D2}] {fileName}");

        Stopwatch sw = Stopwatch.StartNew();

        try
        {
            string sql = File.ReadAllText(path);
            Console.WriteLine($"  ({GetFileSize(path)})");

            if (!ValidateSql(sql))
                throw new Exception("SQL file is empty.");

            Database.Execute(sql);
            sw.Stop();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"      OK ({sw.ElapsedMilliseconds} ms)");
            Console.ResetColor();

            WriteLog($"OK {fileName} {sw.ElapsedMilliseconds}ms");
        }
        catch (Exception ex)
        {
            sw.Stop();
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"      FAILED ({sw.ElapsedMilliseconds} ms)");
            Console.WriteLine(ex.Message);
            Console.ResetColor();

            WriteLog($"FAILED {fileName}");
            WriteLog(ex.ToString());
            throw;
        }

        Console.WriteLine();
    }

    //---------------------------------------------------------
    // Helpers
    //---------------------------------------------------------

    public static bool Exists(string fileName)
        => File.Exists(Path.Combine(SqlFolder, fileName));

    public static string GetSqlFolder() => SqlFolder;

    public static int Count()
        => Directory.Exists(SqlFolder)
            ? Directory.GetFiles(SqlFolder, "*.sql").Length
            : 0;

    //---------------------------------------------------------
    // Private helpers
    //---------------------------------------------------------

    private static void WriteLog(string message)
    {
        Directory.CreateDirectory(LogFolder);

        string logFile = Path.Combine(
            LogFolder,
            $"DbInitialize_{DateTime.Now:yyyyMMdd}.log");

        File.AppendAllText(
            logFile,
            $"{DateTime.Now:HH:mm:ss} {message}{Environment.NewLine}");
    }

    private static bool ValidateSql(string sql)
    {
        if (string.IsNullOrWhiteSpace(sql)) return false;

        foreach (string line in sql.Split(
            Environment.NewLine,
            StringSplitOptions.RemoveEmptyEntries))
        {
            string t = line.Trim();
            if (t.Length == 0 || t.StartsWith("--")) continue;
            return true;
        }

        return false;
    }

    private static string GetFileSize(string path)
    {
        FileInfo fi = new(path);
        if (fi.Length < 1024) return $"{fi.Length} B";
        if (fi.Length < 1024 * 1024) return $"{fi.Length / 1024.0:F1} KB";
        return $"{fi.Length / 1024.0 / 1024.0:F2} MB";
    }

    private static void PrintSummary(int total, long elapsed)
    {
        Console.WriteLine();
        Console.WriteLine("----------------------------------------");
        Console.WriteLine($"Files : {total}");
        Console.WriteLine($"Time  : {elapsed} ms");
        Console.WriteLine("----------------------------------------");
    }
}
