using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Microsoft.Data.Sqlite;

namespace BusinessManagementSuite.Data;

public static class Database
{
    //---------------------------------------------------------
    // Database File
    //---------------------------------------------------------

    private static readonly string DatabaseFolder =
        Path.Combine(
            Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData),
            "BusinessManagementSuite");

    internal static readonly string DatabasePath =
        Path.Combine(DatabaseFolder, "BusinessManagementSuite.db");

    //---------------------------------------------------------
    // Connection String
    //---------------------------------------------------------

    private static readonly string ConnectionString =
        new SqliteConnectionStringBuilder
        {
            DataSource = DatabasePath,
            Mode = SqliteOpenMode.ReadWriteCreate,
            Cache = SqliteCacheMode.Shared,
            Pooling = true
        }.ToString();

    //---------------------------------------------------------
    // Initialize
    //---------------------------------------------------------

    public static void Initialize()
    {
        if (!Directory.Exists(DatabaseFolder))
            Directory.CreateDirectory(DatabaseFolder);

        using SqliteConnection conn = Open();
        using SqliteCommand cmd = conn.CreateCommand();

        cmd.CommandText =
        """
        PRAGMA journal_mode=WAL;
        PRAGMA foreign_keys=ON;
        PRAGMA synchronous=NORMAL;
        PRAGMA busy_timeout=5000;
        """;

        cmd.ExecuteNonQuery();
    }

    //---------------------------------------------------------
    // Connection
    //---------------------------------------------------------

    public static SqliteConnection GetConnection()
        => new SqliteConnection(ConnectionString);

    public static SqliteConnection Open()
    {
        SqliteConnection conn = GetConnection();
        conn.Open();
        return conn;
    }

    //---------------------------------------------------------
    // Execute
    //---------------------------------------------------------

    public static int Execute(
        string sql,
        Dictionary<string, object?>? parameters = null)
    {
        using SqliteConnection conn = Open();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        AddParameters(cmd, parameters);
        return cmd.ExecuteNonQuery();
    }

    //---------------------------------------------------------
    // ExecuteScalar
    //---------------------------------------------------------

    public static object? ExecuteScalar(
        string sql,
        Dictionary<string, object?>? parameters = null)
    {
        using SqliteConnection conn = Open();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        AddParameters(cmd, parameters);
        return cmd.ExecuteScalar();
    }

    //---------------------------------------------------------
    // ExecuteReader - returns DataTable
    //---------------------------------------------------------

    public static DataTable ExecuteReader(
        string sql,
        Dictionary<string, object?>? parameters = null)
    {
        using SqliteConnection conn = Open();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        AddParameters(cmd, parameters);
        using SqliteDataReader reader = cmd.ExecuteReader();
        DataTable table = new();
        table.Load(reader);
        return table;
    }

    //---------------------------------------------------------
    // Transaction
    //---------------------------------------------------------

    public static void ExecuteInTransaction(
        Action<SqliteConnection, SqliteTransaction> action)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tran = conn.BeginTransaction();
        try
        {
            action(conn, tran);
            tran.Commit();
        }
        catch
        {
            tran.Rollback();
            throw;
        }
    }

    //---------------------------------------------------------
    // Parameters
    //---------------------------------------------------------

    private static void AddParameters(
        SqliteCommand command,
        Dictionary<string, object?>? parameters)
    {
        if (parameters == null) return;

        foreach (KeyValuePair<string, object?> p in parameters)
        {
            command.Parameters.AddWithValue(
                p.Key,
                p.Value ?? DBNull.Value);
        }
    }

    //---------------------------------------------------------
    // Utility
    //---------------------------------------------------------

    public static void Backup(string destinationFile)
    {
        if (string.IsNullOrWhiteSpace(destinationFile))
            throw new ArgumentNullException(nameof(destinationFile));

        Directory.CreateDirectory(
            Path.GetDirectoryName(destinationFile)!);

        File.Copy(DatabasePath, destinationFile, true);
    }

    public static void Restore(string sourceFile)
    {
        if (!File.Exists(sourceFile))
            throw new FileNotFoundException(sourceFile);

        File.Copy(sourceFile, DatabasePath, true);
    }

    public static void Vacuum()
        => Execute("VACUUM;");

    public static void Analyze()
        => Execute("ANALYZE;");

    public static bool IntegrityCheck()
    {
        object? result = ExecuteScalar("PRAGMA integrity_check;");
        return result?.ToString() == "ok";
    }

    public static void Optimize()
        => Execute("PRAGMA optimize;");

    public static int GetUserVersion()
    {
        object? value = ExecuteScalar("PRAGMA user_version;");
        return Convert.ToInt32(value);
    }

    public static void SetUserVersion(int version)
        => Execute($"PRAGMA user_version={version};");

    public static long GetDatabaseSize()
    {
        if (!File.Exists(DatabasePath)) return 0;
        return new FileInfo(DatabasePath).Length;
    }

    public static bool Exists()
        => File.Exists(DatabasePath);

    public static void DeleteDatabase()
    {
        if (File.Exists(DatabasePath))
            File.Delete(DatabasePath);
    }
}
