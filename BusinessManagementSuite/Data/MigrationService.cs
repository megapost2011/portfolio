using System;
using System.Security.Cryptography;
using System.Text;

namespace BusinessManagementSuite.Data;

public static class MigrationService
{
    public static int CurrentVersion => Database.GetUserVersion();

    //---------------------------------------------------------
    // Update Version
    //---------------------------------------------------------

    public static void UpdateVersion()
    {
        int latest = GetLatestVersion();
        Database.SetUserVersion(latest);
    }

    //---------------------------------------------------------
    // Latest Version
    //---------------------------------------------------------

    private static int GetLatestVersion()
    {
        // SchemaVersions テーブルが存在しない場合は 0 を返す
        try
        {
            object? result = Database.ExecuteScalar(
                "SELECT IFNULL(MAX(Version),0) FROM SchemaVersions;");

            return Convert.ToInt32(result);
        }
        catch
        {
            return 0;
        }
    }

    //---------------------------------------------------------
    // Is Executed
    //---------------------------------------------------------

    public static bool IsExecuted(string fileName)
    {
        try
        {
            object? result = Database.ExecuteScalar(
                "SELECT COUNT(*) FROM SchemaVersions WHERE FileName=@FileName;",
                new() { ["@FileName"] = fileName });

            return Convert.ToInt32(result) > 0;
        }
        catch
        {
            return false;
        }
    }

    //---------------------------------------------------------
    // Register
    //---------------------------------------------------------

    public static void Register(
        int version,
        string fileName,
        string checksum,
        long elapsed,
        bool success,
        string error)
    {
        try
        {
            Database.Execute(
                """
                INSERT INTO SchemaVersions
                (Version, FileName, Checksum, ExecutedAt, ExecutionTime, Success, ErrorMessage)
                VALUES
                (@Version, @FileName, @Checksum,
                 datetime('now','localtime'), @ExecutionTime, @Success, @Error);
                """,
                new()
                {
                    ["@Version"]       = version,
                    ["@FileName"]      = fileName,
                    ["@Checksum"]      = checksum,
                    ["@ExecutionTime"] = elapsed,
                    ["@Success"]       = success ? 1 : 0,
                    ["@Error"]         = error
                });
        }
        catch
        {
            // SchemaVersions テーブルが未作成の場合は無視
        }
    }

    //---------------------------------------------------------
    // SHA256
    //---------------------------------------------------------

    public static string GetChecksum(string text)
    {
        byte[] bytes = SHA256.HashData(Encoding.UTF8.GetBytes(text));
        return Convert.ToHexString(bytes);
    }
}
