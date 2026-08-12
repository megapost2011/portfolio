using System;
using System.Collections.Generic;
using System.Data;

namespace BusinessManagementSuite.Data;

/// <summary>
/// 全リポジトリ共通の基底クラス（SQLite / DataTable ベース）
/// </summary>
public abstract class RepositoryBase<T>
    where T : class, new()
{
    //---------------------------------------------------------
    // Table / Key
    //---------------------------------------------------------

    protected abstract string TableName { get; }

    protected virtual string KeyName => "Id";

    //---------------------------------------------------------
    // Map
    //---------------------------------------------------------

    protected abstract T Map(DataRow row);

    //---------------------------------------------------------
    // Exists (by id)
    //---------------------------------------------------------

    public virtual bool Exists(long id)
    {
        string sql =
            $"SELECT COUNT(*) FROM {TableName} " +
            $"WHERE {KeyName}=@Id;";

        object? result = Database.ExecuteScalar(
            sql,
            new() { ["@Id"] = id });

        return Convert.ToInt32(result) > 0;
    }

    //---------------------------------------------------------
    // Count
    //---------------------------------------------------------

    public virtual int Count()
    {
        object? result = Database.ExecuteScalar(
            $"SELECT COUNT(*) FROM {TableName};");

        return Convert.ToInt32(result);
    }

    public virtual int Count(string whereClause)
    {
        object? result = Database.ExecuteScalar(
            $"SELECT COUNT(*) FROM {TableName} WHERE {whereClause};");

        return Convert.ToInt32(result);
    }

    //---------------------------------------------------------
    // GetById
    //---------------------------------------------------------

    public virtual T? GetById(long id)
    {
        string sql =
            $"SELECT * FROM {TableName} WHERE {KeyName}=@Id LIMIT 1;";

        DataTable table = Database.ExecuteReader(
            sql,
            new() { ["@Id"] = id });

        if (table.Rows.Count == 0) return null;

        return Map(table.Rows[0]);
    }

    //---------------------------------------------------------
    // GetAll
    //---------------------------------------------------------

    public virtual List<T> GetAll()
        => Query($"SELECT * FROM {TableName};");

    //---------------------------------------------------------
    // Find (raw WHERE clause – caller is responsible for safety)
    //---------------------------------------------------------

    public virtual List<T> Find(string whereClause)
        => Query($"SELECT * FROM {TableName} WHERE {whereClause};");

    //---------------------------------------------------------
    // Find with parameters
    //---------------------------------------------------------

    protected List<T> Find(
        string whereClause,
        Dictionary<string, object?> parameters)
        => Query(
            $"SELECT * FROM {TableName} WHERE {whereClause};",
            parameters);

    //---------------------------------------------------------
    // FirstOrDefault
    //---------------------------------------------------------

    public virtual T? FirstOrDefault(string whereClause)
    {
        DataTable table = Database.ExecuteReader(
            $"SELECT * FROM {TableName} WHERE {whereClause} LIMIT 1;");

        if (table.Rows.Count == 0) return null;

        return Map(table.Rows[0]);
    }

    //---------------------------------------------------------
    // Any
    //---------------------------------------------------------

    public virtual bool Any(string whereClause)
        => Count(whereClause) > 0;

    //---------------------------------------------------------
    // Paging
    //---------------------------------------------------------

    public virtual List<T> GetPage(int page, int pageSize)
    {
        int offset = (page - 1) * pageSize;

        return Query(
            $"SELECT * FROM {TableName} LIMIT {pageSize} OFFSET {offset};");
    }

    //---------------------------------------------------------
    // OrderBy
    //---------------------------------------------------------

    public virtual List<T> OrderBy(
        string column,
        bool ascending = true)
    {
        string order = ascending ? "ASC" : "DESC";

        return Query(
            $"SELECT * FROM {TableName} ORDER BY {column} {order};");
    }

    //---------------------------------------------------------
    // Soft Delete
    //---------------------------------------------------------

    public virtual bool Delete(long id)
    {
        int rows = Database.Execute(
            $"UPDATE {TableName} SET IsDeleted=1 WHERE {KeyName}=@Id;",
            new() { ["@Id"] = id });

        return rows > 0;
    }

    //---------------------------------------------------------
    // Restore
    //---------------------------------------------------------

    public virtual bool Restore(long id)
    {
        int rows = Database.Execute(
            $"UPDATE {TableName} SET IsDeleted=0 WHERE {KeyName}=@Id;",
            new() { ["@Id"] = id });

        return rows > 0;
    }

    //---------------------------------------------------------
    // Physical Delete
    //---------------------------------------------------------

    public virtual bool DeletePhysical(long id)
    {
        int rows = Database.Execute(
            $"DELETE FROM {TableName} WHERE {KeyName}=@Id;",
            new() { ["@Id"] = id });

        return rows > 0;
    }

    //---------------------------------------------------------
    // Clear Table
    //---------------------------------------------------------

    public virtual void Clear()
        => Database.Execute($"DELETE FROM {TableName};");

    //---------------------------------------------------------
    // Table Exists
    //---------------------------------------------------------

    public bool TableExists()
    {
        object? result = Database.ExecuteScalar(
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@Table;",
            new() { ["@Table"] = TableName });

        return Convert.ToInt32(result) > 0;
    }

    //---------------------------------------------------------
    // Transaction helpers
    //---------------------------------------------------------

    public void BeginTransaction()
        => Database.Execute("BEGIN TRANSACTION;");

    public void Commit()
        => Database.Execute("COMMIT;");

    public void Rollback()
        => Database.Execute("ROLLBACK;");

    //---------------------------------------------------------
    // Internal query helpers
    //---------------------------------------------------------

    protected List<T> Query(string sql,
        Dictionary<string, object?>? parameters = null)
    {
        DataTable table = Database.ExecuteReader(sql, parameters);
        List<T> list = new(table.Rows.Count);

        foreach (DataRow row in table.Rows)
            list.Add(Map(row));

        return list;
    }

    //---------------------------------------------------------
    // Log
    //---------------------------------------------------------

    protected virtual void Log(string message)
        => Console.WriteLine($"[{TableName}] {message}");

    public override string ToString() => TableName;
}
