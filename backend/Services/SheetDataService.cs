using System.Data;
using Microsoft.Data.SqlClient;
using Dapper;
using ExcelClone.Api.Models;
using System.Text;
using System.Security.Cryptography;

namespace ExcelClone.Api.Services;

/// <summary>
/// Interface for sheet data operations
/// </summary>
public interface ISheetDataService
{
    Task<SheetQueryResponse> QuerySheetDataAsync(string sheetId, SheetQueryRequest request);
    Task<SheetSaveResponse> SaveSheetDataAsync(string sheetId, SheetSaveRequest request);
    Task<List<LookupValue>> GetLookupValuesAsync(string sheetId, string columnName, string? search = null, int limit = 100);
    Task<Models.ValidationResult> ValidateDataAsync(string sheetId, List<Dictionary<string, object?>> rows);
    Task<SheetStats> GetSheetStatsAsync(string sheetId);
}

/// <summary>
/// Service for managing sheet data operations
/// </summary>
public class SheetDataService : ISheetDataService
{
    private readonly string _connectionString;
    private readonly ILogger<SheetDataService> _logger;

    public SheetDataService(IConfiguration configuration, ILogger<SheetDataService> logger)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection") 
            ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        _logger = logger;
    }

    /// <summary>
    /// Query sheet data with filtering, sorting, and pagination
    /// </summary>
    public async Task<SheetQueryResponse> QuerySheetDataAsync(string sheetId, SheetQueryRequest request)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        
        try
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            // Get sheet definition to determine table and columns
            var sheetDef = await GetSheetDefinitionAsync(connection, sheetId);
            if (sheetDef == null)
            {
                throw new ArgumentException($"Sheet '{sheetId}' not found");
            }

            // Build the query
            var queryBuilder = new QueryBuilder(sheetDef, request);
            var countQuery = queryBuilder.BuildCountQuery();
            var dataQuery = queryBuilder.BuildDataQuery();

            // Execute count query
            var total = await connection.QuerySingleAsync<int>(countQuery, queryBuilder.Parameters);

            // Execute data query
            var rows = await connection.QueryAsync(dataQuery, queryBuilder.Parameters);

            // Convert to response format with row hashes
            var responseRows = rows.Select(row =>
            {
                var rowDict = ((IDictionary<string, object>)row).ToDictionary(
                    kvp => kvp.Key,
                    kvp => kvp.Value == DBNull.Value ? null : kvp.Value
                );

                // Add row hash and version
                rowDict["_rowHash"] = ComputeRowHash(rowDict, sheetDef.Key);
                rowDict["_rowVersion"] = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();

                return rowDict;
            }).ToList();

            stopwatch.Stop();

            return new SheetQueryResponse
            {
                Rows = responseRows,
                Total = total,
                Page = request.Page,
                PageSize = request.PageSize,
                Metadata = new QueryMetadata
                {
                    ExecutionTimeMs = (int)stopwatch.ElapsedMilliseconds,
                    SqlQuery = dataQuery,
                    Parameters = new Dictionary<string, object>() // TODO: Convert DynamicParameters to Dictionary if needed
                }
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error querying sheet data for sheet {SheetId}", sheetId);
            throw;
        }
    }

    /// <summary>
    /// Save sheet data using staging tables and merge procedures
    /// </summary>
    public async Task<SheetSaveResponse> SaveSheetDataAsync(string sheetId, SheetSaveRequest request)
    {
        try
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            using var transaction = (SqlTransaction)await connection.BeginTransactionAsync();

            try
            {
                // Get sheet definition
                var sheetDef = await GetSheetDefinitionAsync(connection, sheetId, transaction);
                if (sheetDef == null)
                {
                    throw new ArgumentException($"Sheet '{sheetId}' not found");
                }

                var results = new SaveResults();

                // Process operations in order: Deletes, Updates, Inserts
                if (request.Deletes.Any())
                {
                    results.Deletes = await ProcessDeletesAsync(connection, transaction, sheetDef, request.Deletes);
                }

                if (request.Updates.Any())
                {
                    results.Updates = await ProcessUpdatesAsync(connection, transaction, sheetDef, request.Updates);
                }

                if (request.Inserts.Any())
                {
                    results.Inserts = await ProcessInsertsAsync(connection, transaction, sheetDef, request.Inserts);
                }

                // Log audit trail
                await LogAuditTrailAsync(connection, transaction, sheetId, request, results);

                await transaction.CommitAsync();

                return new SheetSaveResponse
                {
                    Ok = true,
                    Results = results
                };
            }
            catch (Exception)
            {
                await transaction.RollbackAsync();
                throw;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error saving sheet data for sheet {SheetId}", sheetId);
            return new SheetSaveResponse
            {
                Ok = false,
                ErrorMessage = ex.Message
            };
        }
    }

    private async Task<SheetDefinition?> GetSheetDefinitionAsync(SqlConnection connection, string sheetId, SqlTransaction? transaction = null)
    {
        // This would typically come from a configuration table
        // For now, return a sample definition for FinancialData
        if (sheetId == "financial-data")
        {
            return new SheetDefinition
            {
                Id = "financial-data",
                Name = "Financial Data",
                Table = "FinancialData",
                Key = new List<string> { "InternalOrder", "ItemID" },
                Columns = new List<ColumnDefinition>
                {
                    new() { Name = "InternalOrder", DisplayName = "Internal Order", Type = "string", IsKey = true, IsRequired = true },
                    new() { Name = "ItemID", DisplayName = "Item ID", Type = "int", IsKey = true, IsRequired = true },
                    new() { Name = "Amount", DisplayName = "Amount", Type = "decimal", IsEditable = true, IsRequired = true },
                    new() { Name = "CostCenter", DisplayName = "Cost Center", Type = "string", IsEditable = true },
                    new() { Name = "ModifiedBy", DisplayName = "Modified By", Type = "string", IsEditable = false },
                    new() { Name = "ModifiedAt", DisplayName = "Modified At", Type = "datetime", IsEditable = false }
                }
            };
        }

        return null;
    }

    private async Task<List<OperationResult>> ProcessInsertsAsync(
        SqlConnection connection, 
        SqlTransaction transaction, 
        SheetDefinition sheetDef, 
        List<RowInsert> inserts)
    {
        var results = new List<OperationResult>();

        if (!inserts.Any()) return results;

        // Create staging table data
        var stagingData = CreateStagingDataTable(sheetDef, inserts.Select(i => i.Data));

        // Bulk insert to staging
        var stagingTableName = $"stage.{sheetDef.Table}";
        await BulkInsertToStagingAsync(connection, transaction, stagingTableName, stagingData);

        // Execute merge procedure
        var mergeProc = $"dbo.usp_merge_{sheetDef.Table}_from_stage";
        await connection.ExecuteAsync(mergeProc, transaction: transaction, commandType: CommandType.StoredProcedure);

        // For simplicity, mark all as merged (in real implementation, check actual results)
        foreach (var insert in inserts)
        {
            var key = sheetDef.Key.ToDictionary(k => k, k => insert.Data.GetValueOrDefault(k));
            results.Add(new OperationResult
            {
                Key = key,
                Status = "merged",
                ClientId = insert.ClientId
            });
        }

        return results;
    }

    private async Task<List<OperationResult>> ProcessUpdatesAsync(
        SqlConnection connection, 
        SqlTransaction transaction, 
        SheetDefinition sheetDef, 
        List<RowUpdate> updates)
    {
        var results = new List<OperationResult>();

        foreach (var update in updates)
        {
            // Check current hash for optimistic concurrency
            var currentRow = await GetCurrentRowAsync(connection, transaction, sheetDef, update.Key);
            if (currentRow == null)
            {
                results.Add(new OperationResult
                {
                    Key = update.Key,
                    Status = "missing",
                    Reason = "Row not found",
                    ClientId = update.ClientId
                });
                continue;
            }

            var currentHash = ComputeRowHash(currentRow, sheetDef.Key);
            if (currentHash != update.BeforeHash)
            {
                results.Add(new OperationResult
                {
                    Key = update.Key,
                    Status = "conflict",
                    Reason = "Row was modified by another user",
                    ClientId = update.ClientId,
                    CurrentData = currentRow,
                    CurrentHash = currentHash
                });
                continue;
            }

            // Perform update
            var updateSql = BuildUpdateSql(sheetDef, update);
            var parameters = BuildUpdateParameters(update);
            
            var rowsAffected = await connection.ExecuteAsync(updateSql, parameters, transaction);
            
            results.Add(new OperationResult
            {
                Key = update.Key,
                Status = rowsAffected > 0 ? "merged" : "error",
                ClientId = update.ClientId
            });
        }

        return results;
    }

    private async Task<List<OperationResult>> ProcessDeletesAsync(
        SqlConnection connection, 
        SqlTransaction transaction, 
        SheetDefinition sheetDef, 
        List<RowDelete> deletes)
    {
        var results = new List<OperationResult>();

        foreach (var delete in deletes)
        {
            // Check current hash for optimistic concurrency
            var currentRow = await GetCurrentRowAsync(connection, transaction, sheetDef, delete.Key);
            if (currentRow == null)
            {
                results.Add(new OperationResult
                {
                    Key = delete.Key,
                    Status = "missing",
                    Reason = "Row not found",
                    ClientId = delete.ClientId
                });
                continue;
            }

            var currentHash = ComputeRowHash(currentRow, sheetDef.Key);
            if (currentHash != delete.BeforeHash)
            {
                results.Add(new OperationResult
                {
                    Key = delete.Key,
                    Status = "conflict",
                    Reason = "Row was modified by another user",
                    ClientId = delete.ClientId,
                    CurrentData = currentRow,
                    CurrentHash = currentHash
                });
                continue;
            }

            // Perform delete
            var deleteSql = BuildDeleteSql(sheetDef, delete);
            var parameters = BuildKeyParameters(delete.Key);
            
            var rowsAffected = await connection.ExecuteAsync(deleteSql, parameters, transaction);
            
            results.Add(new OperationResult
            {
                Key = delete.Key,
                Status = rowsAffected > 0 ? "deleted" : "error",
                ClientId = delete.ClientId
            });
        }

        return results;
    }

    private string ComputeRowHash(Dictionary<string, object?> row, List<string> keyColumns)
    {
        // Create hash from data columns (excluding metadata)
        var dataForHash = row.Where(kvp => 
            !kvp.Key.StartsWith("_") && 
            kvp.Key != "ModifiedBy" && 
            kvp.Key != "ModifiedAt")
            .OrderBy(kvp => kvp.Key)
            .Select(kvp => $"{kvp.Key}|{kvp.Value}")
            .ToList();

        var hashInput = string.Join("|", dataForHash);
        using var sha256 = SHA256.Create();
        var hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(hashInput));
        return "0x" + Convert.ToHexString(hashBytes);
    }

    private DataTable CreateStagingDataTable(SheetDefinition sheetDef, IEnumerable<Dictionary<string, object?>> rows)
    {
        var table = new DataTable();
        
        // Add columns based on sheet definition
        foreach (var col in sheetDef.Columns.Where(c => c.IsEditable || c.IsKey))
        {
            var dataType = col.Type switch
            {
                "int" => typeof(int),
                "decimal" => typeof(decimal),
                "datetime" => typeof(DateTime),
                "bool" => typeof(bool),
                _ => typeof(string)
            };
            table.Columns.Add(col.Name, dataType);
        }

        // Add metadata columns
        table.Columns.Add("ModifiedBy", typeof(string));
        table.Columns.Add("ModifiedAt", typeof(DateTime));

        // Add rows
        foreach (var rowData in rows)
        {
            var row = table.NewRow();
            foreach (var col in sheetDef.Columns.Where(c => c.IsEditable || c.IsKey))
            {
                if (rowData.TryGetValue(col.Name, out var value) && value != null)
                {
                    row[col.Name] = value;
                }
            }
            row["ModifiedBy"] = "system"; // TODO: Get from current user context
            row["ModifiedAt"] = DateTime.UtcNow;
            table.Rows.Add(row);
        }

        return table;
    }

    private async Task BulkInsertToStagingAsync(SqlConnection connection, SqlTransaction transaction, string tableName, DataTable data)
    {
        // Clear staging table first
        await connection.ExecuteAsync($"TRUNCATE TABLE {tableName}", transaction: transaction);

        // Bulk insert
        using var bulk = new SqlBulkCopy(connection, SqlBulkCopyOptions.TableLock, transaction)
        {
            DestinationTableName = tableName,
            BatchSize = 5000
        };

        await bulk.WriteToServerAsync(data);
    }

    private async Task<Dictionary<string, object?>?> GetCurrentRowAsync(
        SqlConnection connection, 
        SqlTransaction transaction, 
        SheetDefinition sheetDef, 
        Dictionary<string, object?> key)
    {
        var whereClause = string.Join(" AND ", sheetDef.Key.Select(k => $"{k} = @{k}"));
        var sql = $"SELECT * FROM dbo.{sheetDef.Table} WHERE {whereClause}";
        
        var result = await connection.QueryFirstOrDefaultAsync(sql, key, transaction);
        if (result == null) return null;

        return ((IDictionary<string, object>)result).ToDictionary(
            kvp => kvp.Key,
            kvp => kvp.Value == DBNull.Value ? null : kvp.Value
        );
    }

    private string BuildUpdateSql(SheetDefinition sheetDef, RowUpdate update)
    {
        var setClause = string.Join(", ", update.After.Keys.Select(k => $"{k} = @new_{k}"));
        var whereClause = string.Join(" AND ", sheetDef.Key.Select(k => $"{k} = @key_{k}"));
        
        return $@"
            UPDATE dbo.{sheetDef.Table} 
            SET {setClause}, ModifiedBy = @ModifiedBy, ModifiedAt = @ModifiedAt
            WHERE {whereClause}";
    }

    private string BuildDeleteSql(SheetDefinition sheetDef, RowDelete delete)
    {
        var whereClause = string.Join(" AND ", sheetDef.Key.Select(k => $"{k} = @key_{k}"));
        return $"DELETE FROM dbo.{sheetDef.Table} WHERE {whereClause}";
    }

    private DynamicParameters BuildUpdateParameters(RowUpdate update)
    {
        var parameters = new DynamicParameters();
        
        // Add new values
        foreach (var kvp in update.After)
        {
            parameters.Add($"new_{kvp.Key}", kvp.Value);
        }
        
        // Add key values
        foreach (var kvp in update.Key)
        {
            parameters.Add($"key_{kvp.Key}", kvp.Value);
        }
        
        parameters.Add("ModifiedBy", "system"); // TODO: Get from current user context
        parameters.Add("ModifiedAt", DateTime.UtcNow);
        
        return parameters;
    }

    private DynamicParameters BuildKeyParameters(Dictionary<string, object?> key)
    {
        var parameters = new DynamicParameters();
        foreach (var kvp in key)
        {
            parameters.Add($"key_{kvp.Key}", kvp.Value);
        }
        return parameters;
    }

    private async Task LogAuditTrailAsync(
        SqlConnection connection, 
        SqlTransaction transaction, 
        string sheetId, 
        SheetSaveRequest request, 
        SaveResults results)
    {
        // Log audit information
        var auditSql = @"
            INSERT INTO audit.OperationLog (SheetId, SessionId, OperationType, RowCount, ProcessedAt, ModifiedBy)
            VALUES (@SheetId, @SessionId, @OperationType, @RowCount, @ProcessedAt, @ModifiedBy)";

        if (request.Inserts.Any())
        {
            await connection.ExecuteAsync(auditSql, new
            {
                SheetId = sheetId,
                SessionId = request.ClientSessionId,
                OperationType = "INSERT",
                RowCount = request.Inserts.Count,
                ProcessedAt = DateTime.UtcNow,
                ModifiedBy = "system"
            }, transaction);
        }

        if (request.Updates.Any())
        {
            await connection.ExecuteAsync(auditSql, new
            {
                SheetId = sheetId,
                SessionId = request.ClientSessionId,
                OperationType = "UPDATE",
                RowCount = request.Updates.Count,
                ProcessedAt = DateTime.UtcNow,
                ModifiedBy = "system"
            }, transaction);
        }

        if (request.Deletes.Any())
        {
            await connection.ExecuteAsync(auditSql, new
            {
                SheetId = sheetId,
                SessionId = request.ClientSessionId,
                OperationType = "DELETE",
                RowCount = request.Deletes.Count,
                ProcessedAt = DateTime.UtcNow,
                ModifiedBy = "system"
            }, transaction);
        }
    }

    /// <summary>
    /// Get lookup values for a column
    /// </summary>
    public async Task<List<LookupValue>> GetLookupValuesAsync(string sheetId, string columnName, string? search = null, int limit = 100)
    {
        // For now, return mock data
        // In production, this would query the lookup table defined in the sheet definition
        await Task.Delay(10); // Simulate async operation
        
        var mockValues = new List<LookupValue>
        {
            new() { Value = "CC0001", Display = "Engineering", Description = "Engineering Department" },
            new() { Value = "CC0002", Display = "Marketing", Description = "Marketing Department" },
            new() { Value = "CC0003", Display = "Sales", Description = "Sales Department" },
            new() { Value = "CC0004", Display = "HR", Description = "Human Resources" }
        };

        if (!string.IsNullOrWhiteSpace(search))
        {
            mockValues = mockValues
                .Where(v => v.Display.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                           v.Value.Contains(search, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        return mockValues.Take(limit).ToList();
    }

    /// <summary>
    /// Validate data against sheet definition
    /// </summary>
    public async Task<Models.ValidationResult> ValidateDataAsync(string sheetId, List<Dictionary<string, object?>> rows)
    {
        await Task.Delay(10); // Simulate async operation
        
        var result = new Models.ValidationResult();
        
        // Mock validation - in production this would validate against the sheet definition
        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            
            // Example validation: check required fields
            if (row.ContainsKey("InternalOrder") && (row["InternalOrder"] == null || string.IsNullOrWhiteSpace(row["InternalOrder"]?.ToString())))
            {
                result.Errors.Add(new ValidationError
                {
                    RowIndex = i,
                    ColumnName = "InternalOrder",
                    Message = "Internal Order is required",
                    Value = row["InternalOrder"]?.ToString()
                });
            }
            
            // Example validation: check data types
            if (row.ContainsKey("Amount") && row["Amount"] != null)
            {
                if (!decimal.TryParse(row["Amount"]?.ToString(), out _))
                {
                    result.Errors.Add(new ValidationError
                    {
                        RowIndex = i,
                        ColumnName = "Amount",
                        Message = "Amount must be a valid decimal number",
                        Value = row["Amount"]?.ToString()
                    });
                }
            }
        }
        
        return result;
    }

    /// <summary>
    /// Get statistics about a sheet
    /// </summary>
    public async Task<SheetStats> GetSheetStatsAsync(string sheetId)
    {
        await Task.Delay(10); // Simulate async operation
        
        // Mock stats - in production this would query the actual table
        return new SheetStats
        {
            RowCount = 12500,
            LastModified = DateTime.UtcNow.AddHours(-2),
            LastModifiedBy = "user@example.com",
            ColumnStats = new Dictionary<string, object>
            {
                ["Amount"] = new { Min = 0.00m, Max = 999999.99m, Average = 12345.67m },
                ["Category"] = new { OPEX = 5000, CAPEX = 3000, Revenue = 4000, Other = 500 }
            }
        };
    }
}