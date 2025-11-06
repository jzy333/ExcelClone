using ExcelClone.Api.Models;
using Dapper;
using System.Text;

namespace ExcelClone.Api.Services;

/// <summary>
/// Helper class for building dynamic SQL queries based on sheet definitions and filter criteria
/// </summary>
public class QueryBuilder
{
    private readonly SheetDefinition _sheetDef;
    private readonly SheetQueryRequest _request;
    private readonly DynamicParameters _parameters;

    public QueryBuilder(SheetDefinition sheetDef, SheetQueryRequest request)
    {
        _sheetDef = sheetDef;
        _request = request;
        _parameters = new DynamicParameters();
    }

    public DynamicParameters Parameters => _parameters;

    /// <summary>
    /// Build count query for pagination
    /// </summary>
    public string BuildCountQuery()
    {
        var whereClause = BuildWhereClause();
        
        return $@"
            SELECT COUNT(*)
            FROM dbo.{_sheetDef.Table}
            {whereClause}";
    }

    /// <summary>
    /// Build data query with filtering, sorting, and pagination
    /// </summary>
    public string BuildDataQuery()
    {
        var selectClause = BuildSelectClause();
        var whereClause = BuildWhereClause();
        var orderByClause = BuildOrderByClause();
        var offsetClause = BuildOffsetClause();

        return $@"
            {selectClause}
            FROM dbo.{_sheetDef.Table}
            {whereClause}
            {orderByClause}
            {offsetClause}";
    }

    private string BuildSelectClause()
    {
        var columns = _sheetDef.Columns.Select(c => c.Name);
        return $"SELECT {string.Join(", ", columns)}";
    }

    private string BuildWhereClause()
    {
        if (!_request.Filters.Any())
            return string.Empty;

        var conditions = new List<string>();
        var paramIndex = 0;

        foreach (var filter in _request.Filters)
        {
            var condition = BuildFilterCondition(filter, ref paramIndex);
            if (!string.IsNullOrEmpty(condition))
            {
                conditions.Add(condition);
            }
        }

        return conditions.Any() ? $"WHERE {string.Join(" AND ", conditions)}" : string.Empty;
    }

    private string BuildFilterCondition(FilterCriteria filter, ref int paramIndex)
    {
        var paramName = $"filter_{paramIndex++}";
        var column = filter.Column;

        return filter.Operator.ToLower() switch
        {
            "eq" or "=" => AddParameterAndReturn(paramName, filter.Value, $"{column} = @{paramName}"),
            "ne" or "!=" => AddParameterAndReturn(paramName, filter.Value, $"{column} != @{paramName}"),
            "gt" or ">" => AddParameterAndReturn(paramName, filter.Value, $"{column} > @{paramName}"),
            "lt" or "<" => AddParameterAndReturn(paramName, filter.Value, $"{column} < @{paramName}"),
            "gte" or ">=" => AddParameterAndReturn(paramName, filter.Value, $"{column} >= @{paramName}"),
            "lte" or "<=" => AddParameterAndReturn(paramName, filter.Value, $"{column} <= @{paramName}"),
            "contains" => AddParameterAndReturn(paramName, $"%{filter.Value}%", $"{column} LIKE @{paramName}"),
            "startswith" => AddParameterAndReturn(paramName, $"{filter.Value}%", $"{column} LIKE @{paramName}"),
            "endswith" => AddParameterAndReturn(paramName, $"%{filter.Value}", $"{column} LIKE @{paramName}"),
            "in" => BuildInCondition(column, filter.Values, ref paramIndex),
            "notin" => BuildNotInCondition(column, filter.Values, ref paramIndex),
            "isnull" => $"{column} IS NULL",
            "isnotnull" => $"{column} IS NOT NULL",
            _ => string.Empty
        };
    }

    private string AddParameterAndReturn(string paramName, object? value, string condition)
    {
        _parameters.Add(paramName, value);
        return condition;
    }

    private string BuildInCondition(string column, List<object>? values, ref int paramIndex)
    {
        if (values == null || !values.Any())
            return string.Empty;

        var paramNames = new List<string>();
        foreach (var value in values)
        {
            var paramName = $"filter_{paramIndex++}";
            _parameters.Add(paramName, value);
            paramNames.Add($"@{paramName}");
        }

        return $"{column} IN ({string.Join(", ", paramNames)})";
    }

    private string BuildNotInCondition(string column, List<object>? values, ref int paramIndex)
    {
        if (values == null || !values.Any())
            return string.Empty;

        var paramNames = new List<string>();
        foreach (var value in values)
        {
            var paramName = $"filter_{paramIndex++}";
            _parameters.Add(paramName, value);
            paramNames.Add($"@{paramName}");
        }

        return $"{column} NOT IN ({string.Join(", ", paramNames)})";
    }

    private string BuildOrderByClause()
    {
        if (!_request.Sorts.Any())
        {
            // Default sort by primary key
            var keyColumns = _sheetDef.Key.Select(k => $"{k} ASC");
            return $"ORDER BY {string.Join(", ", keyColumns)}";
        }

        var sortExpressions = _request.Sorts.Select(s => 
        {
            var direction = s.Direction.ToUpper() == "DESC" ? "DESC" : "ASC";
            return $"{s.Column} {direction}";
        });

        return $"ORDER BY {string.Join(", ", sortExpressions)}";
    }

    private string BuildOffsetClause()
    {
        var offset = (_request.Page - 1) * _request.PageSize;
        _parameters.Add("offset", offset);
        _parameters.Add("pageSize", _request.PageSize);
        
        return "OFFSET @offset ROWS FETCH NEXT @pageSize ROWS ONLY";
    }

    /// <summary>
    /// Build query for lookup values
    /// </summary>
    public static string BuildLookupQuery(LookupDefinition lookup, string? searchTerm = null)
    {
        var sql = new StringBuilder();
        sql.AppendLine($"SELECT DISTINCT {lookup.ValueColumn} as Value, {lookup.DisplayColumn} as Display");
        sql.AppendLine($"FROM {lookup.Table}");

        var conditions = new List<string>();

        if (!string.IsNullOrEmpty(lookup.FilterColumn) && lookup.FilterValue != null)
        {
            conditions.Add($"{lookup.FilterColumn} = @FilterValue");
        }

        if (!string.IsNullOrEmpty(searchTerm))
        {
            conditions.Add($"{lookup.DisplayColumn} LIKE @SearchTerm");
        }

        if (conditions.Any())
        {
            sql.AppendLine($"WHERE {string.Join(" AND ", conditions)}");
        }

        sql.AppendLine($"ORDER BY {lookup.DisplayColumn}");

        return sql.ToString();
    }

    /// <summary>
    /// Validate that filter columns exist in sheet definition
    /// </summary>
    public List<string> ValidateFilters()
    {
        var errors = new List<string>();
        var validColumns = _sheetDef.Columns.Select(c => c.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var filter in _request.Filters)
        {
            if (!validColumns.Contains(filter.Column))
            {
                errors.Add($"Column '{filter.Column}' does not exist in sheet '{_sheetDef.Name}'");
            }
        }

        return errors;
    }

    /// <summary>
    /// Validate that sort columns exist in sheet definition
    /// </summary>
    public List<string> ValidateSorts()
    {
        var errors = new List<string>();
        var validColumns = _sheetDef.Columns.Select(c => c.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var sort in _request.Sorts)
        {
            if (!validColumns.Contains(sort.Column))
            {
                errors.Add($"Column '{sort.Column}' does not exist in sheet '{_sheetDef.Name}'");
            }

            if (!new[] { "asc", "desc" }.Contains(sort.Direction.ToLower()))
            {
                errors.Add($"Invalid sort direction '{sort.Direction}'. Must be 'asc' or 'desc'");
            }
        }

        return errors;
    }
}