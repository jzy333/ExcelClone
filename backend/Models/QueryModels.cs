namespace ExcelClone.Api.Models;

/// <summary>
/// Lookup value for dropdowns and autocomplete
/// </summary>
public class LookupValue
{
    public string Value { get; set; } = string.Empty;
    public string Display { get; set; } = string.Empty;
    public string? Description { get; set; }
}

/// <summary>
/// Sheet statistics
/// </summary>
public class SheetStats
{
    public long RowCount { get; set; }
    public DateTime? LastModified { get; set; }
    public string? LastModifiedBy { get; set; }
    public Dictionary<string, object> ColumnStats { get; set; } = new();
}

/// <summary>
/// Validation result
/// </summary>
public class ValidationResult
{
    public bool IsValid => !Errors.Any();
    public List<ValidationError> Errors { get; set; } = new();
}

/// <summary>
/// Validation error
/// </summary>
public class ValidationError
{
    public int RowIndex { get; set; }
    public string ColumnName { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string? Value { get; set; }
}

/// <summary>
/// Request model for querying sheet data
/// </summary>
public class SheetQueryRequest
{
    public int Page { get; set; } = 1;
    public int PageSize { get; set; } = 500;
    public List<FilterCriteria> Filters { get; set; } = new();
    public List<SortCriteria> Sorts { get; set; } = new();
    public SelectionInfo? Selection { get; set; }
}

/// <summary>
/// Filter criteria for querying data
/// </summary>
public class FilterCriteria
{
    public string Column { get; set; } = string.Empty;
    public string Operator { get; set; } = string.Empty; // eq, ne, gt, lt, gte, lte, contains, startswith, endswith, in
    public object? Value { get; set; }
    public List<object>? Values { get; set; } // For 'in' operator
}

/// <summary>
/// Sort criteria for querying data
/// </summary>
public class SortCriteria
{
    public string Column { get; set; } = string.Empty;
    public string Direction { get; set; } = "asc"; // asc, desc
}

/// <summary>
/// Selection information for the query
/// </summary>
public class SelectionInfo
{
    public List<string> RangeRefs { get; set; } = new();
}

/// <summary>
/// Response model for sheet query results
/// </summary>
public class SheetQueryResponse
{
    public List<Dictionary<string, object?>> Rows { get; set; } = new();
    public int Total { get; set; }
    public int Page { get; set; }
    public int PageSize { get; set; }
    public QueryMetadata? Metadata { get; set; }
}

/// <summary>
/// Metadata for query results
/// </summary>
public class QueryMetadata
{
    public DateTime QueryTime { get; set; } = DateTime.UtcNow;
    public int ExecutionTimeMs { get; set; }
    public string? SqlQuery { get; set; }
    public Dictionary<string, object> Parameters { get; set; } = new();
}