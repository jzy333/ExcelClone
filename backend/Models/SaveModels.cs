namespace ExcelClone.Api.Models;

/// <summary>
/// Request model for saving sheet data
/// </summary>
public class SheetSaveRequest
{
    public string ClientSessionId { get; set; } = string.Empty;
    public List<RowInsert> Inserts { get; set; } = new();
    public List<RowUpdate> Updates { get; set; } = new();
    public List<RowDelete> Deletes { get; set; } = new();
}

/// <summary>
/// Row insert operation
/// </summary>
public class RowInsert
{
    public Dictionary<string, object?> Data { get; set; } = new();
    public string? ClientId { get; set; } // For client-side tracking
}

/// <summary>
/// Row update operation
/// </summary>
public class RowUpdate
{
    public Dictionary<string, object?> Key { get; set; } = new();
    public string BeforeHash { get; set; } = string.Empty;
    public Dictionary<string, object?> After { get; set; } = new();
    public string? ClientId { get; set; } // For client-side tracking
}

/// <summary>
/// Row delete operation
/// </summary>
public class RowDelete
{
    public Dictionary<string, object?> Key { get; set; } = new();
    public string BeforeHash { get; set; } = string.Empty;
    public string? ClientId { get; set; } // For client-side tracking
}

/// <summary>
/// Response model for save operations
/// </summary>
public class SheetSaveResponse
{
    public bool Ok { get; set; }
    public SaveResults Results { get; set; } = new();
    public string? ErrorMessage { get; set; }
    public DateTime ProcessedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Results for each operation type
/// </summary>
public class SaveResults
{
    public List<OperationResult> Inserts { get; set; } = new();
    public List<OperationResult> Updates { get; set; } = new();
    public List<OperationResult> Deletes { get; set; } = new();
}

/// <summary>
/// Result for a single operation
/// </summary>
public class OperationResult
{
    public Dictionary<string, object?> Key { get; set; } = new();
    public string Status { get; set; } = string.Empty; // merged, conflict, error, deleted, missing
    public string? Reason { get; set; }
    public string? ClientId { get; set; }
    public Dictionary<string, object?>? CurrentData { get; set; } // For conflicts, shows current server data
    public string? CurrentHash { get; set; }
}

/// <summary>
/// Row data with hash for optimistic concurrency
/// </summary>
public class RowData
{
    public Dictionary<string, object?> Data { get; set; } = new();
    public string RowHash { get; set; } = string.Empty;
    public long RowVersion { get; set; }
    public DateTime ModifiedAt { get; set; }
    public string ModifiedBy { get; set; } = string.Empty;
}