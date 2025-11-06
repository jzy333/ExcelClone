namespace ExcelClone.Api.Models;

/// <summary>
/// Represents a workbook manifest with all sheets and their configurations
/// </summary>
public class WorkbookManifest
{
    public List<SheetDefinition> Sheets { get; set; } = new();
}

/// <summary>
/// Defines a sheet within a workbook
/// </summary>
public class SheetDefinition
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Table { get; set; } = string.Empty;
    public List<string> Key { get; set; } = new();
    public List<ColumnDefinition> Columns { get; set; } = new();
    public string RlsScope { get; set; } = string.Empty;
    public bool IsVisible { get; set; } = true;
    public string Color { get; set; } = "#FFFFFF";
    public SheetProtection? Protection { get; set; }
}

/// <summary>
/// Column definition for a sheet
/// </summary>
public class ColumnDefinition
{
    public string Name { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public bool IsEditable { get; set; } = true;
    public bool IsKey { get; set; } = false;
    public bool IsComputed { get; set; } = false;
    public bool IsRequired { get; set; } = false;
    public bool IsLocked { get; set; } = false;
    public ColumnValidation? Validation { get; set; }
    public object? DefaultValue { get; set; }
    public string Format { get; set; } = string.Empty;
    public string HelpText { get; set; } = string.Empty;
    public LookupDefinition? Lookup { get; set; }
}

/// <summary>
/// Column validation rules
/// </summary>
public class ColumnValidation
{
    public string? Regex { get; set; }
    public decimal? MinValue { get; set; }
    public decimal? MaxValue { get; set; }
    public List<string>? AllowedValues { get; set; }
    public int? MaxLength { get; set; }
}

/// <summary>
/// Lookup definition for dropdown columns
/// </summary>
public class LookupDefinition
{
    public string Table { get; set; } = string.Empty;
    public string ValueColumn { get; set; } = string.Empty;
    public string DisplayColumn { get; set; } = string.Empty;
    public string? FilterColumn { get; set; }
    public object? FilterValue { get; set; }
}

/// <summary>
/// Sheet protection settings
/// </summary>
public class SheetProtection
{
    public bool IsLocked { get; set; } = false;
    public List<string> LockedRanges { get; set; } = new();
    public string? PasswordHash { get; set; }
    public string Reason { get; set; } = string.Empty;
}