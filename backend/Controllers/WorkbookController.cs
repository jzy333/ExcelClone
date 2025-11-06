using Microsoft.AspNetCore.Mvc;
using ExcelClone.Api.Models;
using ExcelClone.Api.Services;

namespace ExcelClone.Api.Controllers;

/// <summary>
/// Controller for workbook operations
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class WorkbookController : ControllerBase
{
    private readonly ILogger<WorkbookController> _logger;

    public WorkbookController(ILogger<WorkbookController> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Get workbook manifest with all sheets and their configurations
    /// </summary>
    [HttpGet("manifest")]
    public ActionResult<WorkbookManifest> GetManifest()
    {
        try
        {
            // For now, return a hardcoded manifest
            // In production, this would come from a configuration service
            var manifest = new WorkbookManifest
            {
                Sheets = new List<SheetDefinition>
                {
                    new()
                    {
                        Id = "financial-data",
                        Name = "Financial Data",
                        Table = "FinancialData",
                        Key = new List<string> { "InternalOrder", "ItemID" },
                        RlsScope = "cost_object",
                        IsVisible = true,
                        Color = "#E3F2FD",
                        Columns = new List<ColumnDefinition>
                        {
                            new()
                            {
                                Name = "InternalOrder",
                                DisplayName = "Internal Order",
                                Type = "string",
                                IsKey = true,
                                IsRequired = true,
                                IsEditable = true,
                                HelpText = "Internal order number for tracking"
                            },
                            new()
                            {
                                Name = "ItemID",
                                DisplayName = "Item ID",
                                Type = "int",
                                IsKey = true,
                                IsRequired = true,
                                IsEditable = true,
                                HelpText = "Unique item identifier"
                            },
                            new()
                            {
                                Name = "Amount",
                                DisplayName = "Amount",
                                Type = "decimal",
                                IsEditable = true,
                                IsRequired = true,
                                Format = "currency",
                                HelpText = "Financial amount in USD",
                                Validation = new ColumnValidation
                                {
                                    MinValue = 0,
                                    MaxValue = 999999999
                                }
                            },
                            new()
                            {
                                Name = "CostCenter",
                                DisplayName = "Cost Center",
                                Type = "string",
                                IsEditable = true,
                                HelpText = "Cost center code",
                                Validation = new ColumnValidation
                                {
                                    Regex = @"^[A-Z]{2}\d{4}$",
                                    MaxLength = 6
                                },
                                Lookup = new LookupDefinition
                                {
                                    Table = "CostCenters",
                                    ValueColumn = "Code",
                                    DisplayColumn = "Name"
                                }
                            },
                            new()
                            {
                                Name = "Description",
                                DisplayName = "Description",
                                Type = "string",
                                IsEditable = true,
                                HelpText = "Item description",
                                Validation = new ColumnValidation
                                {
                                    MaxLength = 500
                                }
                            },
                            new()
                            {
                                Name = "Category",
                                DisplayName = "Category",
                                Type = "string",
                                IsEditable = true,
                                HelpText = "Item category",
                                Validation = new ColumnValidation
                                {
                                    AllowedValues = new List<string> { "OPEX", "CAPEX", "Revenue", "Other" }
                                }
                            },
                            new()
                            {
                                Name = "ModifiedBy",
                                DisplayName = "Modified By",
                                Type = "string",
                                IsEditable = false,
                                IsComputed = true,
                                HelpText = "User who last modified this record"
                            },
                            new()
                            {
                                Name = "ModifiedAt",
                                DisplayName = "Modified At",
                                Type = "datetime",
                                IsEditable = false,
                                IsComputed = true,
                                Format = "datetime",
                                HelpText = "Timestamp of last modification"
                            }
                        }
                    },
                    new()
                    {
                        Id = "budget-data",
                        Name = "Budget Data",
                        Table = "BudgetData",
                        Key = new List<string> { "BudgetYear", "CostCenter" },
                        RlsScope = "cost_object",
                        IsVisible = true,
                        Color = "#F3E5F5",
                        Columns = new List<ColumnDefinition>
                        {
                            new()
                            {
                                Name = "BudgetYear",
                                DisplayName = "Budget Year",
                                Type = "int",
                                IsKey = true,
                                IsRequired = true,
                                IsEditable = true,
                                HelpText = "Budget fiscal year",
                                Validation = new ColumnValidation
                                {
                                    MinValue = 2020,
                                    MaxValue = 2030
                                }
                            },
                            new()
                            {
                                Name = "CostCenter",
                                DisplayName = "Cost Center",
                                Type = "string",
                                IsKey = true,
                                IsRequired = true,
                                IsEditable = true,
                                HelpText = "Cost center code",
                                Lookup = new LookupDefinition
                                {
                                    Table = "CostCenters",
                                    ValueColumn = "Code",
                                    DisplayColumn = "Name"
                                }
                            },
                            new()
                            {
                                Name = "Q1Budget",
                                DisplayName = "Q1 Budget",
                                Type = "decimal",
                                IsEditable = true,
                                Format = "currency",
                                HelpText = "First quarter budget amount"
                            },
                            new()
                            {
                                Name = "Q2Budget",
                                DisplayName = "Q2 Budget",
                                Type = "decimal",
                                IsEditable = true,
                                Format = "currency",
                                HelpText = "Second quarter budget amount"
                            },
                            new()
                            {
                                Name = "Q3Budget",
                                DisplayName = "Q3 Budget",
                                Type = "decimal",
                                IsEditable = true,
                                Format = "currency",
                                HelpText = "Third quarter budget amount"
                            },
                            new()
                            {
                                Name = "Q4Budget",
                                DisplayName = "Q4 Budget",
                                Type = "decimal",
                                IsEditable = true,
                                Format = "currency",
                                HelpText = "Fourth quarter budget amount"
                            },
                            new()
                            {
                                Name = "TotalBudget",
                                DisplayName = "Total Budget",
                                Type = "decimal",
                                IsEditable = false,
                                IsComputed = true,
                                Format = "currency",
                                HelpText = "Computed total of all quarters"
                            },
                            new()
                            {
                                Name = "ModifiedBy",
                                DisplayName = "Modified By",
                                Type = "string",
                                IsEditable = false,
                                IsComputed = true,
                                HelpText = "User who last modified this record"
                            },
                            new()
                            {
                                Name = "ModifiedAt",
                                DisplayName = "Modified At",
                                Type = "datetime",
                                IsEditable = false,
                                IsComputed = true,
                                Format = "datetime",
                                HelpText = "Timestamp of last modification"
                            }
                        }
                    }
                }
            };

            _logger.LogInformation("Retrieved workbook manifest with {SheetCount} sheets", manifest.Sheets.Count);
            return Ok(manifest);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving workbook manifest");
            return StatusCode(500, new { error = "Failed to retrieve workbook manifest" });
        }
    }

    /// <summary>
    /// Get schema information for a specific sheet
    /// </summary>
    [HttpGet("sheet/{sheetId}/schema")]
    public ActionResult<SheetDefinition> GetSheetSchema(string sheetId)
    {
        try
        {
            // Get the manifest and find the requested sheet
            var manifest = GetManifest().Value;
            var sheet = manifest?.Sheets.FirstOrDefault(s => s.Id == sheetId);

            if (sheet == null)
            {
                return NotFound(new { error = $"Sheet '{sheetId}' not found" });
            }

            _logger.LogInformation("Retrieved schema for sheet {SheetId}", sheetId);
            return Ok(sheet);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving sheet schema for {SheetId}", sheetId);
            return StatusCode(500, new { error = "Failed to retrieve sheet schema" });
        }
    }
}