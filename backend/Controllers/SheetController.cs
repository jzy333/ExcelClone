using Microsoft.AspNetCore.Mvc;
using ExcelClone.Api.Models;
using ExcelClone.Api.Services;
using System.ComponentModel.DataAnnotations;

namespace ExcelClone.Api.Controllers;

/// <summary>
/// Controller for sheet data operations
/// </summary>
[ApiController]
[Route("api/sheet")]
public class SheetController : ControllerBase
{
    private readonly ISheetDataService _sheetDataService;
    private readonly ILogger<SheetController> _logger;

    public SheetController(ISheetDataService sheetDataService, ILogger<SheetController> logger)
    {
        _sheetDataService = sheetDataService;
        _logger = logger;
    }

    /// <summary>
    /// Query sheet data with filtering, sorting, and pagination
    /// </summary>
    [HttpPost("{sheetId}/query")]
    public async Task<ActionResult<SheetQueryResponse>> QuerySheet(
        string sheetId, 
        [FromBody] SheetQueryRequest request)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sheetId))
            {
                return BadRequest(new { error = "Sheet ID is required" });
            }

            if (request == null)
            {
                return BadRequest(new { error = "Query request is required" });
            }

            // Validate pagination parameters
            if (request.PageSize <= 0 || request.PageSize > 10000)
            {
                return BadRequest(new { error = "PageSize must be between 1 and 10000" });
            }

            if (request.Page < 1)
            {
                return BadRequest(new { error = "Page must be at least 1" });
            }

            _logger.LogInformation("Querying sheet {SheetId} with page {Page}, pageSize {PageSize}", 
                sheetId, request.Page, request.PageSize);

            var response = await _sheetDataService.QuerySheetDataAsync(sheetId, request);
            
            _logger.LogInformation("Query completed for sheet {SheetId}: {Total} total rows, {ReturnedRows} returned", 
                sheetId, response.Total, response.Rows.Count);

            return Ok(response);
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid query request for sheet {SheetId}", sheetId);
            return BadRequest(new { error = ex.Message });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error querying sheet {SheetId}", sheetId);
            return StatusCode(500, new { error = "Failed to query sheet data" });
        }
    }

    /// <summary>
    /// Save bulk changes to sheet data
    /// </summary>
    [HttpPost("{sheetId}/save")]
    public async Task<ActionResult<SheetSaveResponse>> SaveSheet(
        string sheetId, 
        [FromBody] SheetSaveRequest request)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sheetId))
            {
                return BadRequest(new { error = "Sheet ID is required" });
            }

            if (request == null)
            {
                return BadRequest(new { error = "Save request is required" });
            }

            // Validate that we have some operations to perform
            var totalOperations = (request.Inserts?.Count ?? 0) + 
                                (request.Updates?.Count ?? 0) + 
                                (request.Deletes?.Count ?? 0);

            if (totalOperations == 0)
            {
                return BadRequest(new { error = "No operations specified" });
            }

            if (totalOperations > 10000)
            {
                return BadRequest(new { error = "Too many operations (max 10000)" });
            }

            _logger.LogInformation("Saving sheet {SheetId} with {TotalOperations} operations", 
                sheetId, totalOperations);

            var response = await _sheetDataService.SaveSheetDataAsync(sheetId, request);
            
            _logger.LogInformation("Save completed for sheet {SheetId}: {Inserted} inserted, {Updated} updated, {Deleted} deleted", 
                sheetId, response.Results.Inserts.Count, response.Results.Updates.Count, response.Results.Deletes.Count);

            return Ok(response);
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid save request for sheet {SheetId}", sheetId);
            return BadRequest(new { error = ex.Message });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error saving sheet {SheetId}", sheetId);
            return StatusCode(500, new { error = "Failed to save sheet data" });
        }
    }

    /// <summary>
    /// Get lookup values for a specific column
    /// </summary>
    [HttpGet("{sheetId}/lookup/{columnName}")]
    public async Task<ActionResult<List<LookupValue>>> GetLookupValues(
        string sheetId, 
        string columnName,
        [FromQuery] string? search = null,
        [FromQuery] int limit = 100)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sheetId))
            {
                return BadRequest(new { error = "Sheet ID is required" });
            }

            if (string.IsNullOrWhiteSpace(columnName))
            {
                return BadRequest(new { error = "Column name is required" });
            }

            if (limit <= 0 || limit > 1000)
            {
                return BadRequest(new { error = "Limit must be between 1 and 1000" });
            }

            _logger.LogInformation("Getting lookup values for sheet {SheetId}, column {ColumnName}", 
                sheetId, columnName);

            var values = await _sheetDataService.GetLookupValuesAsync(sheetId, columnName, search, limit);
            
            _logger.LogInformation("Retrieved {Count} lookup values for {SheetId}.{ColumnName}", 
                values.Count, sheetId, columnName);

            return Ok(values);
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid lookup request for sheet {SheetId}, column {ColumnName}", 
                sheetId, columnName);
            return BadRequest(new { error = ex.Message });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting lookup values for sheet {SheetId}, column {ColumnName}", 
                sheetId, columnName);
            return StatusCode(500, new { error = "Failed to get lookup values" });
        }
    }

    /// <summary>
    /// Validate data before saving
    /// </summary>
    [HttpPost("{sheetId}/validate")]
    public async Task<ActionResult<Models.ValidationResult>> ValidateData(
        string sheetId, 
        [FromBody] List<Dictionary<string, object?>> rows)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sheetId))
            {
                return BadRequest(new { error = "Sheet ID is required" });
            }

            if (rows == null || rows.Count == 0)
            {
                return BadRequest(new { error = "Rows to validate are required" });
            }

            if (rows.Count > 1000)
            {
                return BadRequest(new { error = "Too many rows to validate (max 1000)" });
            }

            _logger.LogInformation("Validating {RowCount} rows for sheet {SheetId}", rows.Count, sheetId);

            var result = await _sheetDataService.ValidateDataAsync(sheetId, rows);
            
            _logger.LogInformation("Validation completed for sheet {SheetId}: {ErrorCount} errors", 
                sheetId, result.Errors.Count);

            return Ok(result);
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid validation request for sheet {SheetId}", sheetId);
            return BadRequest(new { error = ex.Message });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error validating data for sheet {SheetId}", sheetId);
            return StatusCode(500, new { error = "Failed to validate data" });
        }
    }

    /// <summary>
    /// Get statistics about sheet data
    /// </summary>
    [HttpGet("{sheetId}/stats")]
    public async Task<ActionResult<SheetStats>> GetSheetStats(string sheetId)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sheetId))
            {
                return BadRequest(new { error = "Sheet ID is required" });
            }

            _logger.LogInformation("Getting statistics for sheet {SheetId}", sheetId);

            var stats = await _sheetDataService.GetSheetStatsAsync(sheetId);
            
            _logger.LogInformation("Retrieved statistics for sheet {SheetId}: {RowCount} rows", 
                sheetId, stats.RowCount);

            return Ok(stats);
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid stats request for sheet {SheetId}", sheetId);
            return BadRequest(new { error = ex.Message });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting statistics for sheet {SheetId}", sheetId);
            return StatusCode(500, new { error = "Failed to get sheet statistics" });
        }
    }
}