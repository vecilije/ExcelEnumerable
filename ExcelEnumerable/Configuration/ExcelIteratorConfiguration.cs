using System.Collections.Generic;

namespace ExcelEnumerable.Configuration
{
  /// <summary>
  /// Configuration for <see cref="IExcelIterator{T}" />
  /// </summary>
  public class ExcelIteratorConfiguration
  {
    /// <summary>
    /// Gets or sets the name of the sheet to select while reading data
    /// </summary>
    public string SheetName { get; set; }
    
    /// <summary>
    /// Gets or sets whether first row, in excel file, contains column names
    /// </summary>
    public bool FirstRowContainsColumnNames { get; set; }
    
    /// <summary>
    /// Gets or sets whether to trim a whitespace characters while reading column names from excel.
    /// Works only when <see cref="FirstRowContainsColumnNames"/> is true
    /// </summary>
    public bool TrimWhitespaceForColumnNames { get; set; }
    
    /// <summary>
    /// Gets or sets whether empty column names should be skipped
    /// Works only when <see cref="FirstRowContainsColumnNames"/> is true
    /// </summary>
    public bool ShouldSkipEmptyColumnNames { get; set; }
    
    /// <summary>
    /// Collection of property to excel column mapping
    /// </summary>
    public IEnumerable<ExcelIteratorPropertyMap> PropertyMaps { get; set; }
  }
}