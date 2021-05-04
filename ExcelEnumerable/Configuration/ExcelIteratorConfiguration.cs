using System.Collections.Generic;

namespace ExcelEnumerable.Configuration
{
  public class ExcelIteratorConfiguration<T> where T : class, new()
  {
    public string SheetName { get; set; }
    public bool SkipWhitespaceForColumnNames { get; set; }
    public bool FirstRowContainsColumnNames { get; set; }
    public bool ShouldSkipEmptyColumnNames { get; set; }
    public IEnumerable<ExcelIteratorPropertyMap<T>> PropertyMaps { get; set; }
  }
}