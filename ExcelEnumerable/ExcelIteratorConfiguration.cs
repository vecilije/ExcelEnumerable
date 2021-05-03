using System.Collections.Generic;

namespace ExcelEnumerable
{
  public class ExcelIteratorConfiguration<T> where T : class, new()
  {
    public string SheetName { get; set; }
    public bool FirstRowContainsColumnNames { get; set; }
    public bool ShouldSkipEmptyColumnNames { get; set; }
    public IEnumerable<PropertyMap<T>> PropertyMaps { get; set; }
  }
}