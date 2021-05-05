using System.Reflection;
using ExcelEnumerable.ValueConverters;

namespace ExcelEnumerable.Configuration
{
  public class ExcelIteratorPropertyMap
  {
    public PropertyInfo Property { get; set; }
    public ExcelIteratorPropertyMapStrategy MapStrategy { get; set; }
    public int ColumnIndex { get; set; }
    public string ColumnName { get; set; }
    public bool Ignored { get; set; }
    public ISourceValueConverter SourceValueConverter { get; set; }
  }
}