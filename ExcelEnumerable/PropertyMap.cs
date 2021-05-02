using System.Reflection;

namespace ExcelEnumerable
{
  public class PropertyMap<T> where T : class, new()
  {
    public PropertyInfo Property { get; set; }
    public PropertyMapStrategy MapStrategy { get; set; }
    public int ColumnIndex { get; set; }
    public string ColumnName { get; set; }
    public bool Ignored { get; set; }
    public ISourceValueConverter SourceValueConverter { get; set; }
  }
}