using System.Reflection;
using ExcelEnumerable.ValueConverters;

namespace ExcelEnumerable.Configuration
{
  /// <summary>
  /// Configuration for property to excel map
  /// </summary>
  public class ExcelIteratorPropertyMap
  {
    /// <summary>
    /// Gets or sets the property fow which this map is intended to.
    /// </summary>
    public PropertyInfo Property { get; set; }
    
    /// <summary>
    /// Gets or sets the map strategy
    /// </summary>
    public ExcelIteratorPropertyMapStrategy MapStrategy { get; set; }
    
    /// <summary>
    /// Gets or sets the index of an excel column.
    /// Works when <see cref="MapStrategy"/> is <see cref="ExcelIteratorPropertyMapStrategy.ByIndex"/>
    /// </summary>
    public int ColumnIndex { get; set; }
    
    /// <summary>
    /// Gets or sets the name of an excel column.
    /// Works when <see cref="MapStrategy"/> is <see cref="ExcelIteratorPropertyMapStrategy.ByName"/>
    /// </summary>
    public string ColumnName { get; set; }
    
    /// <summary>
    /// Gets or sets whether the property is ignored for mapping
    /// </summary>
    public bool Ignored { get; set; }
    
    /// <summary>
    /// Gets or sets the converter from source value (excel) to destination value (property)
    /// </summary>
    public ISourceValueConverter SourceValueConverter { get; set; }
  }
}