using System;
using ExcelEnumerable.Defaults;

namespace ExcelEnumerable
{
  [Obsolete(MessageDefaults.XcelEnumerableColumnAttributeObsolete)]
  [AttributeUsage(AttributeTargets.Property)]
  public class XcelEnumerableColumnAttribute : Attribute
  {
    public XcelEnumerableColumnAttribute(string columnName)
    {
      ColumnName = columnName;
    }
    
    public string ColumnName { get; }
  }
}