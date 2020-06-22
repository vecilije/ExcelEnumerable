using System;

namespace ExcelEnumerable
{
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