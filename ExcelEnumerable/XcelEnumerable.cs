using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelEnumerable.Defaults;

namespace ExcelEnumerable
{
  [Obsolete(MessageDefaults.XcelEnumerableObsolete)]
  public class XcelEnumerable<T> : IExcelEnumerable<T>
    where T : class
  {
    private readonly IExcelDataReader _excelDataReader;
    private readonly Dictionary<int, PropertyInfo> _columnsMapping;
    private readonly string _sheetName;

    public XcelEnumerable(string fileName)
      : this(File.OpenRead(fileName))
    {
    }

    public XcelEnumerable(string fileName, string sheetName)
      : this(File.OpenRead(fileName), sheetName)
    {
    }

    public XcelEnumerable(Stream stream, string sheetName)
      : this(stream)
    {
      if(string.IsNullOrWhiteSpace(sheetName))
        throw new ArgumentException("The provided sheet name is empty.", nameof(sheetName));
      
      _sheetName = sheetName;
    }

    public XcelEnumerable(Stream stream)
    {
      _excelDataReader = ExcelReaderFactory.CreateReader(stream);
      _columnsMapping = new Dictionary<int, PropertyInfo>();
    }

    public IEnumerator<T> GetEnumerator()
    {
      EnsureCorrectSheetSelected();
      
      MapColumnsWithProperties();

      while (_excelDataReader.Read())
      {
        var resultItem = Activator.CreateInstance<T>();
        foreach (var columnMap in _columnsMapping)
          columnMap.Value.SetValue(resultItem, GetValue(columnMap.Value.PropertyType, columnMap.Key));

        yield return resultItem;
      }
    }

    private void EnsureCorrectSheetSelected()
    {
      if (string.IsNullOrWhiteSpace(_sheetName))
        return;

      bool isSheetFound;
      while (true)
      {
        isSheetFound = _excelDataReader.Name == _sheetName;
        if (isSheetFound) break;

        if (!_excelDataReader.NextResult()) break;
      }
      
      if(!isSheetFound)
        throw new InvalidOperationException($"The sheet '{_sheetName}' cannot be found.");
    }

    private void MapColumnsWithProperties()
    {
      //  Read first row for column names mapping
      _excelDataReader.Read();
      var properties = typeof(T).GetProperties();

      //  Map column names
      for (var i = 0; i < _excelDataReader.FieldCount; i++)
      {
        var columnName = _excelDataReader.GetString(i);
        if (string.IsNullOrWhiteSpace(columnName))
          continue;

        var matchingProperty = properties.SingleOrDefault(p =>
          (p.GetCustomAttribute<XcelEnumerableColumnAttribute>(false)?.ColumnName
            .Equals(columnName, StringComparison.OrdinalIgnoreCase) ?? false)
          || p.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));
        if (matchingProperty == null)
          continue;

        _columnsMapping.Add(i, matchingProperty);
      }
    }

    private object GetValue(Type propertyType, int columnIndex)
    {
      var value = _excelDataReader.GetValue(columnIndex);
      return value == null
        ? (propertyType.IsValueType ? Activator.CreateInstance(propertyType) : null)
        : Convert.ChangeType(value, propertyType);
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
      return GetEnumerator();
    }

    public void Dispose()
    {
      _excelDataReader?.Dispose();
    }
  }
}