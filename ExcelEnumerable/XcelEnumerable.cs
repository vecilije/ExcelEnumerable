using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ExcelDataReader;

namespace ExcelEnumerable
{
  public class XcelEnumerable<TResult> : IExcelEnumerable<TResult>
    where TResult : class
  {
    private readonly Stream _stream;

    public XcelEnumerable(string fileName)
      : this(File.OpenRead(fileName))
    {
    }

    public XcelEnumerable(Stream stream)
    {
      _stream = stream;
    }

    public IEnumerator<TResult> GetEnumerator()
    {
      return ExcelEnumerableGatewayEnumerator.Create(_stream);
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
      return GetEnumerator();
    }

    public void Dispose()
    {
      _stream?.Dispose();
    }


    private class ExcelEnumerableGatewayEnumerator : IEnumerator<TResult>
    {
      private readonly IExcelDataReader _excelDataReader;
      private readonly Dictionary<int, PropertyInfo> _columnsMapping;

      private ExcelEnumerableGatewayEnumerator(
        IExcelDataReader excelDataReader,
        Dictionary<int, PropertyInfo> columnsMapping)
      {
        _excelDataReader = excelDataReader;
        _columnsMapping = columnsMapping;
      }

      public static ExcelEnumerableGatewayEnumerator Create(Stream stream)
      {
        var excelDataReader = ExcelReaderFactory.CreateReader(stream);

        //  Read first row for column names mapping
        excelDataReader.Read();
        var properties = typeof(TResult).GetProperties();

        //  Map column names
        var columnsMapping = new Dictionary<int, PropertyInfo>();

        for (var i = 0; i < excelDataReader.FieldCount; i++)
        {
          var columnName = excelDataReader.GetString(i);
          if (string.IsNullOrWhiteSpace(columnName))
          {
            continue;
          }

          var matchingProperty = properties.SingleOrDefault(p =>
            (p.GetCustomAttribute<XcelEnumerableColumnAttribute>(false)?.ColumnName
              .Equals(columnName, StringComparison.OrdinalIgnoreCase) ?? false)
            || p.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));
          if (matchingProperty == null)
          {
            continue;
          }

          columnsMapping.Add(i, matchingProperty);
        }

        return new ExcelEnumerableGatewayEnumerator(excelDataReader, columnsMapping);
      }

      public bool MoveNext()
      {
        if (!_excelDataReader.Read()) return false;

        Current = CreateCurrent();
        return true;
      }

      private TResult CreateCurrent()
      {
        var resultInstance = Activator.CreateInstance<TResult>();

        foreach (var map in _columnsMapping)
        {
          map.Value.SetValue(resultInstance,
            GetValue(map.Value.PropertyType, map.Key));
        }

        return resultInstance;
      }

      private object GetValue(Type propertyType, int columnIndex)
      {
        var value = _excelDataReader.GetValue(columnIndex);
        return value == null
          ? (propertyType.IsValueType ? Activator.CreateInstance(propertyType) : null)
          : Convert.ChangeType(value, propertyType);
      }

      public void Reset()
      {
        _excelDataReader.Reset();
      }

      public TResult Current { get; private set; }

      object IEnumerator.Current => Current;

      public void Dispose()
      {
        Current = null;
        _excelDataReader?.Dispose();
      }
    }
  }
}