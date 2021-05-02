using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDataReader;

namespace ExcelEnumerable
{
  internal class ExcelIterator<T> : IExcelIterator<T> where T : class, new()
  {
    private readonly ExcelIteratorConfiguration<T> _configuration;
    private readonly IExcelDataReader _dataReader;

    private IDictionary<string, int> _columnIndexAndNamePairs;

    public ExcelIterator(Stream stream, ExcelIteratorConfiguration<T> configuration)
    {
      _configuration = configuration;
      _dataReader = ExcelReaderFactory.CreateReader(stream);
    }

    public IEnumerator<T> GetEnumerator()
    {
      _dataReader.Reset();

      if (!string.IsNullOrWhiteSpace(_configuration.SheetName)) EnsureSheetSelected();

      if (_configuration.FirstRowContainsColumnNames) EnsureColumnNamesResolved();

      while (_dataReader.Read())
      {
        var instance = Activator.CreateInstance<T>();

        foreach (var propertyMap in _configuration.PropertyMaps.Where(p => !p.Ignored))
        {
          var valueFromExcel = propertyMap.MapStrategy == PropertyMapStrategy.ByName
            ? _dataReader.GetValue(ResolveIndexByName(propertyMap))
            : _dataReader.GetValue(propertyMap.ColumnIndex);

          propertyMap.Property.SetValue(instance,
            propertyMap.SourceValueConverter.ConvertValue(valueFromExcel, propertyMap.Property.PropertyType));
        }

        yield return instance;
      }
    }

    private void EnsureSheetSelected()
    {
      if (string.IsNullOrWhiteSpace(_configuration.SheetName))
        return;

      bool isSheetFound;
      while (true)
      {
        isSheetFound = _dataReader.Name == _configuration.SheetName;
        if (isSheetFound) break;

        if (!_dataReader.NextResult()) break;
      }

      if (!isSheetFound)
        throw new InvalidOperationException(MessageDefaults.SheetNameNotFound(_configuration.SheetName));
    }

    private void EnsureColumnNamesResolved()
    {
      if (_columnIndexAndNamePairs != null) return;

      if (!_dataReader.Read()) return;

      _columnIndexAndNamePairs = new Dictionary<string, int>();

      for (var columnIndex = 0; columnIndex < _dataReader.FieldCount; columnIndex++)
        _columnIndexAndNamePairs.Add(_dataReader.GetString(columnIndex), columnIndex);
    }

    private int ResolveIndexByName(PropertyMap<T> propertyMap) =>
      _columnIndexAndNamePairs.TryGetValue(propertyMap.ColumnName, out var columnIndex)
        ? columnIndex : throw new InvalidOperationException(
          MessageDefaults.CannotMapColumnByName(propertyMap.ColumnName));

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public void Dispose() => _dataReader?.Dispose();
  }
}