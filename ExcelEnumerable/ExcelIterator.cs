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

      EnsureCorrectSheetSelected();

      EnsureColumnNamesResolved();

      while (_dataReader.Read())
      {
        var instance = Activator.CreateInstance<T>();

        foreach (var propertyMap in _configuration.PropertyMaps.Where(p => !p.Ignored))
        {
          var valueFromExcel = propertyMap.MapStrategy == PropertyMapStrategy.ByName
            ? _dataReader.GetValue(ResolveIndexByName(propertyMap))
            : _dataReader.GetValue(propertyMap.ColumnIndex);

          propertyMap.Property.SetValue(instance,
            propertyMap.SourceValueConverter.ConvertValue(valueFromExcel));
        }

        yield return instance;
      }
    }

    private void EnsureCorrectSheetSelected()
    {
      if (string.IsNullOrWhiteSpace(_configuration.SheetName)) return;

      while (!_dataReader.Name.Equals(_configuration.SheetName, StringComparison.OrdinalIgnoreCase))
        if (!_dataReader.NextResult()) 
          throw new InvalidOperationException(MessageDefaults.SheetNotFound(_configuration.SheetName));
    }

    private void EnsureColumnNamesResolved()
    {
      if (!_configuration.FirstRowContainsColumnNames)
      {
        _columnIndexAndNamePairs = new Dictionary<string, int>();
        return;
      };

      if (_columnIndexAndNamePairs != null) return;

      if (!_dataReader.Read()) throw new InvalidOperationException(MessageDefaults.CannotReadColumnNames);
      
      if (_dataReader.FieldCount == 0) throw new InvalidOperationException(MessageDefaults.CannotReadColumnNames);

      _columnIndexAndNamePairs = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

      for (var columnIndex = 0; columnIndex < _dataReader.FieldCount; columnIndex++)
      {
        var columnName = _dataReader.GetString(columnIndex);

        if (string.IsNullOrWhiteSpace(columnName))
          if(_configuration.ShouldSkipEmptyColumnNames) continue;
          else throw new InvalidOperationException(MessageDefaults.EmptyColumnNameFound);
        
        _columnIndexAndNamePairs.Add(columnName, columnIndex);
      }
    }

    private int ResolveIndexByName(PropertyMap<T> propertyMap) =>
      _columnIndexAndNamePairs.TryGetValue(propertyMap.ColumnName, out var columnIndex)
        ? columnIndex
        : throw new InvalidOperationException(MessageDefaults.CannotMapColumnByName(propertyMap.ColumnName));

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public void Dispose() => _dataReader?.Dispose();
  }
}