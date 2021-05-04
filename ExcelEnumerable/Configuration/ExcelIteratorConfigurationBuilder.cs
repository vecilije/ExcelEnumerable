using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using ExcelEnumerable.Defaults;
using ExcelEnumerable.ValueConverters;

namespace ExcelEnumerable.Configuration
{
  public class ExcelIteratorConfigurationBuilder<T> where T : class, new()
  {
    private string _sheetName;
    private bool _firstRowContainsColumnNames;
    private bool _emptyColumnNamesSkipped;
    private readonly List<ExcelIteratorPropertyMap<T>> _propertyMaps;
    private bool _skipWhitespaceForColumnNames;

    private ExcelIteratorConfigurationBuilder()
    {
      _skipWhitespaceForColumnNames = false;
      _sheetName = string.Empty;
      _firstRowContainsColumnNames = true;
      _emptyColumnNamesSkipped = true;

      _propertyMaps = typeof(T).GetProperties()
        .Select(p => new ExcelIteratorPropertyMap<T>
        {
          Property = p,
          MapStrategy = ExcelIteratorPropertyMapStrategy.ByName,
          ColumnName = p.Name,
          SourceValueConverter = ValueConverterFactory.Create(p.PropertyType)
        })
        .ToList();
    }

    public static ExcelIteratorConfigurationBuilder<T> Default() => new ExcelIteratorConfigurationBuilder<T>();
    
    public void UseSheetName(string name)
    {
      if (string.IsNullOrWhiteSpace(name)) 
        throw new ArgumentException(MessageDefaults.SheetNameEmpty, nameof(name));

      _sheetName = name;
    }

    public void SkipWhitespaceForColumnNames(bool flag = true) => _skipWhitespaceForColumnNames = flag;

    public void SkipEmptyColumnNames(bool flag = true) => _emptyColumnNamesSkipped = flag;

    public void FirstRowContainsColumnNames(bool flag = true) => _firstRowContainsColumnNames = flag;

    public void MapByIndex<TProperty>(Expression<Func<T, TProperty>> propertyExpression, int columnIndex)
    {
      if (columnIndex < 0) throw new ArgumentException(MessageDefaults.InvalidColumnIndex(columnIndex));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.MapStrategy = ExcelIteratorPropertyMapStrategy.ByIndex;
      propertyMap.ColumnIndex = columnIndex;
    }

    public void MapByName<TProperty>(Expression<Func<T, TProperty>> propertyExpression, string columnName)
    {
      if (string.IsNullOrWhiteSpace(columnName))
        throw new ArgumentException(MessageDefaults.InvalidColumnName(columnName));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.MapStrategy = ExcelIteratorPropertyMapStrategy.ByName;
      propertyMap.ColumnName = columnName;
    }

    public void ConvertSourceValue<TProperty>(Expression<Func<T, TProperty>> propertyExpression,
      Func<object, TProperty> convert)
    {
      if (convert == null) throw new ArgumentNullException(nameof(convert));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.SourceValueConverter = new CustomSourceValueConverter(sourceValue => convert(sourceValue));
    }

    public void ConvertSourceValue<TProperty>(Expression<Func<T, TProperty>> propertyExpression,
      ISourceValueConverter converter)
    {
      if (converter == null) throw new ArgumentNullException(nameof(converter));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.SourceValueConverter = converter;
    }

    public void Ignore<TProperty>(Expression<Func<T, TProperty>> propertyExpression)
    {
      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.Ignored = true;
    }

    private ExcelIteratorPropertyMap<T> GetPropertyMapOrThrowException<TProperty>(Expression<Func<T, TProperty>> propertyExpression)
    {
      var propertyName = (propertyExpression.Body as MemberExpression)?.Member.Name;

      if (string.IsNullOrWhiteSpace(propertyName))
        throw new ArgumentException(MessageDefaults.InvalidPropertyExpression, nameof(propertyExpression));

      var propertyMap = _propertyMaps.FirstOrDefault(p => p.Property.Name.Equals(propertyName))
                        ?? throw new ArgumentException(MessageDefaults.PropertyNotFound<T>(propertyName),
                          nameof(propertyExpression));
      
      return propertyMap;
    }

    public ExcelIteratorConfiguration<T> Build() =>
      new ExcelIteratorConfiguration<T>
      {
        SkipWhitespaceForColumnNames = _skipWhitespaceForColumnNames,
        SheetName = _sheetName,
        FirstRowContainsColumnNames = _firstRowContainsColumnNames,
        ShouldSkipEmptyColumnNames = _emptyColumnNamesSkipped,
        PropertyMaps = _propertyMaps
      };
  }
}