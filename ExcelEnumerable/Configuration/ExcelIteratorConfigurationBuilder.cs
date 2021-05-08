using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using ExcelEnumerable.Defaults;
using ExcelEnumerable.ValueConverters;

namespace ExcelEnumerable.Configuration
{
  /// <summary>
  /// Builder for easier instantiation of <see cref="ExcelIteratorConfiguration"/>
  /// </summary>
  /// <typeparam name="T">
  /// The type which represents one item (row) in an excel file.
  /// Must be a class and have parameterless constructor.
  /// </typeparam>
  public class ExcelIteratorConfigurationBuilder<T> where T : class, new()
  {
    private string _sheetName;
    private bool _firstRowContainsColumnNames;
    private bool _emptyColumnNamesSkipped;
    private readonly List<ExcelIteratorPropertyMap> _propertyMaps;
    private bool _trimWhitespaceForColumnNames;

    private ExcelIteratorConfigurationBuilder()
    {
      _trimWhitespaceForColumnNames = false;
      _sheetName = string.Empty;
      _firstRowContainsColumnNames = true;
      _emptyColumnNamesSkipped = true;

      _propertyMaps = typeof(T).GetProperties()
        .Select(p => new ExcelIteratorPropertyMap
        {
          Property = p,
          MapStrategy = ExcelIteratorPropertyMapStrategy.ByName,
          ColumnName = p.Name,
          SourceValueConverter = ValueConverterFactory.Create(p.PropertyType)
        })
        .ToList();
    }

    /// <summary>
    /// Creates instance of <see cref="ExcelIteratorConfigurationBuilder{T}"/> with default configuration:
    /// <see cref="ExcelIteratorConfiguration.FirstRowContainsColumnNames"/> = true,
    /// <see cref="ExcelIteratorConfiguration.SheetName"/> = string.Empty,
    /// <see cref="ExcelIteratorConfiguration.TrimWhitespaceForColumnNames"/> = true,
    /// <see cref="ExcelIteratorConfiguration.ShouldSkipEmptyColumnNames"/> = true
    /// </summary>
    /// <returns>An instance of <see cref="ExcelIteratorConfigurationBuilder{T}"/></returns>
    public static ExcelIteratorConfigurationBuilder<T> Default() => new ExcelIteratorConfigurationBuilder<T>();
    
    /// <summary>
    /// Sets the sheet name to select
    /// </summary>
    /// <param name="name">The name of the sheet</param>
    /// <exception cref="ArgumentException">The sheet name is empty</exception>
    public void UseSheetName(string name)
    {
      if (string.IsNullOrWhiteSpace(name)) 
        throw new ArgumentException(MessageDefaults.SheetNameEmpty, nameof(name));

      _sheetName = name;
    }

    /// <summary>
    /// Sets whether to trim whitespace while reading column names from excel file
    /// </summary>
    /// <param name="flag">The optional boolean value, default true</param>
    public void TrimWhitespaceForColumnNames(bool flag = true) => _trimWhitespaceForColumnNames = flag;

    /// <summary>
    /// Sets whether, while reading, empty column names should be skipped.
    /// If false, <see cref="IExcelIterator{T}"/> will throw InvalidOperationException while reading column names
    /// </summary>
    /// <param name="flag">The optional boolean value, default true</param>
    public void SkipEmptyColumnNames(bool flag = true) => _emptyColumnNamesSkipped = flag;

    /// <summary>
    /// Sets whether first row contains column names in excel file.
    /// When false, all properties should be mapped by column index
    /// </summary>
    /// <param name="flag">The optional boolean value, default true</param>
    public void FirstRowContainsColumnNames(bool flag = true) => _firstRowContainsColumnNames = flag;

    /// <summary>
    /// Maps specified property to excel column by index.
    /// This can be combined with <see cref="MapByName{TProperty}"/>.
    /// When <see cref="FirstRowContainsColumnNames"/> is false, then all properties have to be mapped by index.
    /// </summary>
    /// <param name="propertyExpression">Expression for specifying the property which is intended for mapping</param>
    /// <param name="columnIndex">The column index from excel</param>
    /// <typeparam name="TProperty">The type of specified property from expression</typeparam>
    /// <exception cref="ArgumentException">When column index is less than 0</exception>
    public void MapByIndex<TProperty>(Expression<Func<T, TProperty>> propertyExpression, int columnIndex)
    {
      if (columnIndex < 0) throw new ArgumentException(MessageDefaults.InvalidColumnIndex(columnIndex));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.MapStrategy = ExcelIteratorPropertyMapStrategy.ByIndex;
      propertyMap.ColumnIndex = columnIndex;
    }

    /// <summary>
    /// Maps specified property to excel column by index.
    /// Extremely useful when property name in a class cannot be mapped to an excel column by default.
    /// Works only when <see cref="FirstRowContainsColumnNames"/> is set to true.
    /// </summary>
    /// <param name="propertyExpression">Expression for specifying the property which is intended for mapping</param>
    /// <param name="columnName">Column name</param>
    /// <typeparam name="TProperty">The type of specified property from expression</typeparam>
    /// <exception cref="ArgumentException">When column name is empty</exception>
    public void MapByName<TProperty>(Expression<Func<T, TProperty>> propertyExpression, string columnName)
    {
      if (string.IsNullOrWhiteSpace(columnName))
        throw new ArgumentException(MessageDefaults.InvalidColumnName(columnName));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.MapStrategy = ExcelIteratorPropertyMapStrategy.ByName;
      propertyMap.ColumnName = columnName;
    }

    /// <summary>
    /// Defines a custom conversion from source value (excel) to destination value (property)
    /// </summary>
    /// <param name="propertyExpression">Expression for specifying the property which is intended for conversion</param>
    /// <param name="convert">The actual conversion method</param>
    /// <typeparam name="TProperty">The type of specified property from expression</typeparam>
    /// <exception cref="ArgumentNullException">When conversion method is null</exception>
    public void ConvertSourceValue<TProperty>(Expression<Func<T, TProperty>> propertyExpression,
      Func<object, TProperty> convert)
    {
      if (convert == null) throw new ArgumentNullException(nameof(convert));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.SourceValueConverter = new CustomSourceValueConverter(sourceValue => convert(sourceValue));
    }

    /// <summary>
    /// Defines a custom conversion from source value (excel) to destination value (property)
    /// </summary>
    /// <param name="propertyExpression">Expression for specifying the property which is intended for conversion</param>
    /// <param name="converter">The converter, implementation of <see cref="ISourceValueConverter"/></param>
    /// <typeparam name="TProperty">The type of specified property from expression</typeparam>
    /// <exception cref="ArgumentNullException">When converter is null</exception>
    public void ConvertSourceValue<TProperty>(Expression<Func<T, TProperty>> propertyExpression,
      ISourceValueConverter converter)
    {
      if (converter == null) throw new ArgumentNullException(nameof(converter));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.SourceValueConverter = converter;
    }

    /// <summary>
    /// Excludes property from mapping
    /// </summary>
    /// <param name="propertyExpression">Expression for specifying the property which is intended for ignore</param>
    /// <typeparam name="TProperty">The type of specified property from expression</typeparam>
    public void Ignore<TProperty>(Expression<Func<T, TProperty>> propertyExpression)
    {
      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.Ignored = true;
    }

    private ExcelIteratorPropertyMap GetPropertyMapOrThrowException<TProperty>(Expression<Func<T, TProperty>> propertyExpression)
    {
      var propertyName = (propertyExpression.Body as MemberExpression)?.Member.Name;

      if (string.IsNullOrWhiteSpace(propertyName))
        throw new ArgumentException(MessageDefaults.InvalidPropertyExpression, nameof(propertyExpression));

      var propertyMap = _propertyMaps.FirstOrDefault(p => p.Property.Name.Equals(propertyName))
                        ?? throw new ArgumentException(MessageDefaults.PropertyNotFound<T>(propertyName),
                          nameof(propertyExpression));
      
      return propertyMap;
    }

    /// <summary>
    /// Create a new instance of <see cref="ExcelIteratorConfiguration"/> based on the current state of builder
    /// </summary>
    /// <returns>A new instance of <see cref="ExcelIteratorConfiguration"/></returns>
    public ExcelIteratorConfiguration Build() =>
      new ExcelIteratorConfiguration
      {
        TrimWhitespaceForColumnNames = _trimWhitespaceForColumnNames,
        SheetName = _sheetName,
        FirstRowContainsColumnNames = _firstRowContainsColumnNames,
        ShouldSkipEmptyColumnNames = _emptyColumnNamesSkipped,
        PropertyMaps = _propertyMaps
      };
  }
}