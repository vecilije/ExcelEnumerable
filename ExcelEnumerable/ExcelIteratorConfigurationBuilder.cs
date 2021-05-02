using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelEnumerable
{
  public class ExcelIteratorConfigurationBuilder<T> where T : class, new()
  {
    private readonly ExcelIteratorConfiguration<T> _configuration;

    private ExcelIteratorConfigurationBuilder()
    {
      _configuration = new ExcelIteratorConfiguration<T>
      {
        FirstRowContainsColumnNames = true,
        PropertyMaps = typeof(T).GetProperties()
          .Select(p => new PropertyMap<T>
          {
            Property = p,
            MapStrategy = PropertyMapStrategy.ByName,
            ColumnName = p.Name,
            SourceValueConverter = new DefaultSourceValueConverter()
          })
          .ToList()
      };
    }

    public static ExcelIteratorConfigurationBuilder<T> Default() => new ExcelIteratorConfigurationBuilder<T>();

    public void MapByIndex<TProperty>(Expression<Func<T, TProperty>> propertyExpression, int columnIndex)
    {
      if (columnIndex <= 0) throw new ArgumentException(MessageDefaults.InvalidColumnIndex(columnIndex));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.MapStrategy = PropertyMapStrategy.ByIndex;
      propertyMap.ColumnIndex = columnIndex;
    }

    public void MapByName<TProperty>(Expression<Func<T, TProperty>> propertyExpression, string columnName)
    {
      if (string.IsNullOrWhiteSpace(columnName))
        throw new ArgumentException(MessageDefaults.InvalidColumnName(columnName));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.MapStrategy = PropertyMapStrategy.ByName;
      propertyMap.ColumnName = columnName;
    }

    public void ConvertValue<TProperty>(Expression<Func<T, TProperty>> propertyExpression,
      Func<object, TProperty> convert)
    {
      if (convert == null) throw new ArgumentNullException(nameof(convert));

      var propertyMap = GetPropertyMapOrThrowException(propertyExpression);

      propertyMap.SourceValueConverter = new CustomValueConverter(sourceValue => convert(sourceValue));
    }

    public void ConvertValue<TProperty>(Expression<Func<T, TProperty>> propertyExpression,
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

    private PropertyMap<T> GetPropertyMapOrThrowException<TProperty>(Expression<Func<T, TProperty>> propertyExpression)
    {
      var propertyName = (propertyExpression.Body as MemberExpression)?.Member.Name;

      if (string.IsNullOrWhiteSpace(propertyName))
        throw new ArgumentException(MessageDefaults.InvalidPropertyExpression, nameof(propertyExpression));

      var propertyMap = _configuration.PropertyMaps.FirstOrDefault(p => p.Property.Name.Equals(propertyName))
                        ?? throw new ArgumentException(MessageDefaults.PropertyNotFound<T>(propertyName),
                          nameof(propertyExpression));
      return propertyMap;
    }

    public ExcelIteratorConfiguration<T> Build() => _configuration;
  }
}