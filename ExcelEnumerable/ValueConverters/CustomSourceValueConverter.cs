using System;

namespace ExcelEnumerable.ValueConverters
{
  internal class CustomSourceValueConverter : ISourceValueConverter
  {
    private readonly Func<object, object> _convert;

    public CustomSourceValueConverter(Func<object, object> convert)
    {
      _convert = convert;
    }

    public object ConvertValue(object sourceValue) => _convert(sourceValue);
  }
}