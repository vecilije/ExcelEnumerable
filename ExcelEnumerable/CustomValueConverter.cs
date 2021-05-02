using System;

namespace ExcelEnumerable
{
  internal class CustomValueConverter : ISourceValueConverter
  {
    private readonly Func<object, object> _convert;

    public CustomValueConverter(Func<object, object> convert)
    {
      _convert = convert;
    }

    public object ConvertValue(object sourceValue, Type destinationType) => _convert(sourceValue);
  }
}