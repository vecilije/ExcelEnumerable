using System;

namespace ExcelEnumerable.ValueConverters
{
  internal class DefaultSourceValueConverter : ISourceValueConverter
  {
    private readonly Type _destinationType;

    public DefaultSourceValueConverter(Type destinationType)
    {
      _destinationType = destinationType;
    }
    
    public object ConvertValue(object sourceValue)
    {
      return sourceValue == null
        ? _destinationType.IsValueType ? Activator.CreateInstance(_destinationType) : null
        : Convert.ChangeType(sourceValue, _destinationType);
    }
  }
}