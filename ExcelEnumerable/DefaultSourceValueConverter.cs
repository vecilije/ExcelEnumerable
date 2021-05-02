using System;

namespace ExcelEnumerable
{
  internal class DefaultSourceValueConverter : ISourceValueConverter
  {
    public object ConvertValue(object sourceValue, Type destinationType)
    {
      return sourceValue == null
        ? (destinationType.IsValueType ? Activator.CreateInstance(destinationType) : null)
        : Convert.ChangeType(sourceValue, destinationType);
    }
  }
}