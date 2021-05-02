using System;

namespace ExcelEnumerable
{
  public interface ISourceValueConverter
  {
    object ConvertValue(object sourceValue, Type destinationType);
  }
}