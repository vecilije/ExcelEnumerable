using System;

namespace ExcelEnumerable.ValueConverters
{
  internal static class ValueConverterFactory
  {
    public static ISourceValueConverter Create(Type destinationType) =>
      new DefaultSourceValueConverter(destinationType);
  }
}