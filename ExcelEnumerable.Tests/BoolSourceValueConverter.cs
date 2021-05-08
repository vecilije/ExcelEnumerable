using System;
using ExcelEnumerable.ValueConverters;

namespace ExcelEnumerable.Tests
{
  public class BoolSourceValueConverter : ISourceValueConverter
  {
    private readonly string _trueString;
    private readonly string _falseString;

    public BoolSourceValueConverter(string trueString, string falseString)
    {
      _trueString = trueString;
      _falseString = falseString;
    }

    public object ConvertValue(object sourceValue)
    {
      var sourceString = sourceValue.ToString();

      if (sourceString == _trueString) return true;

      if (sourceString == _falseString) return false;

      throw new InvalidOperationException("Unable to map sourceValue to bool.");
    }
  }
}