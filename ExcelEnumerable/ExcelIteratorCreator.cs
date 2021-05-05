using System;
using System.IO;
using ExcelEnumerable.Configuration;

namespace ExcelEnumerable
{
  public static class ExcelIteratorCreator
  {
    public static IExcelIterator<T> Create<T>(Stream fileStream, 
      Action<ExcelIteratorConfigurationBuilder<T>> builderAction = null) 
      where T : class, new()
    {
      var configurationBuilder = ExcelIteratorConfigurationBuilder<T>.Default();
      
      builderAction?.Invoke(configurationBuilder);

      return new ExcelIterator<T>(fileStream, configurationBuilder.Build());
    }

    public static IExcelIterator<T> Create<T>(Stream fileStream, ExcelIteratorConfiguration configuration)
      where T : class, new() =>
      new ExcelIterator<T>(fileStream, configuration);
  }
}