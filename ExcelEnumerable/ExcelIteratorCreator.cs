using System;
using System.IO;

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
  }
}