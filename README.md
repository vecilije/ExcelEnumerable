# ExcelEnumerable
Lightweight .NET library which enables mapping POCO objects to Excel rows, with ability to apply C# LINQ expressions. It uses [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) for reading Excel files.

Any contribution is very welcome, so, please feel free to fork and create pull requests, including any issue you find with the library.

## Migration to version 2.x.x
`IExcelEnumerable<T>`, `XcelEnumerable<T>` and `[XcelEnumerableColumn("ColumnName")]` are still existing, but are obsolete and will be removed in future versions. Consider using `ExcelIteratorCreator` and `IExcelIterator<T>`.

## Basic Usage
As an example, let's take following Excel file structure:

Id | FirstName | LastName | IsActive
------ | ------- | ------ | ------
1 | First | User | TRUE
2 | Second | User | FALSE

First, we need a class which represents a row:

```c#
public class User
{
    public int Id { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public bool IsActive { get; set; }
}
```

Then, we create a new instance of `IExcelIterator<T>` by using `ExcelIteratorCreator`, by specifying the above type as a generic argument of `Create` method, passing a file stream, and then apply LINQ expressions:

```c#
using (var excelIterator = ExcelIteratorCreator.Create<User>(stream))
{
    var items = excelIterator.Where(u => u.IsActive).ToList();
}
```

`IExcelIterator<T>` implements `IEnumerable<T>`, which means that it can be used in `foreach` statements, and various LINQ expressions can be applied to it.

## Configuration
Additional configuration for column mapping, value conversion and default behaviour can be passed as a second argument of the `ExcelIteratorCreator.Create(stream, config)` method.

```c#
var builder = ExcelIteratorConfigurationBuilder<ExampleRow>.Default();

//  Select sheet by name. Without this configuration, the first sheet will be used by default.
builder.UseSheetName("Some Sheet");

//  Specifying that first row in excel file contains column names. Default: true
builder.FirstRowContainsColumnNames(flag: true);

//  Skips any empty column names, only works when FirstRowContainsColumnNames is true. Default: true
builder.SkipEmptyColumnNames(flag: true);

//  Trim any whitespace when reading colum names from excel file. This can be useful for easy 
//  mapping properties like FirstName, LastName, IsActive with column names 'First Name', 
//  '   Last   Name', ' Is Active' without manually configuring map for each property.
//  When using this behaviour, all columns will be read from excel without whitespace, so 
//  for example, using .MapByName(p => p.FirstName, "First Name") will not work anymore.
//  Only works when FirstRowContainsColumnNames is true. Default: false
builder.TrimWhitespaceForColumnNames(flag: true);

//  Map property by column name (case insensitive).
//  Without this configuration, the column name must exactly match property name (case insensitive)
//  works only when FirstRowContainsColumnNames is true
builder.MapByName(p => p.FirstName, "First NAME");

//  Map property by column index. When FirstRowContainsColumnNames is false, all columns have 
//  to be mapped by index.
builder.MapByIndex(p => p.LastName, 2);

//  Ignore property mapping
builder.Ignore(p => p.Address);

//  Configure custom value conversion for property or implement ISourceValueConverter 
//  and pass an instance of it.
builder.ConvertSourceValue(p => p.IsActive, sourceValue => sourceValue?.ToString() == "TRUE");
builder.ConvertSourceValue(p => p.IsActive, new MySourceValueConverter());

//  Build the configuration
var config = builder.Build();

//  Pass the configuration
using (var excelIterator = ExcelIteratorCreator.Create<User>(stream), config)
{
    ...
}
```

## Important note when targeting .NET Core Application
When targeting .NET Core, it's important to add dependency on `System.Text.Encoding.CodePages` and register the code page provider during application initialization:
```c#
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
```