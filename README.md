# ExcelEnumerable
[ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) wrapper for supporting LINQ while reading excel files.

## Motivation
The project has been created in order to enable easy way to use C# LINQ while reading from excel files.

## Usage
Install a NuGet package `ExcelEnumerable`, then create a class that represents row with column names which you want to fetch from excel file. You can use `XcelEnumerableColumnAttribute` to specify name of the column when the property name cannot match it:

```c#
public class SampleRow
{
  public string Email { get; set; }
  
  [XcelEnumerableColumn("First Name")]
  public string FirstName { get; set; } 
}
```

Create a new instance of `XcelEnumerable` and specify the previously created type, which represents the row in an excel file, as a generic argument, passing stream or a file name to constructor. After that, apply LINQ expressions:

```c#
using(var excelEnumerable = new XcelEnumerable<SampleRow>(stream)) // or new XcelEnumerable<SampleRow>("fileName"))

{
  var filteredRows = excelEnumerable.Where(r => r.Email.Contains("domain.com")).ToList();
}
```

By default, the first sheet will be selected. In order to read data from another sheet, provide a sheet name in constructor (the name is case sensitive):

```c#
using(var excelEnumerable = new XcelEnumerable<SampleRow>(stream, "Sheet Name"))
{
  ...
}
``` 

## Important when targeting .NET Core Application
When targeting .NET Core, it's important to add dependency on `System.Text.Encoding.CodePages` and register the code page provider during application initialization:

```c#
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
```
