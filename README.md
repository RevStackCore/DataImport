# RevStackCore.DataImport

Import and Model bind CSV and Excel files. The library is essentially a wrapper around CsvHelper that also provides support for Excel(.xlsx) files.

[![Build status](https://ci.appveyor.com/api/projects/status/pejda29yjhfwhwq6?svg=true)](https://ci.appveyor.com/project/tachyon1337/dataimport)

# Nuget Installation

``` bash
Install-Package RevStackCore.DataImport

```

# Api

```cs
public interface IDataImport
{
    IEnumerable<T> ImportCsc<T>(string filePath, bool hasHeader = true, bool matchCase = false) where T : class;
    Task<IEnumerable<T>> ImportCsvAsync<T>(string filePath, bool hasHeader = true, bool matchCase = false) where T : class;
    IEnumerable<T> ImportCsv<T>(Stream file, bool hasHeader = true, bool matchCase = false) where T : class;
    Task<IEnumerable<T>> ImportCsvAsync<T>(Stream file, bool hasHeader = true, bool matchCase = false) where T : class;
    IEnumerable<T> ImportExcel<T>(string filePath, bool ignoreHeader = false, bool matchCase=false) where T : class;
    Task<IEnumerable<T>> ImportExcelAsync<T>(string filePath, bool ignoreHeader = false, bool matchCase = false) where T : class;
    IEnumerable<T> ImportExcel<T>(Stream file, bool ignoreHeader = false, bool matchCase = false) where T : class;
    Task<IEnumerable<T>> ImportExcelAsync<T>(Stream file, bool ignoreHeader = false, bool matchCase = false) where T : class;
    void ExportCsv<T>(IEnumerable<T> items, string filePath, bool useQuotes = true) where T : class;
    Task ExportCsvAsync<T>(IEnumerable<T> items, string filePath, bool useQuotes = true) where T : class;
    Stream ExportCsvStream<T>(IEnumerable<T> items, bool useQuotes = true) where T : class;
    Task<Stream> ExportCsvStreamAsync<T>(IEnumerable<T> items, bool useQuotes = true) where T : class;
}

public class FileImport : IDataImport
{
    public IEnumerable<T> ImportCsv<T>(string filePath, bool hasHeader = true, bool matchCase = false) where T : class
    public Task<IEnumerable<T>> ImportCsvAsync<T>(string filePath, bool hasHeader = true, bool matchCase = false) where T : class
    public IEnumerable<T> ImportCsv<T>(Stream file, bool hasHeader = true, bool matchCase = false) where T : class
    public Task<IEnumerable<T>> ImportCsvAsync<T>(Stream file, bool hasHeader = true, bool matchCase = false) where T : class
    public IEnumerable<T> ImportExcel<T>(string filePath, bool ignoreHeader=false, bool matchCase=false) where T : class
    public Task<IEnumerable<T>> ImportExcelAsync<T>(string filePath, bool ignoreHeader = false, bool matchCase=false) where T : class
    public IEnumerable<T> ImportExcel<T>(Stream file, bool ignoreHeader = false, bool matchCase=false) where T : class
    public Task<IEnumerable<T>> ImportExcelAsync<T>(Stream file, bool ignoreHeader = false, bool matchCase=false) where T : class
    public void ExportCsv<T>(IEnumerable<T> items, string filePath, bool useQuotes = true) where T : class
    public Task ExportCsvAsync<T>(IEnumerable<T> items, string filePath, bool useQuotes = true) where T : class
    public Stream ExportCsvStream<T>(IEnumerable<T> items, bool useQuotes = true) where T : class
    public Task<Stream> ExportCsvStreamAsync<T>(IEnumerable<T> items, bool useQuotes = true) where T : class
}
```

# Model Binding

FileImport will model bind the imported data to the passed model class reference based on the column header fields in the .csv or .xlsx file. If the column casing of the file doesn't match the Pascal case of the model class properties, set matchCase=true in the applicable api call. If there is no header, or if the header columns do not match the model class properties at all, use the IndexAttribute property attribute to map fields to class properties by index.

For CSV, if the file has no column header:
hasHeader=false;

For Excel, if the file has no column header:
ignoreHeader=true;

For Excel, if column header does not match model class properties:
ignoreHeader=true;

## Data Annotation
```cs
using CsvHelper.Configuration.Attributes;

public class MyModel
{
    [Index(0)]
    public string MyProperty1 {get; set;}
    [Index(1)]
    public string MyProperty2 {get; set;}
    [Index(2)]
    public string MyProperty3 {get; set;}
}

//Name attribute 
public class MyModel2
{
    [Name("Property One")]
    public string Property1 {get; set;}
    [Name("Property Two")]
    public string Property2 {get; set;}
    [Name("Property Three")]
    public string Property3 {get; set;}
}
```


# Example Usage
```cs
using System;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using RevStackCore.DataImport;

class Program
{
    static async Task Main(string[] args)
    {
        var serviceCollection = new ServiceCollection();
        ConfigureServices(serviceCollection);
        var serviceProvider = serviceCollection.BuildServiceProvider();
        var fileImport = serviceProvider.GetService<IDataImport>();
        var result = await fileImport.ImportCsvAsync<MyModel>("/path/to/file.csv");
        foreach(var model in result)
        {
            //do something with (MyModel)model
        }
    }

    private static void ConfigureServices(IServiceCollection services)
    {
        services
                .AddSingleton<IDataImport, FileImport>()

        services.AddLogging();

    }
}

```




