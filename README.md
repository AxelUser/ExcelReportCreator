# ExcelReportCreator
Small Library to create Excel-reports from entities.

# Usage example

```csharp
byte[] binary = ReportBuilder<TestReportEntity>
    .Create("Test report")
    //Create template with object initializer.
    .AddColumn(entity => new ReportColumn()
    {
        Title = "Title",
        Value = entity.Title
    })
    //Or with constructor.
    .AddColumn(entity => new ReportColumn("Date", entity.Date.ToShortDateString()))
    .Build(entities);
```

[Visit project's wiki](https://github.com/AxelUser/ExcelReportCreator/wiki)
