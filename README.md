# EPPlus.TableGrid, EPPlus.TableGrid.Core
Easily create Excel documents from any .NET object collection. 
Easily configure worksheet columns and aggregate summary for each column.  
Feedback are welcome

Install Package for .NET Framework
```
Install-Package EPPlus.TableGrid
```
Install Package for .NET Core
```
Install-Package EPPlus.TableGrid.Core
```

Example of generated Excel shreadsheet with grouping by native language column and aggregate summary for 2 columns (Budget(SUM), Age(AVG)). Each column of the table is configurable (See usage example).   

![output](/screenshots/TableGridExample.png)


Table grid options accept the following parameters:
 - `Collection` - data to show on table (required)
 - `Columns` - list of columns to display
 - `GroupOptions` - grouping settings (options) for a given Collection. There are two types of grouping: ByColumn and ByRow.
 - `DefaultColumnOptions` - default column configuration
 - `PrintHeaders` - set column headers visibility
 - `PrintHeaderColumnNumbers` - set header column numbers visibility (if true it will be located under the header)
 - `RowNumberColumn` - set settings for row numbers column
 - `TableStyle` - set Excel standard table styles (Note: does not work for grouped table. Only for simple (plain) table)
 ===========================================================================
 # Usage example
 ```csharp
  var gridOptions = GetGridOptions(collection);
  var bytes = Spreadsheet.GenerateTableGrid(gridOptions);
  File.WriteAllBytes(GetFilePath(), bytes);
 ```
 
 ```csharp
 TgOptions<Person> GetGridOptions(IEnumerable<Person> persons)
 {
     return new TgOptions<Person>()
     {
         Collection = persons,
         DefaultColumnOptions = new TgDefaultColumnOptions()
         {
             AutoWidth = true,
             Style = new TgExcelStyle
             {
                 HorizontalAlignment = ExcelHorizontalAlignment.Center
             },
             HeaderStyle = new TgExcelStyle
             {
                 HorizontalAlignment = ExcelHorizontalAlignment.Center,
                 VerticalAlignment = ExcelVerticalAlignment.Center,
                 WrapText = true,
                 Font = new TgExcelFont() { IsBold = true }
             }
         },
         Columns = new List<TgColumn>()
         {
             new TgColumn<Person>()
             {
                 Header = "Custom Title",
                 Property = it => it.FirstName,
                 Width = 20,
             },
             new TgColumn<Person>()
             {
                 Property = it => it.LastName,
                 Width = 20
             },
             new TgColumn<Person>()
             {
                 Property = it => it.Email,
                 Style = new TgExcelStyle() {HorizontalAlignment = ExcelHorizontalAlignment.Left}
             },
             new TgColumn<Person>()
             {
                 Property = it => it.Gender
             },
             new TgColumn<Person>()
             {
                 Property = it => it.Budget,
                 Width = 13,
                 Summary = new TgColumnSummary()
                 {
                     AggregateFunction = new AggregateFunction(AggregateFunctionType.Sum),
                     Style = new TgExcelStyle()
                     {
                         HorizontalAlignment = ExcelHorizontalAlignment.Right,
                         Font = new TgExcelFont() {IsBold = true}
                     }
                 }
             },
             new TgColumn<Person>()
             {
                 Property = it => it.Age,
                 Width = 6,
                 Summary = new TgColumnSummary()
                 {
                     AggregateFunction = new AggregateFunction(AggregateFunctionType.Average),
                     Style = new TgExcelStyle()
                     {
                         Font = new TgExcelFont() {IsBold = true}
                     }
                 }
             },
             new TgColumn<Person>()
             {
                 Property = it => it.StreetAddress,
                 Style = new TgExcelStyle()
                 {
                     HorizontalAlignment = ExcelHorizontalAlignment.Right
                 }
             },
         },
         GroupOptions = new TgGroupOptions<Person>()
         {
             GroupingType = GroupingType.GroupHeaderOnColumn,
             GroupingColumn = item => item.NativeLanguage,
         },
         PrintHeaders = true,
         RowNumberColumn = new TgRowNumberColumn(),
         PrintHeaderColumnNumbers = true,
     };
 }
 ```

 ```csharp
  public class Person
  {
      public string FirstName { get; set; }
      public string LastName { get; set; }
      public string Email { get; set; }
      public string Gender { get; set; }
      public string IpAddress { get; set; }
      public decimal Budget { get; set; }
      public int Age { get; set; }
      public string StreetAddress { get; set; }
      public string NativeLanguage { get; set; }
  }
  ```
  ```csharp
  string GetFilePath()
  {
      var folderPath = @"C:\tableGridOutput";
      Directory.CreateDirectory(folderPath);
      var path = $@"{folderPath}\{DateTime.Now:yyyyMMdd_HH_mm_ss}.xlsx";
      return path;
  }
 ```
