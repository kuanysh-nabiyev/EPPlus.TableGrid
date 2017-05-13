# EPPlus.TableGrid.Core
Generate table by configuring grid options. 
 - Use GenerateTableGrid extension method which accept tableGridOptions as first argument
 - Using second(third) argument you can set location of a table on a worksheet

Install Package
```
Install-Package EPPlus.TableGrid
```

Table grid options accept the following parameters:
 - Collection - data to show on table (required)
 - Columns - list of columns to display
 - GroupOptions - grouping settings (options) for a given Collection. There are two types of grouping: ByColumn and ByRow.
 - PrintHeaders - set column headers visibility
 - PrintHeaderColumnNumbers - set header column numbers visibility (if true it will be located under the header)
 - TableStyle - set Excel standard table styles (Note: does not work for grouped table. Only for simple (plain) table)
 - RowNumberColumn - set settings for row numbers column
 - DefaultColumnWidth - default width for columns widthout particular width
 
 ===========================================================================
 # Usage example
 ```
  var gridOptions = new TgOptions<Person>()
  {
      Collection = new List<Person>()
      {
          new Person("Eric", "Ap #816-6335 Pede. Road", new DateTime(1990, 10, 25), 34.2m),
          new Person("Caesar","Ap #362-9181 Cum Street", new DateTime(1991, 11, 11), 45),
          new Person("Lionel","P.O. Box 923, 806 Sit Rd.", new DateTime(1990, 12, 25), 23),
          new Person("Lionel","P.O. Box 923, 806 Sit Rd.2", new DateTime(1992, 6, 25), 7.8m),
          new Person("Caesar","P.O. Box 923, 806 Sit Rd.2", new DateTime(1992, 2, 25), 67),
          new Person("Caesar","P.O. Box 923, 806 Sit Rd.3", new DateTime(1990, 4, 20), 11.7m)
      };
  }
  
  using (var package = new ExcelPackage())
  {
      var worksheet = package.Workbook.Worksheets.Add("Sheet1");
      var excelRange = worksheet.GenerateTableGrid(gridOptions /*, optionally you can set location of a table*/);

      using (Stream stream = File.Create("yourfilepath"))
      {
          package.SaveAs(stream);
      }
  }
 ```
 
 # Table Grid Options examples:
 Basic example:
 ```
  var gridOptions = new TgOptions<Person>()
  {
      Collection = sampleData
  }
 ```
 
 With table grouping options example:
 ```
  var gridOptions = new TgOptions<Person>()
  {
      Collection = sampleData,
      GroupOptions = new TgGroupOptions<Person>()
      {
          GroupingType = GroupingType.GroupHeaderOnColumn,
          GroupingColumn = item => item.FirstName,
          IsGroupCollapsable = true
      }
  }
 ```
 
 With columns example:
  ```
  var gridOptions = new TgOptions<Person>()
  {
      Collection = sampleData,
      Columns =  new List<TgColumn>()
      {
        new TgColumn()
        {
            Header = "Birthdate Title", PropertyName = "Birthdate", AutoWidth = false,
            Style = new TgExcelStyle
            {
                HorizontalAlignment = ExcelHorizontalAlignment.Center,
                DisplayFormat = "dd.MM.yyy",
                Border = new TgBorder()
                {
                    Left = new TgExcelBorderItem()
                    {
                        Color = Color.Aqua,
                        Style = ExcelBorderStyle.DashDot
                    }
                }
            },
            HeaderStyle = new TgExcelStyle() {WrapText = true}
        },
        new TgColumn()
        {
            Header = "FirstNameTitle", PropertyName = "FirstName", Width = 20,
            Style = new TgExcelStyle()
            {
                WrapText = true,
                HorizontalAlignment = ExcelHorizontalAlignment.Left,
                VerticalAlignment = ExcelVerticalAlignment.Center,
                Fill = new TgExcelFill()
                {
                    BackgroundColor = Color.Brown
                }
            }
        },
        new TgColumn()
        {
            Header = "AddressTitle", PropertyName = "Address", AutoWidth = true,
            HeaderStyle = new TgExcelStyle() {HorizontalAlignment = ExcelHorizontalAlignment.Left}
        },
        new TgColumn()
        {
            Header = "Budget Title", PropertyName = nameof(person.Budget), AutoWidth = true
        },
      } 
  }
 ```
 
 Other examples:
 ```
  var gridOptions = new TgOptions<Person>()
  {
      Collection = sampleData,
      DefaultColumnWidth = 30,
      PrintHeaders = true,
      RowNumberColumn = new TgRowNumberColumn(),
      PrintHeaderColumnNumbers = true,
      TableStyle = TableStyles.Medium18
  }
 ```
 
 ![alt text](https://drive.google.com/open?id=0B520zNYNGOEbMDBmUVBiWVkzcFE)
