using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using EPPlus.TableGrid.Configurations;
using EPPlus.TableGrid.Configurations.Styles;
using EPPlus.TableGrid.Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.TableGrid.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var result = new List<Person>()
            {
                new Person("Eric dsgfsdg sdgfsdgfshj sdhgfsdgfjsd", "Ap #816-6335 Pede. Road", new DateTime(1990, 10, 25), 34.2m),
                new Person("Caesar","Ap #362-9181 Cum Street", new DateTime(1991, 11, 11), 45),
                new Person("Lionel","P.O. Box 923, 806 Sit Rd.", new DateTime(1990, 12, 25), 23),
                new Person("Lionel","P.O. Box 923, 806 Sit Rd.2", new DateTime(1992, 6, 25), 7.8m),
                new Person("Caesar","P.O. Box 923, 806 Sit Rd.2", new DateTime(1992, 2, 25), 67),
                new Person("Caesar","P.O. Box 923, 806 Sit Rd.3", new DateTime(1990, 4, 20), 11.7m)
            };

            var person = result.First();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                var gridOptions = new TgOptions<Person>()
                {
                    Collection = result,
                    DefaultColumnWidth = 30,
                    Columns = new List<TgColumn>()
                    {
                        new TgColumn()
                        {
                            Header = "Birthdate Title", PropertyName = nameof(person.Birthdate), AutoWidth = false,
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
                                Border = new TgBorder()
                                {
                                    Right = new TgExcelBorderItem()
                                    {
                                        Color = Color.Red,
                                        Style = ExcelBorderStyle.Double
                                    }
                                }
                            }
                        },
                        new TgColumn()
                        {
                            Header = "AddressTitle", PropertyName = "Address", AutoWidth = true,
                            Style = new TgExcelStyle()
                            {
                                HorizontalAlignment = ExcelHorizontalAlignment.Right,
                                Border = new TgBorder()
                                {
                                    Left = new TgExcelBorderItem()
                                    {
                                        Color = Color.Green,
                                        Style = ExcelBorderStyle.Double
                                    }
                                }
                            },
                            HeaderStyle = new TgExcelStyle() {HorizontalAlignment = ExcelHorizontalAlignment.Left}
                        },
                        new TgColumn()
                        {
                            Header = "Budget Title", PropertyName = nameof(person.Budget), AutoWidth = true,
                            Style = new TgExcelStyle()
                            {
                                HorizontalAlignment = ExcelHorizontalAlignment.Right,
                            },
                            Summary = new TgColumnSummary()
                            {
                                AggregateFunction = new AggregateFunction(AggregateFunctionType.Sum),
                                Style = new TgExcelStyle()
                                {
                                    HorizontalAlignment = ExcelHorizontalAlignment.Left,
                                    Fill = new TgExcelFill()
                                    {
                                        BackgroundColor = Color.Brown
                                    }
                                }
                            }
                        },
                    },
                    GroupOptions = new TgGroupOptions<Person>()
                    {
                        GroupingType = GroupingType.GroupHeaderOnColumn,
                        GroupingColumn = item => item.FirstName,
                        IsGroupCollapsable = true
                    },
                    PrintHeaders = true,
                    RowNumberColumn = new TgRowNumberColumn(),
                    PrintHeaderColumnNumbers = true,
                    //TableStyle = TableStyles.Medium18
                };
                
                var excelRange = worksheet.GenerateTableGrid(gridOptions, "C10");
                Console.WriteLine(excelRange.Address);

                var path = $@"D:\interviews\files\{DateTime.Now:yyyyMMdd_HH_mm_ss}.xlsx";
                using (Stream stream = File.Create(path))
                {
                    package.SaveAs(stream);
                }
            }
        }

    }
}