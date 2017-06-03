using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using EPPlus.TableGrid.Core.Configurations;
using EPPlus.TableGrid.Core.Configurations.Styles;
using Newtonsoft.Json;
using OfficeOpenXml.Style;

namespace EPPlus.TableGrid.Core.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string json = File.ReadAllText("Persons.json");
            var persons = JsonConvert.DeserializeObject<List<Person>>(json);

            var gridOptions = GetGridOptions(persons);

            var bytes = Spreadsheet.GenerateTableGrid(gridOptions);
            var path = GetFilePath();
            File.WriteAllBytes(path, bytes);

            //using (var package = new ExcelPackage())
            //{
            //    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
            //    var excelRange = worksheet.GenerateTableGrid(gridOptions);
            //    Console.WriteLine(excelRange.Address);

            //    var path = GetFilePath();
            //    using (Stream stream = File.Create(path))
            //    {
            //        package.SaveAs(stream);
            //    }
            //}
        }

        private static TgOptions<Person> GetGridOptions(IEnumerable<Person> result)
        {
            return new TgOptions<Person>()
            {
                Collection = result,
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


        private static string GetFilePath()
        {
            var folderPath = @"C:\tableGridOutput";
            Directory.CreateDirectory(folderPath);
            var path = $@"{folderPath}\{DateTime.Now:yyyyMMdd_HH_mm_ss}.xlsx";
            return path;
        }
    }
}