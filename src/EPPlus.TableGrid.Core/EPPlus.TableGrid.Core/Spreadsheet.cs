using EPPlus.TableGrid.Core.Configurations;
using EPPlus.TableGrid.Core.Extensions;
using OfficeOpenXml;

namespace EPPlus.TableGrid.Core
{
    public class Spreadsheet
    {
        public static byte[] GenerateTableGrid<T>(TgOptions<T> gridOptions, string worksheetName = null) where T : class
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(worksheetName ?? typeof(T).Name);
                worksheet.GenerateTableGrid(gridOptions);
                return package.GetAsByteArray();
            }
        }
    }
}