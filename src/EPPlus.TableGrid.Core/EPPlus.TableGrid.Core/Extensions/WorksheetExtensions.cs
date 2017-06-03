using EPPlus.TableGrid.Core.Configurations;
using EPPlus.TableGrid.Core.TableLoaders;
using OfficeOpenXml;

namespace EPPlus.TableGrid.Core.Extensions
{
    public static class WorksheetExtensions
    {
        /// <summary>
        /// Generate table by a given grid options
        /// </summary>
        /// <typeparam name="T">Type of business objects to show</typeparam>
        /// <param name="worksheet">worksheet</param>
        /// <param name="gridOptions">table grid options</param>
        /// <param name="row">start table row (table top position)</param>
        /// <param name="column">start table column (table left position)</param>
        /// <returns>ExcelRange of a generated table</returns>
        public static ExcelRange GenerateTableGrid<T>(this ExcelWorksheet worksheet, 
            TgOptions<T> gridOptions, int row = 1, int column = 1) where T : class
        {
            return GenerateTableGrid(worksheet, gridOptions, new ExcelCellAddress(row, column));
        }

        /// <summary>
        /// Generate table by a given grid options
        /// </summary>
        /// <typeparam name="T">Type of business objects to show</typeparam>
        /// <param name="worksheet">worksheet</param>
        /// <param name="gridOptions">table grid options</param>
        /// <param name="cellAddress">table top-left position</param>
        /// <returns>ExcelRange of a generated table</returns>
        public static ExcelRange GenerateTableGrid<T>(this ExcelWorksheet worksheet, 
            TgOptions<T> gridOptions, string cellAddress) where T : class
        {
            return GenerateTableGrid(worksheet, gridOptions, new ExcelCellAddress(cellAddress));
        }

        private static ExcelRange GenerateTableGrid<T>(ExcelWorksheet worksheet,
            TgOptions<T> gridOptions, ExcelCellAddress topLeftCellAddress) where T : class
        {
            SetTableLocation(gridOptions, topLeftCellAddress);

            var loaderFactory = new TableLoaderFactory();
            var tableLoader = loaderFactory.Create(gridOptions, worksheet);
            return tableLoader.Load();
        }

        private static void SetTableLocation<T>(TgOptions<T> gridOptions, ExcelCellAddress topLeftCellAddress) 
            where T : class
        {
            if (topLeftCellAddress == null)
                topLeftCellAddress = new ExcelCellAddress(1, 1);

            gridOptions.TableTopPosition = topLeftCellAddress.Row;
            gridOptions.TableLeftPosition = topLeftCellAddress.Column;
        }
    }
}