using System.Collections.Generic;
using System.Linq;
using EPPlus.TableGrid.Core.Configurations;
using OfficeOpenXml;

namespace EPPlus.TableGrid.Core.TableLoaders
{
    internal class SimpleTableLoader<T> : TableLoader<T>
    {
        public SimpleTableLoader(TgOptions<T> gridOptions, ExcelWorksheet worksheet) 
            : base(gridOptions, worksheet) {}

        protected override void PopulateData()
        {
            int row = GridOptions.HeaderRowsCount + GridOptions.TableTopPosition;
            int rowNumber = 1;
            foreach (var item in GridOptions.Collection)
            {
                if (GridOptions.PrintRowNumbers)
                    Worksheet.Cells[row, GridOptions.RowNumberColumn.PositionInSheet].Value = rowNumber++;

                foreach (var gridColumn in GridOptions.Columns)
                {
                    var cellValue = gridColumn.PropertyInfo.GetValue(item, null);
                    Worksheet.Cells[row, gridColumn.PositionInSheet].Value = cellValue;
                }
                row++;
            }

            base.PrintSummary(GridOptions.HeaderRowsCount, row - 1);
        }

        protected override int GetRowsCount()
        {
            var rowsCount = GridOptions.HeaderRowsCount + GridOptions.Collection.Count() - 1;
            if (GridOptions.HasColumnSummary)
                rowsCount++;
            return rowsCount;
        }

        protected override IEnumerable<TgColumn> SetPositionInSheetForEachColumn()
        {
            var dataStartColumn = GridOptions.TableLeftPosition + (GridOptions.PrintRowNumbers ? 1 : 0);
            int col = dataStartColumn;
            foreach (var gridColumn in GridOptions.Columns)
            {
                gridColumn.PositionInSheet = col++;
                gridColumn.OrderNumber = col - dataStartColumn;
            }

            if (GridOptions.PrintRowNumbers)
                GridOptions.RowNumberColumn.PositionInSheet = GridOptions.TableLeftPosition;

            return GridOptions.Columns;
        }
    }
}