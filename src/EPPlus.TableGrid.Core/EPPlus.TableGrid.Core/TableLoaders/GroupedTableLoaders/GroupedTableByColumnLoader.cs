using System.Collections.Generic;
using System.Linq;
using EPPlus.TableGrid.Core.Configurations;
using OfficeOpenXml;

namespace EPPlus.TableGrid.Core.TableLoaders.GroupedTableLoaders
{
    internal class GroupedTableByColumnLoader<T> : GroupedTableLoader<T>
    {
        public GroupedTableByColumnLoader(TgOptions<T> gridOptions, ExcelWorksheet worksheet)
            : base(gridOptions, worksheet) { }

        protected override void PopulateData()
        {
            TgColumn groupingTgColumn = GetGroupingTgColumn();
            int row = GridOptions.HeaderRowsCount + GridOptions.TableTopPosition;
            int groupIncrement = 1;

            foreach (var groupItems in base.GroupedCollection)
            {
                if (GridOptions.PrintRowNumbers)
                {
                    int col = GridOptions.RowNumberColumn.PositionInSheet;
                    MergeGroupRows(row, col, groupItems.Count());
                    Worksheet.Cells[row, col].Value = groupIncrement++;
                }

                int groupingColumnOrder = groupingTgColumn.PositionInSheet;
                MergeGroupRows(row, groupingColumnOrder, groupItems.Count());

                foreach (var item in groupItems)
                {
                    foreach (var gridColumn in GridOptions.Columns)
                    {
                        var cellValue = gridColumn.PropertyInfo.GetValue(item, null);
                        Worksheet.Cells[row, gridColumn.PositionInSheet].Value = cellValue;
                    }
                    row++;
                }

                base.PrintSummary(row - groupItems.Count(), row - 1);
                if (GridOptions.HasColumnSummary)
                    row++;
            }
        }

        private void MergeGroupRows(int row, int col, int groupItemsCount)
        {
            var lastGroupRow = row + groupItemsCount - 1;
            var groupingSheetColumn = Worksheet.Cells[row, col, lastGroupRow, col];
            groupingSheetColumn.Merge = true;

            if (GridOptions.GroupOptions.IsGroupCollapsable)
            {
                for (var i = row; i <= lastGroupRow; i++)
                {
                    Worksheet.Row(i).OutlineLevel = 1;
                }
            }
        }

        protected override int GetRowsCount()
        {
            var rowsCount = GridOptions.HeaderRowsCount + GridOptions.Collection.Count() - 1;
            if (GridOptions.HasColumnSummary)
                rowsCount += base.GroupedCollection.Count();
            return rowsCount;
        }

        protected override IEnumerable<TgColumn> SetPositionInSheetForEachColumn()
        {
            string groupingColumnName = GridOptions.GroupOptions.GetGroupingColumnName();

            var dataStartColumn = GridOptions.TableLeftPosition + (GridOptions.PrintRowNumbers ? 1 : 0);
            int col = dataStartColumn;
            foreach (var gridColumn in GridOptions.Columns)
            {
                if (gridColumn.PropertyName == groupingColumnName)
                {
                    gridColumn.PositionInSheet = dataStartColumn;
                    gridColumn.OrderNumber = 1;
                }
                else
                {
                    col++;
                    gridColumn.PositionInSheet = col;
                    gridColumn.OrderNumber = col - dataStartColumn + 1;
                }
            }

            if (GridOptions.PrintRowNumbers)
                GridOptions.RowNumberColumn.PositionInSheet = GridOptions.TableLeftPosition;

            return GridOptions.Columns;
        }
    }
}