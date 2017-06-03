using System.Collections.Generic;
using System.Linq;
using EPPlus.TableGrid.Core.Configurations;
using EPPlus.TableGrid.Core.Extensions;
using OfficeOpenXml;

namespace EPPlus.TableGrid.Core.TableLoaders.GroupedTableLoaders
{
    internal class GroupedTableByRowLoader<T> : GroupedTableLoader<T>
    {
        public GroupedTableByRowLoader(TgOptions<T> gridOptions, ExcelWorksheet worksheet) 
            : base(gridOptions, worksheet) {}

        protected override void PopulateData()
        {
            var groupingColumn = GetGroupingTgColumn();
            int row = GridOptions.HeaderRowsCount + GridOptions.TableTopPosition;
            var columnsCount = GridOptions.DisplayableColumns.Max(it => it.PositionInSheet);

            foreach (var groupItems in base.GroupedCollection)
            {
                MergeAndSetGroupingRowStyle(row, columnsCount, groupingColumn.Style, groupItems.Count());
                Worksheet.Cells[row++, GridOptions.TableLeftPosition].Value = groupItems.Key;

                int groupRows = 1;
                foreach (var item in groupItems)
                {
                    if (GridOptions.PrintRowNumbers)
                    {
                        var rowNumberCell = Worksheet.Cells[row, GridOptions.TableLeftPosition];
                        rowNumberCell.Value = groupRows++;
                    }

                    foreach (var gridColumn in GridOptions.DisplayableColumns)
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

        private void MergeAndSetGroupingRowStyle(
            int row, int lastColumnInSheet, TgExcelStyle style, int groupItemsCount)
        {
            var fromCol = GridOptions.TableLeftPosition;
            var toCol = lastColumnInSheet;
            var groupingRow = Worksheet.Cells[row, fromCol, row, toCol];
            groupingRow.Merge = true;
            groupingRow.Style.SetStyle(style);

            if (GridOptions.GroupOptions.IsGroupCollapsable)
            {
                var toRow = row + groupItemsCount + (GridOptions.HasColumnSummary ? 1 : 0);
                for (var i = row + 1; i <= toRow; i++)
                {
                    Worksheet.Row(i).OutlineLevel = 1;
                }
            }
        }

        protected override int GetRowsCount()
        {
            var rowsCount = GridOptions.HeaderRowsCount + GroupedCollection.Count() + GridOptions.Collection.Count() - 1;
            if (GridOptions.HasColumnSummary)
                rowsCount += GroupedCollection.Count();
            return rowsCount;
        }

        protected override IEnumerable<TgColumn> SetPositionInSheetForEachColumn()
        {
            string groupingColumnName = GridOptions.GroupOptions.GetGroupingColumnName();

            int dataStartColumn = GridOptions.TableLeftPosition + (GridOptions.PrintRowNumbers ? 1 : 0);
            int col = dataStartColumn;
            foreach (var gridColumn in GridOptions.Columns)
            {
                if (gridColumn.PropertyName == groupingColumnName)
                    gridColumn.PositionInSheet = -1;
                else
                {
                    gridColumn.PositionInSheet = col++;
                    gridColumn.OrderNumber = col - dataStartColumn;
                }
            }

            if (GridOptions.PrintRowNumbers)
                GridOptions.RowNumberColumn.PositionInSheet = GridOptions.TableLeftPosition;

            return GridOptions.Columns;
        }
    }
}