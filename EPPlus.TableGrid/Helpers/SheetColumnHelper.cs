using EPPlus.TableGrid.Configurations;
using EPPlus.TableGrid.Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.TableGrid.Helpers
{
    public class SheetColumnHelper
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly int _row;

        public SheetColumnHelper(ExcelWorksheet worksheet, int row)
        {
            _worksheet = worksheet;
            _row = row;
        }

        public void SetHeaderText(TgColumn tgColumn)
        {
            var cell = _worksheet.Cells[_row, tgColumn.PositionInSheet];
            cell.Value = tgColumn.Header ?? tgColumn.PropertyName;
        }

        public void SetHeaderText(TgColumnBase tgColumn)
        {
            var cell = _worksheet.Cells[_row, tgColumn.PositionInSheet];
            cell.Value = tgColumn.Header;
        }

        public void SetHeaderStyle(TgColumnBase tgColumn)
        {
            var cell = _worksheet.Cells[_row, tgColumn.PositionInSheet];
            cell.Style.SetStyle(tgColumn.HeaderStyle);
        }

        public void SetStyle(TgColumnBase tgColumn, int rowsCount)
        {
            var sheetColumn = _worksheet
                .Cells[_row, tgColumn.PositionInSheet, _row + rowsCount, tgColumn.PositionInSheet];
            sheetColumn.Style.SetStyle(tgColumn.Style);
        }

        public void SetWidth(TgColumnBase tgColumn, int defaultWidth = 0)
        {
            var sheetColumn = _worksheet.Column(tgColumn.PositionInSheet);

            if (tgColumn.AutoWidth)
                sheetColumn.AutoFit(tgColumn.Width);
            else
            {
                if (tgColumn.Width <= 0)
                {
                    if (defaultWidth <= 0)
                        tgColumn.AutoWidth = true;
                    else
                        sheetColumn.Width = defaultWidth;
                }
                else
                {
                    sheetColumn.Width = tgColumn.Width;
                }
            }
        }

        public void SetHeaderColumnNumberStyle(TgColumn gridColumn)
        {
            var cell = _worksheet.Cells[_row + 1, gridColumn.PositionInSheet];
            cell.Style.SetStyle(gridColumn.HeaderStyle);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
    }
}