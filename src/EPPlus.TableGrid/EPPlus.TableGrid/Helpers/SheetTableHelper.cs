using System.Linq;
using EPPlus.TableGrid.Configurations;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.TableGrid.Helpers
{
    internal class SheetTableHelper<T>
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly TgOptions<T> _gridOptions;
        private readonly int _rowsCount;
        private readonly SheetColumnHelper _sheetColumnHelper;

        public SheetTableHelper(ExcelWorksheet worksheet, TgOptions<T> gridOptions, int rowsCount)
        {
            _worksheet = worksheet;
            _gridOptions = gridOptions;
            _rowsCount = rowsCount;
            _sheetColumnHelper = 
                new SheetColumnHelper(worksheet, gridOptions.TableTopPosition, gridOptions.DefaultColumnOptions);
        }

        public void SetHeaderTextForEachColumn()
        {
            if (_gridOptions.PrintHeaders)
            {
                if (_gridOptions.PrintRowNumbers)
                    _sheetColumnHelper.SetHeaderText(_gridOptions.RowNumberColumn);

                foreach (var gridColumn in _gridOptions.DisplayableColumns)
                {
                    if (_gridOptions.PrintHeaderColumnNumbers)
                    {
                        var columnNumbersRow = _gridOptions.TableTopPosition + 1;
                        _worksheet.Cells[columnNumbersRow, gridColumn.PositionInSheet].Value = gridColumn.OrderNumber;
                    }
                    _sheetColumnHelper.SetHeaderText(gridColumn);
                }
            }
        }

        public void SetHeaderStyleForEachColumn()
        {
            if (_gridOptions.PrintHeaders)
            {
                if (_gridOptions.PrintRowNumbers)
                    _sheetColumnHelper.SetHeaderStyle(_gridOptions.RowNumberColumn);

                foreach (var gridColumn in _gridOptions.DisplayableColumns)
                {
                    if (_gridOptions.PrintHeaderColumnNumbers)
                        _sheetColumnHelper.SetHeaderColumnNumberStyle(gridColumn);

                    _sheetColumnHelper.SetHeaderStyle(gridColumn);
                }
            }
        }

        public void SetStyleForEachColumn()
        {
            if (_gridOptions.PrintRowNumbers)
                _sheetColumnHelper.SetStyle(_gridOptions.RowNumberColumn, _rowsCount);

            foreach (var gridColumn in _gridOptions.DisplayableColumns)
            {
                _sheetColumnHelper.SetStyle(gridColumn, _rowsCount);
            }
        }

        public void SetWidthForEachColumn()
        {
            if (_gridOptions.PrintRowNumbers)
                _sheetColumnHelper.SetWidth(_gridOptions.RowNumberColumn);

            foreach (var gridColumn in _gridOptions.DisplayableColumns)
            {
                _sheetColumnHelper.SetWidth(gridColumn);
            }
        }

        public virtual void SetTableStyle()
        {
            if (_gridOptions.TableStyle != TableStyles.None)
            {
                var table = _worksheet.Tables.Add(GetGeneratedTable(), Name: string.Empty);
                table.ShowHeader = _gridOptions.PrintHeaders;
                table.TableStyle = _gridOptions.TableStyle;
            }
        }

        public void PrintSummary(int fromRow, int toRow)
        {
            var columnsWithSummary = _gridOptions.DisplayableColumns.Where(it => it.Summary != null).ToList();
            columnsWithSummary.ForEach(column => _sheetColumnHelper.PrintSummary(column, fromRow, toRow));
        }

        public ExcelRange GetGeneratedTable()
        {
            var fromRow = _gridOptions.TableTopPosition;
            var fromCol = _gridOptions.TableLeftPosition;
            var toRow = fromRow + _rowsCount;
            var toCol = _gridOptions.DisplayableColumns.Max(it => it.PositionInSheet);
            return _worksheet.Cells[fromRow, fromCol, toRow, toCol];
        }
    }
}