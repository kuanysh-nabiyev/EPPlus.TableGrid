using System.Collections.Generic;
using System.Linq;
using EPPlus.TableGrid.Configurations;
using EPPlus.TableGrid.Extensions;
using EPPlus.TableGrid.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.TableGrid.TableLoaders
{
    internal abstract class TableLoader<T>
    {
        private readonly TgOptions<T> _gridOptions;
        private readonly ExcelWorksheet _worksheet;
        private readonly SheetColumnHelper _sheetColumnHelper;
        private IEnumerable<TgColumn> _displayableColumns;
        private int _rowsCount;

        protected TableLoader(TgOptions<T> gridOptions, ExcelWorksheet worksheet)
        {
            _gridOptions = gridOptions;
            _worksheet = worksheet;
            _sheetColumnHelper = 
                new SheetColumnHelper(worksheet, gridOptions.TableTopPosition, gridOptions.DefaultColumnOptions);
        }

        protected TgOptions<T> GridOptions => _gridOptions;

        protected ExcelWorksheet Worksheet => _worksheet;

        protected IEnumerable<TgColumn> DisplayableColumns
        {
            get
            {
                if (_displayableColumns == null)
                    _displayableColumns = _gridOptions.Columns.Where(it => it.PositionInSheet != -1).ToList();
                return _displayableColumns;
            }
        }

        protected int RowsCount
        {
            get
            {
                if (_rowsCount == 0)
                    _rowsCount = GetRowsCount();
                return _rowsCount;
            }
        }

        public ExcelRange Load()
        {
            GridOptions.SetPropertyInfoForEachColumn();
            GridOptions.Validate();
            SetPositionInSheetForEachColumn();
            SetHeaderTextForEachColumn();
            SetStyleForEachColumn();
            SetHeaderStyleForEachColumn();
            PopulateData();
            SetWidthForEachColumn();
            SetTableStyle();
            return GetGeneratedTable();
        }

        protected abstract void PopulateData();
        protected abstract int GetRowsCount();
        protected abstract IEnumerable<TgColumn> SetPositionInSheetForEachColumn();

        private void SetHeaderTextForEachColumn()
        {
            if (_gridOptions.PrintHeaders)
            {
                if (_gridOptions.PrintRowNumbers)
                    _sheetColumnHelper.SetHeaderText(_gridOptions.RowNumberColumn);

                foreach (var gridColumn in DisplayableColumns)
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

        private void SetHeaderStyleForEachColumn()
        {
            if (_gridOptions.PrintHeaders)
            {
                if (_gridOptions.PrintRowNumbers)
                    _sheetColumnHelper.SetHeaderStyle(_gridOptions.RowNumberColumn);

                foreach (var gridColumn in DisplayableColumns)
                {
                    if (_gridOptions.PrintHeaderColumnNumbers)
                        _sheetColumnHelper.SetHeaderColumnNumberStyle(gridColumn);

                    _sheetColumnHelper.SetHeaderStyle(gridColumn);
                }
            }
        }

        private void SetStyleForEachColumn()
        {
            if (_gridOptions.PrintRowNumbers)
                _sheetColumnHelper.SetStyle(_gridOptions.RowNumberColumn, RowsCount);

            foreach (var gridColumn in DisplayableColumns)
            {
                _sheetColumnHelper.SetStyle(gridColumn, RowsCount);
            }
        }

        private void SetWidthForEachColumn()
        {
            if (_gridOptions.PrintRowNumbers)
                _sheetColumnHelper.SetWidth(_gridOptions.RowNumberColumn);

            foreach (var gridColumn in DisplayableColumns)
            {
                _sheetColumnHelper.SetWidth(gridColumn);
            }
        }

        protected virtual void SetTableStyle()
        {
            if (_gridOptions.TableStyle != TableStyles.None)
            {
                var table = _worksheet.Tables.Add(GetGeneratedTable(), Name: string.Empty);
                table.ShowHeader = _gridOptions.PrintHeaders;
                table.TableStyle = _gridOptions.TableStyle;
            }
        }

        protected void PrintSummary(int fromRow, int toRow)
        {
            foreach (var gridColumn in DisplayableColumns)
            {
                if (gridColumn.Summary != null)
                {
                    var aggFunction = gridColumn.Summary.AggregateFunction;
                    var aggFunctionName = aggFunction.Type.GetDisplayName();
                    string groupBeginAddress = Worksheet
                        .Cells[fromRow, gridColumn.PositionInSheet].Address;
                    string groupEndAddress = Worksheet
                        .Cells[toRow, gridColumn.PositionInSheet].Address;

                    var cell = Worksheet.Cells[toRow + 1, gridColumn.PositionInSheet];
                    cell.Style.SetStyle(gridColumn.Summary.Style);
                    cell.Style.Locked = true;

                    if (aggFunction.HasCondition)
                    {
                        var condition = aggFunction.Condition;
                        cell.Formula = $"{aggFunctionName}({groupBeginAddress}:{groupEndAddress},\"{condition}\")";
                    }
                    else
                    {
                        cell.Formula = $"{aggFunctionName}({groupBeginAddress}:{groupEndAddress})";
                    }
                }
            }
        }

        private ExcelRange GetGeneratedTable()
        {
            var fromRow = _gridOptions.TableTopPosition;
            var fromCol = _gridOptions.TableLeftPosition;
            var toRow = fromRow + RowsCount;
            var toCol = DisplayableColumns.Max(it => it.PositionInSheet);
            return _worksheet.Cells[fromRow, fromCol, toRow, toCol];
        }
    }
}