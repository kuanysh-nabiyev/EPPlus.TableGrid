using System.Collections.Generic;
using System.Linq;
using EPPlus.TableGrid.Core.Configurations;
using EPPlus.TableGrid.Core.Extensions;
using EPPlus.TableGrid.Core.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.TableGrid.Core.TableLoaders
{
    internal abstract class TableLoader<T>
    {
        private SheetTableHelper<T> _sheetTableHelper;

        protected TableLoader(TgOptions<T> gridOptions, ExcelWorksheet worksheet)
        {
            GridOptions = gridOptions;
            Worksheet = worksheet;
        }

        protected TgOptions<T> GridOptions { get; }

        protected ExcelWorksheet Worksheet { get; }

        public ExcelRange Load()
        {
            GridOptions.InitializeColumnsIfEmpty();
            GridOptions.AddGroupColumnIfNotSet();
            GridOptions.Validate();
            GridOptions.SetPropertyInfoForEachColumn();
            this.SetPositionInSheetForEachColumn();

            _sheetTableHelper = new SheetTableHelper<T>(Worksheet, GridOptions, GetRowsCount());

            _sheetTableHelper.SetHeaderTextForEachColumn();
            _sheetTableHelper.SetStyleForEachColumn();
            _sheetTableHelper.SetHeaderStyleForEachColumn();
            this.PopulateData();
            _sheetTableHelper.SetWidthForEachColumn();
            _sheetTableHelper.SetTableStyle();
            return _sheetTableHelper.GetGeneratedTable();
        }

        protected abstract void PopulateData();
        protected abstract int GetRowsCount();
        protected abstract IEnumerable<TgColumn> SetPositionInSheetForEachColumn();

        protected void PrintSummary(int fromRow, int toRow)
        {
            _sheetTableHelper.PrintSummary(fromRow, toRow);
        }
    }
}