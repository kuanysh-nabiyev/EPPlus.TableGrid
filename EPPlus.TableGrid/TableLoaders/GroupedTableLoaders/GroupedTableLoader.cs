using System;
using System.Collections.Generic;
using System.Linq;
using EPPlus.TableGrid.Configurations;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.TableGrid.TableLoaders.GroupedTableLoaders
{
    internal abstract class GroupedTableLoader<T> : TableLoader<T>
    {
        private IEnumerable<IGrouping<object, T>> _groupedCollection;

        protected GroupedTableLoader(TgOptions<T> gridOptions, ExcelWorksheet worksheet) 
            : base(gridOptions, worksheet) { }

        protected IEnumerable<IGrouping<object, T>> GroupedCollection
        {
            get
            {
                if (_groupedCollection == null)
                {
                    _groupedCollection = GridOptions.Collection
                        .GroupBy(GridOptions.GroupOptions.GroupingColumn.Compile());
                }

                return _groupedCollection;
            }
        }

        protected override void SetTableStyle()
        {
            //if (GridOptions.TableStyle != TableStyles.None)
            //    throw new NotSupportedException(
            //        @"Unfortunately, it's not possible to set table style for merged cells.
            //        Please, remove tableStyle property and set style for each column");
        }

        protected TgColumn GetGroupingTgColumn()
        {
            var groupByPropertyName = base.GridOptions.GroupOptions.GetGroupingColumnName();
            var groupingColumn = base.GridOptions.Columns
                .SingleOrDefault(col => col.PropertyName == groupByPropertyName);
            return groupingColumn;
        }
    }
}