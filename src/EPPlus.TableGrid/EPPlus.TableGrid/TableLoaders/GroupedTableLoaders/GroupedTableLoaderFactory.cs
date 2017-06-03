using System;
using EPPlus.TableGrid.Configurations;
using OfficeOpenXml;

namespace EPPlus.TableGrid.TableLoaders.GroupedTableLoaders
{
    internal class GroupedTableLoaderFactory
    {
        public GroupedTableLoader<T> Create<T>(TgOptions<T> gridOptions, ExcelWorksheet worksheet)
        {
            GroupingType groupingType = gridOptions.GroupOptions.GroupingType;
            switch (groupingType)
            {
                case GroupingType.GroupHeaderOnColumn:
                    return new GroupedTableByColumnLoader<T>(gridOptions, worksheet);
                case GroupingType.GroupHeaderOnRow:
                    return new GroupedTableByRowLoader<T>(gridOptions, worksheet);
                default:
                    throw new ArgumentException(nameof(groupingType));
            }
        }
    }
}