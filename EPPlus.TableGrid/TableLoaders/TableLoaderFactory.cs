using EPPlus.TableGrid.Configurations;
using EPPlus.TableGrid.TableLoaders.GroupedTableLoaders;
using OfficeOpenXml;

namespace EPPlus.TableGrid.TableLoaders
{
    internal class TableLoaderFactory
    {
        public TableLoader<T> Create<T>(TgOptions<T> gridOptions, ExcelWorksheet worksheet)
        {
            if (gridOptions.GroupOptions == null)
            {
                return new SimpleTableLoader<T>(gridOptions, worksheet);
            }
            else
            {
                var loaderFactory = new GroupedTableLoaderFactory();
                return loaderFactory.Create(gridOptions, worksheet);
            }
        }
    }
}