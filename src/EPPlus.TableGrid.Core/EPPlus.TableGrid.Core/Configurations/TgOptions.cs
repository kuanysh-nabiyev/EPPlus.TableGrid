using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using EPPlus.TableGrid.Core.Exceptions;
using EPPlus.TableGrid.Core.Extensions;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace EPPlus.TableGrid.Core.Configurations
{
    public class TgOptions<T> 
    {
        public TgOptions()
        {
            Collection = Enumerable.Empty<T>();
            Columns = Enumerable.Empty<TgColumn>();
            DefaultColumnOptions = new TgDefaultColumnOptions();
        }

        /// <summary>collection to print</summary>
        public IEnumerable<T> Collection { get; set; }

        /// <summary>list of columns to display</summary>
        public IEnumerable<TgColumn> Columns { get; set; }

        /// <summary>grouping settings (options) for a given Collection</summary>
        public TgGroupOptions<T> GroupOptions { get; set; }

        /// <summary>show headers on top of the table</summary>
        public bool PrintHeaders { get; set; }

        /// <summary>excel table style (Note: does not work for groupable table)</summary>
        public TableStyles TableStyle { get; set; }

        /// <summary>column for row enumeration</summary>
        public TgRowNumberColumn RowNumberColumn { get; set; }

        /// <summary>
        /// show header column numbers (if true it will be located under the header). 
        /// PrintHeaders must be true to show numbers
        /// </summary>
        public bool PrintHeaderColumnNumbers { get; set; }

        /// <summary>Default column option for all columns</summary>
        public TgDefaultColumnOptions DefaultColumnOptions { get; set; }

        /// <summary>start row position</summary>
        internal int TableTopPosition { get; set; }

        /// <summary>start column position</summary>
        internal int TableLeftPosition { get; set; }

        /// <summary>
        /// Use RowNumberColumn property to set row number feature
        /// </summary>
        internal bool PrintRowNumbers => RowNumberColumn != null;

        internal int HeaderRowsCount
        {
            get
            {
                if (PrintHeaders)
                {
                    if (PrintHeaderColumnNumbers)
                        return 2;
                    return 1;
                }
                return 0;
            }
        }

        /// <summary>
        /// Is there aggregate function for any columns
        /// </summary>
        internal bool HasColumnSummary
        {
            get
            {
                if (Columns.Any(it => it.Summary != null))
                    return true;
                return false;
            }
        }

        internal IEnumerable<TgColumn> DisplayableColumns => Columns.Where(it => it.PositionInSheet != -1);

        internal void InitializeColumnsIfEmpty()
        {
            if (!Columns.Any())
            {
                var type = typeof(T);
                var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance).ToList();
                var columnsFromProperties = new List<TgColumn>();
                columnsFromProperties.AddRange(properties.Select(it => new TgColumn(it.Name) {PropertyInfo = it}));
                Columns = columnsFromProperties;
            }
        }

        internal void AddGroupColumnIfNotSet()
        {
            if (GroupOptions != null)
            {
                if (Columns.Count(it => it.PropertyName == GroupOptions.GetGroupingColumnName()) == 0)
                {
                    IList<TgColumn> columns = Columns as IList<TgColumn>;
                    columns.Insert(0, new TgColumn<T>
                    {
                        Property = GroupOptions.GroupingColumn,
                        AutoWidth = true,
                        Style = new TgExcelStyle()
                        {
                            VerticalAlignment = ExcelVerticalAlignment.Center
                        }
                    });
                    Columns = columns;
                }
            }
        }

        internal void SetPropertyInfoForEachColumn()
        {
            var type = typeof(T);
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance).ToList();

            if (properties.Count == 0)
                throw new IncorrectTgOptionsException($"Class {type.Name} must have at least one property. Length is zero");

            foreach (var tgColumn in Columns)
            {
                var propertyInfo = properties.Find(item => item.Name == tgColumn.PropertyName);
                if (propertyInfo == null)
                    throw new IncorrectTgOptionsException($"Property with name {tgColumn.PropertyName} does not exists in class {type.Name}");

                tgColumn.PropertyInfo = propertyInfo;
            }
        }

        internal void Validate()
        {
            if (Collection == null)
                throw new RequiredPropertyException(nameof(Collection), this.GetType());
            if (Columns == null)
                throw new RequiredPropertyException(nameof(Columns), this.GetType());
            else
            {
                foreach (var tgColumn in Columns)
                {
                    tgColumn.Validate();
                }
            }

            if (GroupOptions != null)
            {
                GroupOptions.Validate();

                if (Columns.All(it => it.PropertyName != GroupOptions.GetGroupingColumnName()))
                    throw new IncorrectTgOptionsException("Columns must contain grouping property");
            }
        }
    }
}