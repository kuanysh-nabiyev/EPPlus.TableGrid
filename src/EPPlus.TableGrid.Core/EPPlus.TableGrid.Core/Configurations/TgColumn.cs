using System;
using System.Linq.Expressions;
using System.Reflection;
using EPPlus.TableGrid.Core.Exceptions;
using EPPlus.TableGrid.Core.Extensions;
using OfficeOpenXml.Style;

namespace EPPlus.TableGrid.Core.Configurations
{
    public class TgColumn : TgColumnBase
    {
        public TgColumn() { }

        public TgColumn(string propertyName)
        {
            PropertyName = propertyName;
        }

        /// <summary>property name</summary>
        public virtual string PropertyName { get; set; }

        /// <summary>column summary (sum, count, average and etc.)</summary>
        public TgColumnSummary Summary { get; set; }

        internal PropertyInfo PropertyInfo { get; set; }

        internal int OrderNumber { get; set; }

        internal virtual void Validate()
        {
            if (PropertyName == null)
                throw new RequiredPropertyException(nameof(PropertyName), this.GetType());

            Summary?.Validate();
        }
    }

    public class TgColumn<T> : TgColumn
    {
        public TgColumn()
        {
        }

        public TgColumn(Expression<Func<T, object>> propertyExpression)
        {
            Property = propertyExpression;
        }

        /// <summary>property expression to set property name</summary>
        public Expression<Func<T, object>> Property { get; set; }

        public override string PropertyName => Property?.GetPropertyName();

        internal override void Validate()
        {
            if (PropertyName == null && Property == null)
                throw new RequiredPropertyException(nameof(Property), this.GetType());

            Summary?.Validate();
        }
    }

    public class TgRowNumberColumn : TgColumnBase
    {
        public TgRowNumberColumn()
        {
            Header = "#";
            AutoWidth = true;
            Style = new TgExcelStyle()
            {
                HorizontalAlignment = ExcelHorizontalAlignment.Center,
                VerticalAlignment = ExcelVerticalAlignment.Center
            };
            HeaderStyle = new TgExcelStyle()
            {
                HorizontalAlignment = ExcelHorizontalAlignment.Center,
                VerticalAlignment = ExcelVerticalAlignment.Center
            };
        }
    }

    public abstract class TgColumnBase
    {
        /// <summary>header text</summary>
        public string Header { get; set; }

        /// <summary>column width</summary>
        public int Width { get; set; }

        /// <summary>width that corresponds to column text</summary>
        public bool AutoWidth { get; set; }

        /// <summary>column style</summary>
        public TgExcelStyle Style { get; set; }

        /// <summary>column header style</summary>
        public TgExcelStyle HeaderStyle { get; set; }

        internal virtual int PositionInSheet { get; set; }
    }
}