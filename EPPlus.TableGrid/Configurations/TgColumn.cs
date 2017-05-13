using System.Reflection;
using EPPlus.TableGrid.Exceptions;
using OfficeOpenXml.Style;

namespace EPPlus.TableGrid.Configurations
{
    public class TgColumn : TgColumnBase
    {
        public TgColumn() { }

        public TgColumn(string propertyName)
        {
            PropertyName = propertyName;
        }

        /// <summary>business object property name</summary>
        public string PropertyName { get; set; }

        /// <summary>column summary (sum, count, average and etc.)</summary>
        public TgColumnSummary Summary { get; set; }

        internal PropertyInfo PropertyInfo { get; set; }

        internal int OrderNumber { get; set; }

        internal void Validate()
        {
            if (PropertyName == null)
                throw new RequiredPropertyException(nameof(PropertyName), this.GetType());

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

        /// <summary>column width. if AutoWidth is true, Width property will be minimal width</summary>
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