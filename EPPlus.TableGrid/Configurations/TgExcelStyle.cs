using EPPlus.TableGrid.Configurations.Styles;
using OfficeOpenXml.Style;

namespace EPPlus.TableGrid.Configurations
{
    public class TgExcelStyle
    {
        /// <summary>Fill</summary>
        public TgExcelFill Fill { get; set; }

        /// <summary>Is wrap text</summary>
        public bool WrapText { get; set; }

        /// <summary>Border</summary>
        public TgBorder Border { get; set; }

        /// <summary>Font</summary>
        public TgExcelFont Font { get; set; }

        /// <summary>
        /// If true the formula is hidden when the sheet is protected.
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// Horizontal alignment
        /// </summary>
        public ExcelHorizontalAlignment HorizontalAlignment { get; set; }

        /// <summary>
        /// Vertical alignment
        /// </summary>
        public ExcelVerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// The margin between the border and the text
        /// </summary>
        public int Indent { get; set; }

        /// <summary>
        /// If true the cell is locked for editing when the sheet is protected
        /// </summary>
        public bool Locked { get; set; }

        /// <summary>
        /// Display format (dd.MM.yyyy - for date;)
        /// </summary>
        public string DisplayFormat { get; set; }

        /// <summary>
        /// If true the cell has a quote prefix, which indicates the value of the cell is prefixed with a single quote.
        /// </summary>
        public bool QuotePrefix { get; set; }

        /// <summary>
        /// Reading order
        /// </summary>
        public ExcelReadingOrder ReadingOrder { get; set; }

        /// <summary>
        /// Shrink to fit
        /// </summary>
        public bool ShrinkToFit { get; set; }

        /// <summary>
        /// Text orientation in degrees. Values range from 0 to 180.
        /// </summary>
        public int TextRotation { get; set; }
    }
}   