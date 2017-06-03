using System.Drawing;
using OfficeOpenXml.Style;

namespace EPPlus.TableGrid.Core.Configurations.Styles
{
    public class TgExcelFill
    {
        /// <summary>The background color</summary>
        public Color BackgroundColor { get; set; }

        /// <summary>Access to properties for gradient fill.</summary>
        public TgExcelGradientFill Gradient { get; set; }

        /// <summary>The pattern for solid fills.</summary>
        public ExcelFillStyle PatternType { get; set; }

        /// <summary>The color of the pattern</summary>
        public Color PatternColor { get; set; }

        internal void ToExcelFill(ExcelFill excelFill)
        {
            excelFill.PatternType = PatternType;
            if (excelFill.PatternType == ExcelFillStyle.None)
            {
                if (BackgroundColor != Color.Empty || PatternColor != Color.Empty)
                    excelFill.PatternType = ExcelFillStyle.Solid;
            }

            if (BackgroundColor != Color.Empty)
                excelFill.BackgroundColor.SetColor(BackgroundColor);
            if (PatternColor != Color.Empty)
                excelFill.PatternColor.SetColor(PatternColor);

            if (Gradient != null)
            {
                excelFill.Gradient.Type = Gradient.Type;
                excelFill.Gradient.Bottom = Gradient.Bottom;
                excelFill.Gradient.Left = Gradient.Left;
                excelFill.Gradient.Right = Gradient.Right;
                excelFill.Gradient.Top = Gradient.Top;
                excelFill.Gradient.Degree = Gradient.Degree;
                if (Gradient.Color1 != Color.Empty)
                    excelFill.Gradient.Color1.SetColor(Gradient.Color1);
                if (Gradient.Color2 != Color.Empty)
                    excelFill.Gradient.Color2.SetColor(Gradient.Color2);
            }
        }
    }

    public class TgExcelGradientFill
    {
        /// <summary>Linear or Path gradient</summary>
        public ExcelFillGradientType Type { get; set; }

        /// <summary>Gradient Color 1</summary>
        public Color Color1 { get; set; }

        /// <summary>Gradient Color 2</summary>
        public Color Color2 { get; set; }

        /// <summary>
        /// Specifies in percentage format (from the top to the bottom) the position of the bottom edge of the inner rectangle (color 1). For bottom, 0 means the bottom edge of the inner rectangle is on the top edge of the cell, and 1 means it is on the bottom edge of the cell.
        /// </summary>
        public double Bottom { get; set; }

        /// <summary>
        /// Specifies in percentage format(from the top to the bottom) the position of the top edge of the inner rectangle (color 1). For top, 0 means the top edge of the inner rectangle is on the top edge of the cell, and 1 means it is on the bottom edge of the cell. (applies to From Corner and From Center gradients).
        /// </summary>
        public double Top { get; set; }

        /// <summary>
        /// Specifies in percentage format (from the left to the right) the position of the left edge of the inner rectangle (color 1). For left, 0 means the left edge of the inner rectangle is on the left edge of the cell, and 1 means it is on the right edge of the cell. (applies to From Corner and From Center gradients).
        /// </summary>
        public double Left { get; set; }

        /// <summary>
        /// Specifies in percentage format (from the left to the right) the position of the right edge of the inner rectangle (color 1). For right, 0 means the right edge of the inner rectangle is on the left edge of the cell, and 1 means it is on the right edge of the cell. (applies to From Corner and From Center gradients).
        /// </summary>
        public double Right { get; set; }

        /// <summary>Angle of the linear gradient</summary>
        public double Degree { get; set; }
    }
}