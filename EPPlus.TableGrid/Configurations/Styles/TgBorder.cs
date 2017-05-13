using System.Drawing;
using OfficeOpenXml.Style;

namespace EPPlus.TableGrid.Configurations.Styles
{
    public class TgBorder
    {
        /// <summary>Left border style</summary>
        public TgExcelBorderItem Left { get; set; }

        /// <summary>Bottom border style</summary>
        public TgExcelBorderItem Bottom { get; set; }

        /// <summary>Right border style</summary>
        public TgExcelBorderItem Right { get; set; }

        /// <summary>Top border style</summary>
        public TgExcelBorderItem Top { get; set; }

        /// <summary>Diagonal border style</summary>
        public TgExcelBorderItem Diagonal { get; set; }

        /// <summary>A diagonal from the top left to bottom right of the cell</summary>
        public bool DiagonalDown { get; set; }

        /// <summary>A diagonal from the bottom left to top right of the cell</summary>
        public bool DiagonalUp { get; set; }

        internal void ToExcelBorder(Border excelBorder)
        {
            Left?.ToExcelBorderItem(excelBorder.Left);
            Right?.ToExcelBorderItem(excelBorder.Right);
            Bottom?.ToExcelBorderItem(excelBorder.Bottom);
            Top?.ToExcelBorderItem(excelBorder.Top);
            Diagonal?.ToExcelBorderItem(excelBorder.Diagonal);
            DiagonalDown = excelBorder.DiagonalDown;
            DiagonalUp = excelBorder.DiagonalUp;
        }
    }

    public class TgExcelBorderItem
    {
        /// <summary>The line style of the border</summary>
        public ExcelBorderStyle Style { get; set; }

        /// <summary>The color of the border</summary>
        public Color Color { get; set; }

        internal void ToExcelBorderItem(ExcelBorderItem excelBorderItem)
        {
            excelBorderItem.Style = Style;
            if (Style != ExcelBorderStyle.None)
            {
                if (Color != Color.Empty)
                    excelBorderItem.Color.SetColor(Color);
            }
        }
    }
}