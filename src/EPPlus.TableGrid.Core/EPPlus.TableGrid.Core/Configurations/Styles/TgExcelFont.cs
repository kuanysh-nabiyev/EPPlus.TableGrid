using System.Drawing;
using OfficeOpenXml.Style;

namespace EPPlus.TableGrid.Core.Configurations.Styles
{
    public class TgExcelFont
    {
        public TgExcelFont()
        {
            Name = "Calibri";
            Size = 11;
        }

        /// <summary>Cell color</summary>
        public Color Color { get; set; }

        /// <summary>UnderLine type</summary>
        public ExcelUnderLineType UnderLineType { get; set; }

        /// <summary>Font-Vertical Align</summary>
        public ExcelVerticalAlignmentFont VerticalAlign { get; set; }

        /// <summary>Is Font-bold</summary>
        public bool IsBold { get; set; }

        /// <summary>Font family</summary>
        public int Family { get; set; }

        /// <summary>Is Font-italic</summary>
        public bool IsItalic { get; set; }

        /// <summary>The name of the font</summary>
        public string Name { get; set; }

        /// <summary>The Size of the font</summary>
        public float Size { get; set; }

        /// <summary>Is Font-Strikeout</summary>
        public bool IsStrike { get; set; }

        /// <summary>Is Font-Underline</summary>
        public bool IsUnderLine { get; set; }

        internal void ToExcelFont(ExcelFont excelFont)
        {
            if (Color != Color.Empty)
                excelFont.Color.SetColor(Color);
            excelFont.UnderLineType = UnderLineType;
            excelFont.VerticalAlign = VerticalAlign;
            excelFont.Bold = IsBold;
            excelFont.Family = Family;
            excelFont.Italic = IsItalic;
            excelFont.Size = Size;
            excelFont.Strike = IsStrike;
            excelFont.UnderLine = IsUnderLine;
            excelFont.Name = Name;
        }
    }
}