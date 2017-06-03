using EPPlus.TableGrid.Configurations;
using OfficeOpenXml.Style;

namespace EPPlus.TableGrid.Extensions
{
    internal static class ExcelStyleExtensions
    {
        public static void SetStyle(this ExcelStyle excelStyle, TgExcelStyle tgExcelStyle)
        {
            if (tgExcelStyle == null)
                return;
            tgExcelStyle.Fill?.ToExcelFill(excelStyle.Fill);
            tgExcelStyle.Border?.ToExcelBorder(excelStyle.Border);
            tgExcelStyle.Font?.ToExcelFont(excelStyle.Font);

            excelStyle.WrapText = tgExcelStyle.WrapText;
            excelStyle.Hidden = tgExcelStyle.Hidden;
            excelStyle.HorizontalAlignment = tgExcelStyle.HorizontalAlignment;
            excelStyle.VerticalAlignment = tgExcelStyle.VerticalAlignment;
            excelStyle.Indent = tgExcelStyle.Indent;
            excelStyle.Locked = tgExcelStyle.Locked;
            excelStyle.Numberformat.Format = tgExcelStyle.DisplayFormat;
            excelStyle.QuotePrefix = tgExcelStyle.QuotePrefix;
            excelStyle.ReadingOrder = tgExcelStyle.ReadingOrder;
            excelStyle.ShrinkToFit = tgExcelStyle.ShrinkToFit;
            excelStyle.TextRotation = tgExcelStyle.TextRotation;
        }
    }
}