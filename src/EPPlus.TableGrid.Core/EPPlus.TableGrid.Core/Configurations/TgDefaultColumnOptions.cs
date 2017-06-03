namespace EPPlus.TableGrid.Core.Configurations
{
    public class TgDefaultColumnOptions
    {
        /// <summary>
        /// adjust column width to the longest text
        /// </summary>
        public bool AutoWidth { get; set; }

        /// <summary>
        /// column width. if AutoWidth is true, Width property will be minimal width. 
        /// </summary>
        public int Width { get; set; }

        /// <summary>column style</summary>
        public TgExcelStyle Style { get; set; }

        /// <summary>column header style</summary>
        public TgExcelStyle HeaderStyle { get; set; }
    }
}