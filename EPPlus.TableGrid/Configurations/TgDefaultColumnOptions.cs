namespace EPPlus.TableGrid.Configurations
{
    public class TgDefaultColumnOptions
    {
        /// <summary>
        /// column width. if AutoWidth is true, Width property will be minimal width. 
        /// if Width property doesn't set, autowidth will be applied by default
        /// </summary>
        public int Width { get; set; }

        /// <summary>column style</summary>
        public TgExcelStyle Style { get; set; }

        /// <summary>column header style</summary>
        public TgExcelStyle HeaderStyle { get; set; }
    }
}