namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルの書式情報を提供します。
    /// </summary>
    public class CellFormatItem
    {
        /// <summary>
        /// 書式を指定します。
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 新しい CellFormatItem インスタンスを生成します。
        /// </summary>
        /// <param name="format">書式。</param>
        public CellFormatItem(string format)
        {
            Format = format;
        }
    }
}
