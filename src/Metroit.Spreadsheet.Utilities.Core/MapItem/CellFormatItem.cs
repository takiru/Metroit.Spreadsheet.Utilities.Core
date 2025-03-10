namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルの書式情報を提供します。
    /// </summary>
    public class CellFormatItem
    {
        /// <summary>
        /// 書式を取得します。
        /// </summary>
        public string Format { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="format">書式。</param>
        public CellFormatItem(string format)
        {
            Format = format;
        }
    }
}
