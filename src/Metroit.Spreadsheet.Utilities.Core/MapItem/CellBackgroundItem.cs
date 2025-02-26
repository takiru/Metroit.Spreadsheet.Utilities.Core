using System.Drawing;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルの背景情報を提供します。
    /// </summary>
    public class CellBackgroundItem
    {
        /// <summary>
        /// 背景色 を指定します。
        /// </summary>
        public Color Color { get; set; }

        /// <summary>
        /// 新しい CellBackgroundItem インスタンスを生成します。
        /// </summary>
        public CellBackgroundItem(Color color)
        {
            Color = color;
        }
    }
}
