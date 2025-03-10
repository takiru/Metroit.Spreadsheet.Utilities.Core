using System.Drawing;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルの背景情報を提供します。
    /// </summary>
    public class CellBackgroundItem
    {
        /// <summary>
        /// 背景色を取得します。
        /// </summary>
        public Color Color { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        public CellBackgroundItem(Color color)
        {
            Color = color;
        }
    }
}
