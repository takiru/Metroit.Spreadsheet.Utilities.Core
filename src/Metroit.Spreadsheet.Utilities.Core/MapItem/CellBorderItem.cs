using Metroit.Spreadsheet.Utilities.Core.Mapping;
using System.Drawing;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルの罫線情報を提供します。
    /// </summary>
    public class CellBorderItem
    {
        /// <summary>
        /// 罫線の位置 を指定します。
        /// </summary>
        public BorderPosition Position { get; set; }

        /// <summary>
        /// 罫線のスタイル を指定します。
        /// </summary>
        public LineStyle Style { get; set; }

        /// <summary>
        /// 罫線の太さ を指定します。
        /// </summary>
        public BorderWeightType Weight { get; set; }

        /// <summary>
        /// 罫線の色 を指定します。
        /// </summary>
        public Color Color { get; set; }

        /// <summary>
        /// 新しい CellBorderItem インスタンスを生成します。
        /// </summary>
        /// <param name="position">罫線の位置。</param>
        /// <param name="style">罫線のスタイル。</param>
        /// <param name="weight">罫線の太さ。</param>
        /// <param name="color">罫線の色。</param>
        public CellBorderItem(BorderPosition position, LineStyle style, BorderWeightType weight, Color color)
        {
            Position = position;
            Style = style;
            Weight = weight;
            Color = color;
        }
    }
}
