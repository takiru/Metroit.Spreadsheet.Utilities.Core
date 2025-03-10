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
        /// 罫線の位置を取得します。
        /// </summary>
        public BorderPosition Position { get; }

        /// <summary>
        /// 罫線のスタイルを取得します。
        /// </summary>
        public LineStyle Style { get; }

        /// <summary>
        /// 罫線の太さを取得します。
        /// </summary>
        public BorderWeightType Weight { get; }

        /// <summary>
        /// 罫線の色を取得します。
        /// </summary>
        public Color Color { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
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
