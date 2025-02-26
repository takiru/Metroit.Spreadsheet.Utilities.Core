using System;
using System.Drawing;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// Excelセルの罫線情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class CellBorderAttribute : Attribute
    {
        /// <summary>
        /// 罫線の位置 を指定します。既定は None です。
        /// </summary>
        public BorderPosition Position { get; } = BorderPosition.None;

        /// <summary>
        /// 罫線のスタイル を指定します。既定は LineStyleNone です。
        /// </summary>
        public LineStyle Style { get; } = LineStyle.LineStyleNone;

        /// <summary>
        /// 罫線の太さ を指定します。既定は Thin です。
        /// </summary>
        public BorderWeightType Weight { get; } = BorderWeightType.Thin;

        /// <summary>
        /// 罫線の色 を指定します。既定は null です。
        /// </summary>
        public string Color { get; } = null;

        /// <summary>
        /// 新しい CellBorderAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="position">罫線の位置。</param>
        /// <param name="style">罫線のスタイル。</param>
        /// <param name="weight">罫線の太さ。</param>
        /// <param name="color">罫線の色。</param>
        public CellBorderAttribute(BorderPosition position, LineStyle style, BorderWeightType weight, string color)
        {
            Position = position;
            Style = style;
            Weight = weight;
            Color = color;
        }

        /// <summary>
        /// 罫線の色 を取得します。
        /// </summary>
        /// <returns>罫線の色。</returns>
        public Color GetColor()
        {
            try
            {
                return ColorTranslator.FromHtml(Color);

            }
            catch
            {
                return System.Drawing.Color.Empty;
            }
        }
    }
}
