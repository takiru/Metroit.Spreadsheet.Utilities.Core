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
        /// 罫線の色を取得します。既定は <see cref="Color.Black"/> です。
        /// </summary>
        public Color Color { get; } = Color.Black;

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="position">罫線の位置。</param>
        /// <param name="style">罫線のスタイル。</param>
        /// <param name="weight">罫線の太さ。</param>
        public CellBorderAttribute(BorderPosition position, LineStyle style, BorderWeightType weight)
        {
            Position = position;
            Style = style;
            Weight = weight;
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="position">罫線の位置。</param>
        /// <param name="style">罫線のスタイル。</param>
        /// <param name="weight">罫線の太さ。</param>
        /// <param name="color">Red や #FF0000 などのHTML色表現で指定する色。</param>
        public CellBorderAttribute(BorderPosition position, LineStyle style, BorderWeightType weight, string color)
        {
            Position = position;
            Style = style;
            Weight = weight;
            Color = ColorTranslator.FromHtml(color);
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="position">罫線の位置。</param>
        /// <param name="style">罫線のスタイル。</param>
        /// <param name="weight">罫線の太さ。</param>
        /// <param name="red">RGB形式で指定する赤色。</param>
        /// <param name="green">RGB形式で指定する緑色。</param>
        /// <param name="blue">RGB形式で指定する青色。</param>
        public CellBorderAttribute(BorderPosition position, LineStyle style, BorderWeightType weight, int red, int green, int blue)
        {
            Position = position;
            Style = style;
            Weight = weight;
            Color = Color.FromArgb(red, green, blue);
        }
    }
}
