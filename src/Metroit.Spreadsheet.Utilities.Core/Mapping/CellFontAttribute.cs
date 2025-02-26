using System;
using System.ComponentModel;
using System.Drawing;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// Excelセルのフォント情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CellFontAttribute : Attribute
    {
        /// <summary>
        /// フォント を指定します。既定は null です。
        /// </summary>
        public string FontFamily { get; } = null;

        /// <summary>
        /// フォントスタイル を指定します。既定は null です。
        /// </summary>
        public FontStyle? FontStyle { get; } = null;

        /// <summary>
        /// フォントサイズ を指定します。既定は null です。
        /// </summary>
        public float? Size { get; } = null;

        /// <summary>
        /// フォント色 を指定します。既定は null です。
        /// </summary>
        public string Color { get; } = null;

        /// <summary>
        /// 新しい CellFontAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="fontFamily">フォント。</param>
        /// <param name="fontStyle">フォントスタイル。</param>
        public CellFontAttribute(string fontFamily, FontStyle fontStyle)
        {
            FontFamily = fontFamily;
            FontStyle = fontStyle;
        }

        /// <summary>
        /// 新しい CellFontAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="fontFamily">フォント。</param>
        /// <param name="fontStyle">フォントスタイル。</param>
        /// <param name="size">フォントサイズ。</param>
        public CellFontAttribute(string fontFamily, FontStyle fontStyle, float size) : this(fontFamily, fontStyle)
        {
            Size = size;
        }

        /// <summary>
        /// 新しい CellFontAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="fontFamily">フォント。</param>
        /// <param name="fontStyle">フォントスタイル。</param>
        /// <param name="size">フォントサイズ。</param>
        /// <param name="color">フォント色。</param>
        public CellFontAttribute(string fontFamily, FontStyle fontStyle, float size, string color) : this(fontFamily, fontStyle, size)
        {
            Color = color;
        }

        /// <summary>
        /// 新しい CellFontAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="size">フォントサイズ。</param>
        public CellFontAttribute(float size)
        {
            Size = size;
        }

        /// <summary>
        /// 新しい CellFontAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="color">フォント色。</param>
        public CellFontAttribute(string color)
        {
            Color = color;
        }

        /// <summary>
        /// 新しい CellFontAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="size">フォントサイズ。</param>
        /// <param name="color">フォント色。</param>
        public CellFontAttribute(float size, string color) : this(size)
        {
            Color = color;
        }

        /// <summary>
        /// フォント を取得します。
        /// </summary>
        /// <returns>フォント。</returns>
        public Font GetFont()
        {
            try
            {
                var converter = TypeDescriptor.GetConverter(typeof(Font));
                return (Font)converter.ConvertFromString(FontFamily);

            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// フォント色 を取得します。
        /// </summary>
        /// <returns>フォント色。</returns>
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
