using System;
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
        /// フォント名を取得します。既定は未指定を示す <see cref="string.Empty"/> です。
        /// </summary>
        public string FamilyName { get; } = string.Empty;

        /// <summary>
        /// フォントスタイル を指定します。既定は <see cref="FontStyle.Regular"/> です。
        /// </summary>
        public FontStyle FontStyle { get; } = FontStyle.Regular;

        /// <summary>
        /// フォントサイズ を指定します。既定は未指定を示す -1 です。
        /// </summary>
        public float Size { get; } = -1;

        /// <summary>
        /// フォント色 を指定します。既定は <see cref="Color.Black"/> です。
        /// </summary>
        public Color Color { get; } = Color.Black;

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="familyName">フォント名。</param>
        /// <param name="fontStyle">フォントスタイル。</param>
        public CellFontAttribute(string familyName, FontStyle fontStyle)
        {
            FamilyName = familyName;
            FontStyle = fontStyle;
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="familyName">フォント名。</param>
        /// <param name="size">フォントサイズ。</param>
        /// <param name="fontStyle">フォントスタイル。</param>
        /// <remarks>
        /// <paramref name="size"/> は 0 より大きくなければなりません。
        /// </remarks>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="size"/> が 0 以下です。</exception>
        public CellFontAttribute(string familyName, float size, FontStyle fontStyle) : this(familyName, fontStyle)
        {
            if (size <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(size));
            }

            Size = size;
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="familyName">フォント名。</param>
        /// <param name="size">フォントサイズ。</param>
        /// <param name="fontStyle">フォントスタイル。</param>
        /// <param name="color">Red や #FF0000 などのHTML色表現で指定する色。</param>
        public CellFontAttribute(string familyName, float size, FontStyle fontStyle, string color) : this(familyName, size, fontStyle)
        {
            Color = ColorTranslator.FromHtml(color);
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="size">フォントサイズ。</param>
        public CellFontAttribute(float size)
        {
            if (size <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(size));
            }

            Size = size;
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="color">Red や #FF0000 などのHTML色表現で指定する色。</param>
        public CellFontAttribute(string color)
        {
            Color = ColorTranslator.FromHtml(color);
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="red">RGB形式で指定する赤色。</param>
        /// <param name="green">RGB形式で指定する緑色。</param>
        /// <param name="blue">RGB形式で指定する青色。</param>
        public CellFontAttribute(int red, int green, int blue)
        {
            Color = Color.FromArgb(red, green, blue);
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="size">フォントサイズ。</param>
        /// <param name="color">Red や #FF0000 などのHTML色表現で指定する色。</param>
        public CellFontAttribute(float size, string color) : this(size)
        {
            Color = ColorTranslator.FromHtml(color);
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="size">フォントサイズ。</param>
        /// <param name="red">RGB形式で指定する赤色。</param>
        /// <param name="green">RGB形式で指定する緑色。</param>
        /// <param name="blue">RGB形式で指定する青色。</param>
        public CellFontAttribute(float size, int red, int green, int blue) : this(size)
        {
            Color = Color.FromArgb(red, green, blue);
        }
    }
}
