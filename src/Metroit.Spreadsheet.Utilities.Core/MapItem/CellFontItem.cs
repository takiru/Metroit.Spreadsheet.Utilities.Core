using System.Drawing;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルのフォント情報を提供します。
    /// </summary>
    public class CellFontItem
    {
        /// <summary>
        /// フォント を指定します。
        /// </summary>
        public string FontFamily { get; set; }

        /// <summary>
        /// フォントスタイル を指定します。
        /// </summary>
        public FontStyle? FontStyle { get; set; }

        /// <summary>
        /// フォントサイズ を指定します。
        /// </summary>
        public float? Size { get; set; }

        /// <summary>
        /// フォント色 を指定します。
        /// </summary>
        public Color Color { get; set; }

        /// <summary>
        /// 新しい CellFontItem インスタンスを生成します。
        /// </summary>
        /// <param name="font">フォント。</param>
        /// <param name="fontStyle">フォントスタイル。</param>
        /// <param name="size">フォントサイズ。</param>
        /// <param name="color">フォント色。</param>
        public CellFontItem(Font font, FontStyle? fontStyle, float? size, Color color)
        {
            FontFamily = font.FontFamily.Name;
            FontStyle = fontStyle;
            Size = size;
            Color = color;
        }
    }
}
