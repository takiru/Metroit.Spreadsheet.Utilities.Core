using System.Drawing;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルのフォント情報を提供します。
    /// </summary>
    public class CellFontItem
    {
        /// <summary>
        /// フォント名を取得します。
        /// </summary>
        public string FamilyName { get; }

        /// <summary>
        /// フォントサイズを取得します。
        /// </summary>
        public float Size { get; }

        /// <summary>
        /// フォントスタイルを取得します。
        /// </summary>
        public FontStyle FontStyle { get; }

        /// <summary>
        /// フォント色を取得します。
        /// </summary>
        public Color Color { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="familyName">フォント名。</param>
        /// <param name="size">フォントサイズ。</param>
        /// <param name="fontStyle">フォントスタイル。</param>
        /// <param name="color">フォント色。</param>
        public CellFontItem(string familyName, float size, FontStyle fontStyle, Color color)
        {
            FamilyName = familyName;
            Size = size;
            FontStyle = fontStyle;
            Color = color;
        }
    }
}
