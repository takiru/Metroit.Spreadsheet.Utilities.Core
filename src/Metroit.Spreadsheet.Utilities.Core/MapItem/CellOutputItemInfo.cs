using System.Collections.Generic;
using System.Reflection;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// 出力マッピングされたプロパティについての情報を提供します。
    /// </summary>
    public class CellOutputItemInfo
    {
        /// <summary>
        /// マッピングされたプロパティを取得します。
        /// </summary>
        public PropertyInfo MapProperty { get; }

        /// <summary>
        /// 出力セルインデックス情報を取得します。
        /// </summary>
        public CellOutputMapItem Cell { get; }

        /// <summary>
        /// 結合情報を取得します。
        /// </summary>
        public CellMergeItem CellMerge { get; }

        /// <summary>
        /// 書式情報を取得します。
        /// </summary>
        public CellFormatItem CellFormat { get; }

        /// <summary>
        /// フォント情報を取得します。
        /// </summary>
        public CellFontItem CellFont { get; }

        /// <summary>
        /// 文字列の修飾情報を取得します。
        /// </summary>
        public CellCharacterDecorationItem CellCharacterDecoration { get; }

        /// <summary>
        /// 表示位置情報を取得します。
        /// </summary>
        public CellAlignmentItem CellAlignment { get; }

        /// <summary>
        /// 背景情報を取得します。
        /// </summary>
        public CellBackgroundItem CellBackground { get; }

        /// <summary>
        /// 罫線情報を取得します。
        /// </summary>
        public List<CellBorderItem> CellBorders { get; }

        /// <summary>
        /// 新しい CellOutputItemInfo インスタンスを生成します。
        /// </summary>
        /// <param name="pi">マッピングされたプロパティ。</param>
        /// <param name="cell">出力セルインデックス情報。</param>
        /// <param name="cellMerge">結合情報。</param>
        /// <param name="cellFormat">書式情報。</param>
        /// <param name="cellFont">フォント情報。</param>
        /// <param name="cellCharacterDecoration">文字列の修飾情報。</param>
        /// <param name="cellAlignment">表示位置情報。</param>
        /// <param name="cellBackground">背景情報。</param>
        /// <param name="cellBorders">罫線情報。</param>
        public CellOutputItemInfo(PropertyInfo pi, CellOutputMapItem cell, CellMergeItem cellMerge,
            CellFormatItem cellFormat, CellFontItem cellFont, CellCharacterDecorationItem cellCharacterDecoration,
            CellAlignmentItem cellAlignment, CellBackgroundItem cellBackground, List<CellBorderItem> cellBorders)
        {
            MapProperty = pi;
            Cell = cell;
            CellMerge = cellMerge;
            CellFormat = cellFormat;
            CellFont = cellFont;
            CellCharacterDecoration = cellCharacterDecoration;
            CellAlignment = cellAlignment;
            CellBackground = cellBackground;
            CellBorders = cellBorders;
        }
    }
}
