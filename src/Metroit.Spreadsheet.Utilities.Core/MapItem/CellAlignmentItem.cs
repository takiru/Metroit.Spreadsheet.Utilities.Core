using Metroit.Spreadsheet.Utilities.Core.Mapping;
using System;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルの表示位置情報を提供します。
    /// </summary>
    public class CellAlignmentItem : Attribute
    {
        /// <summary>
        /// 水平位置を取得します。
        /// </summary>
        public HorizontalAlignment Horizontal { get; }

        /// <summary>
        /// 垂直位置を取得します。
        /// </summary>
        public VerticalAlignment Vertical { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="horizontal">水平位置。</param>
        /// <param name="vertical">垂直位置。</param>
        public CellAlignmentItem(HorizontalAlignment horizontal, VerticalAlignment vertical)
        {
            Horizontal = horizontal;
            Vertical = vertical;
        }
    }
}
