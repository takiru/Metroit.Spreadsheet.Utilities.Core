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
        /// 横位置 を指定します。
        /// </summary>
        public HorizontalAlignment? Horizontal { get; set; }

        /// <summary>
        /// 縦位置 を指定します。
        /// </summary>
        public VerticalAlignment? Vertical { get; set; }

        /// <summary>
        /// 新しい CellAlignmentItem インスタンスを生成します。
        /// </summary>
        /// <param name="horizontal">横位置。</param>
        /// <param name="vertical">縦位置。</param>
        public CellAlignmentItem(HorizontalAlignment? horizontal, VerticalAlignment? vertical)
        {
            Horizontal = horizontal;
            Vertical = vertical;
        }
    }
}
