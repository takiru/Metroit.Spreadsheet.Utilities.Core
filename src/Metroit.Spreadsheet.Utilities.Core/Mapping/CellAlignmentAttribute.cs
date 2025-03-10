using System;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// Excelセルの表示情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CellAlignmentAttribute : Attribute
    {
        /// <summary>
        /// 横位置 を指定します。既定は <see cref="HorizontalAlignment.General"/> です。
        /// </summary>
        public HorizontalAlignment Horizontal { get; } = HorizontalAlignment.General;

        /// <summary>
        /// 縦位置 を指定します。既定は <see cref="VerticalAlignment.Center"/> です。
        /// </summary>
        public VerticalAlignment Vertical { get; } = VerticalAlignment.Center;

        /// <summary>
        /// 新しい CellAlignmentAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="horizontal">横位置。</param>
        public CellAlignmentAttribute(HorizontalAlignment horizontal)
        {
            Horizontal = horizontal;
        }

        /// <summary>
        /// 新しい CellAlignmentAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="vertical">縦位置。</param>
        public CellAlignmentAttribute(VerticalAlignment vertical)
        {
            Vertical = vertical;
        }

        /// <summary>
        /// 新しい CellAlignmentAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="horizontal">横位置。</param>
        /// <param name="vertical">縦位置。</param>
        public CellAlignmentAttribute(HorizontalAlignment horizontal, VerticalAlignment vertical) : this(horizontal)
        {
            Vertical = vertical;
        }
    }
}
