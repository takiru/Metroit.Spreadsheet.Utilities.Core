using System;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// Excelセルの結合セル情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CellMergeAttribute : Attribute
    {
        /// <summary>
        /// 終了行インデックスまたは相対値 を指定します。
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// 終了列インデックスまたは相対値 を指定します。
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// 結合を行う配置方法 を指定します。既定は Relative です。
        /// </summary>
        public CellMergePosition Position { get; } = CellMergePosition.Relative;

        /// <summary>
        /// 新しい CellMergeAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="row">終了行インデックスまたは相対値。</param>
        /// <param name="column">終了列インデックスまたは相対値。</param>
        public CellMergeAttribute(int row, int column)
        {
            Row = row;
            Column = column;
        }

        /// <summary>
        /// 新しい CellMergeAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="row">終了行インデックスまたは相対値。position を <see cref="CellMergePosition.Absolute"/> とした時、0 以上でなければなりません。</param>
        /// <param name="column">終了列インデックスまたは相対値。position を <see cref="CellMergePosition.Absolute"/> とした時、0 以上でなければなりません。</param>
        /// <param name="position">結合を行う配置方法。</param>
        public CellMergeAttribute(int row, int column, CellMergePosition position)
        {
            if (position == CellMergePosition.Absolute)
            {
                if (row < 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(row));
                }
                if (column < 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(column));
                }
            }

            Row = row;
            Column = column;
            Position = position;
        }
    }
}
