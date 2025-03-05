using Metroit.Spreadsheet.Utilities.Core.Mapping;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// スプレッドシートのセルの結合セル情報を提供します。
    /// </summary>
    public class CellMergeItem
    {
        /// <summary>
        /// 終了行インデックスまたは相対値を取得します。
        /// </summary>
        public int OriginalRow { get; }

        /// <summary>
        /// 終了列インデックスまたは相対値を取得します。
        /// </summary>
        public int OriginalColumn { get; }

        /// <summary>
        /// 結合を行う配置方法を取得します。
        /// </summary>
        public CellMergePosition Position { get; }

        /// <summary>
        /// 結合範囲となる終了行インデックスを取得します。
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// 結合範囲となる終了列インデックスを取得します。
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        public CellMergeItem(int originalRow, int originalColumn, CellMergePosition position, int row, int column)
        {
            OriginalRow = originalRow;
            Row = row;
            OriginalColumn = originalColumn;
            Column = column;
            Position = position;
        }
    }
}
