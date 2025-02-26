using Metroit.Spreadsheet.Utilities.Core.Mapping;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルの結合セル情報を提供します。
    /// </summary>
    public class CellMergeItem
    {
        /// <summary>
        /// 終了行インデックスまたは相対値 を指定します。
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// 終了列インデックスまたは相対値 を指定します。
        /// </summary>
        public int Column { get; set; }

        /// <summary>
        /// 結合を行う配置方法 を指定します。
        /// </summary>
        public CellMergePosition Position { get; set; }

        /// <summary>
        /// 新しい CellMergeItem インスタンスを生成します。
        /// </summary>
        public CellMergeItem(int row, int column, CellMergePosition position)
        {
            Row = row;
            Column = column;
            Position = position;
        }
    }
}
