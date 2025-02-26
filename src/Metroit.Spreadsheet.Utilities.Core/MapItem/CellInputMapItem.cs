namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルの入力セルインデックス情報を提供します。
    /// </summary>
    public class CellInputMapItem
    {
        /// <summary>
        /// 行インデックスを指定します。
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// 列インデックスを指定します。
        /// </summary>
        public int Column { get; set; }

        /// <summary>
        /// 数式かどうかを指定します。
        /// </summary>
        public bool IsFormula { get; }

        /// <summary>
        /// 新しい CellInputMapItem インスタンスを生成します。
        /// </summary>
        public CellInputMapItem(int row, int column, bool isFormula)
        {
            Row = row;
            Column = column;
            IsFormula = isFormula;
        }
    }
}
