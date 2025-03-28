using System;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// スプレッドシートのセルから読み込みを行う情報を提供します。
    /// </summary>
    //public class CellInputMapItem : IReadOnlyCellMapItem
    public class CellInputMapItem
    {
        /// <summary>
        /// 読み込みを行うプロパティ名を取得します。
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// オリジナルの行インデックスを取得します。
        /// </summary>
        public int OriginalRow { get; }

        /// <summary>
        /// オリジナルの列インデックスを取得します。
        /// </summary>
        public int OriginalColumn { get; }

        /// <summary>
        /// 読み込みを行う行インデックスを取得します。
        /// </summary>
        public int Row { get; private set; }

        /// <summary>
        /// 読み込みを行う列インデックスを取得します。
        /// </summary>
        public int Column { get; private set; }

        /// <summary>
        /// 数式かどうかを取得します。
        /// </summary>
        public bool IsFormula { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="name">読み込みを行うプロパティ名。</param>
        /// <param name="row">0 から始まる行インデックス、または未指定を表す - 1。</param>
        /// <param name="column">0 から始まる列インデックス、または未指定を表す - 1。</param>
        /// <param name="isFormula">数式のセルの場合は true, それ以外なら false。</param>
        public CellInputMapItem(string name, int row, int column, bool isFormula)
        {
            Name = name;
            OriginalRow = row;
            Row = row;
            OriginalColumn = column;
            Column = column;
            IsFormula = isFormula;
        }

        /// <summary>
        /// 読み込みを行う行インデックスを変更します。
        /// </summary>
        /// <param name="row">読み込みを行う行インデックス。</param>
        /// <exception cref="ArgumentOutOfRangeException">行は 0 以上である必要があります。</exception>
        public void ChangeRow(int row)
        {
            if (row < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(Row), row, "Row must be greater than or equal to 0.");
            }
            Row = row;
        }

        /// <summary>
        /// 読み込みを行う列インデックスを変更します。
        /// </summary>
        /// <param name="column">読み込みを行う列インデックス。</param>
        public void ChangeColumn(int column)
        {
            if (column < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(Column), column, "Column must be greater than or equal to 0.");
            }
            Column = column;
        }
    }
}
