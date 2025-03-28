using System;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// スプレッドシートのセルに書き出しを行う情報を提供します。
    /// </summary>
    public class CellOutputMapItem : IReadOnlyCellMapItem
    {
        /// <summary>
        /// 書き出しを行うプロパティ名を取得します。
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// オリジナルの行インデックスを取得します。未指定の場合、-1 を取得します。
        /// </summary>
        public int OriginalRow { get; }

        /// <summary>
        /// オリジナルの列インデックスを取得します。未指定の場合、-1 を取得します。
        /// </summary>
        public int OriginalColumn { get; }

        /// <summary>
        /// 書き出しを行う開始行インデックスを取得します。
        /// </summary>
        public int StartRow { get; private set; }

        /// <summary>
        /// 書き出しを行う開始列インデックスを取得します。
        /// </summary>
        public int StartColumn { get; private set; }

        /// <summary>
        /// 書き出しを行う終了行インデックスを取得します。
        /// </summary>
        public int EndRow { get; private set; }

        /// <summary>
        /// 書き出しを行う終了列インデックスを取得します。
        /// </summary>
        public int EndColumn { get; private set; }

        /// <summary>
        /// 数式かどうかを取得します。
        /// </summary>
        public bool IsFormula { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="name">書き出しを行うプロパティ名。</param>
        /// <param name="row">0 から始まる行インデックス、または未指定を表す - 1。</param>
        /// <param name="column">0 から始まる列インデックス、または未指定を表す - 1。</param>
        /// <param name="isFormula">数式のセルの場合は true, それ以外なら false。</param>
        public CellOutputMapItem(string name, int row, int column, bool isFormula)
        {
            Name = name;
            OriginalRow = row;
            StartRow = row;
            EndRow = row;
            OriginalColumn = column;
            StartColumn = column;
            EndColumn = column;
            IsFormula = isFormula;
        }

        /// <summary>
        /// 書き出しを行う行インデックスを変更します。
        /// </summary>
        /// <param name="row">書き出しを行う行インデックス。</param>
        /// <exception cref="ArgumentOutOfRangeException">行は 0 以上である必要があります。</exception>
        public void ChangeRow(int row)
        {
            if (row < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(StartRow), row, "Row must be greater than or equal to 0.");
            }
            StartRow = row;
            EndRow = row;
        }

        /// <summary>
        /// 書き出しを行う列インデックスを変更します。
        /// </summary>
        /// <param name="column">書き出しを行う列インデックス。</param>
        public void ChangeColumn(int column)
        {
            if (column < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(StartColumn), column, "Column must be greater than or equal to 0.");
            }
            StartColumn = column;
            EndColumn = column;
        }

        /// <summary>
        /// セル結合に伴うセル範囲に変更します。
        /// </summary>
        /// <param name="startRow">開始行インデックス。</param>
        /// <param name="startColumn">開始列インデックス。</param>
        /// <param name="endRow">終了行インデックス。</param>
        /// <param name="endColumn">終了列インデックス。</param>
        internal void ChangeCell(int startRow, int startColumn, int endRow, int endColumn)
        {
            StartRow = startRow;
            StartColumn = startColumn;
            EndRow = endRow;
            EndColumn = endColumn;
        }
    }
}
