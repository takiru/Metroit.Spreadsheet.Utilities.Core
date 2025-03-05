namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// スプレッドシートのセル情報を読み取り専用で提供します。
    /// </summary>
    public interface IReadOnlyCellMapItem
    {
        /// <summary>
        /// 読み込みを行うプロパティ名を取得します。
        /// </summary>
        string Name { get; }

        /// <summary>
        /// オリジナルの行インデックスを取得します。
        /// </summary>
        int OriginalRow { get; }

        /// <summary>
        /// オリジナルの列インデックスを取得します。
        /// </summary>
        int OriginalColumn { get; }

        /// <summary>
        /// 読み込みを行う行インデックスを取得します。
        /// </summary>
        int Row { get; }

        /// <summary>
        /// 読み込みを行う列インデックスを取得します。
        /// </summary>
        int Column { get; }

        /// <summary>
        /// 数式かどうかを取得します。
        /// </summary>
        bool IsFormula { get; }
    }
}
