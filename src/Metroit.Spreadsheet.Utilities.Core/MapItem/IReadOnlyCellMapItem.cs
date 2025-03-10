namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// スプレッドシートのセル情報を読み取り専用で提供します。
    /// </summary>
    public interface IReadOnlyCellMapItem
    {
        /// <summary>
        /// プロパティ名を取得します。
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
        /// 開始行インデックスを取得します。
        /// </summary>
        int StartRow { get; }

        /// <summary>
        /// 開始列インデックスを取得します。
        /// </summary>
        int StartColumn { get; }

        /// <summary>
        /// 終了行インデックスを取得します。
        /// </summary>
        int EndRow { get; }

        /// <summary>
        /// 終了列インデックスを取得します。
        /// </summary>
        int EndColumn { get; }

        /// <summary>
        /// 数式かどうかを取得します。
        /// </summary>
        bool IsFormula { get; }
    }
}
