using System.Reflection;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// 入力マッピングされたプロパティについての情報を提供します。
    /// </summary>
    public class CellInputItemInfo
    {
        /// <summary>
        /// マッピングされたプロパティを取得します。
        /// </summary>
        public PropertyInfo MapProperty { get; }

        /// <summary>
        /// 入力セルインデックス情報を取得します。
        /// </summary>
        public CellInputMapItem Cell { get; }

        /// <summary>
        /// 新しい CellInputItemInfo インスタンスを生成します。
        /// </summary>
        /// <param name="pi">マッピングされたプロパティ。</param>
        /// <param name="cell">入力セルインデックス情報。</param>
        public CellInputItemInfo(PropertyInfo pi, CellInputMapItem cell)
        {
            MapProperty = pi;
            Cell = cell;
        }
    }
}
