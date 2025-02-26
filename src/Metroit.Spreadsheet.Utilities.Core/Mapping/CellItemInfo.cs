using System.Reflection;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// マッピングされたプロパティについての情報を提供します。
    /// </summary>
    public class CellItemInfo
    {
        /// <summary>
        /// マッピングされたプロパティを取得します。
        /// </summary>
        public PropertyInfo MapProperty { get; }

        /// <summary>
        /// マッピングされた行インデックスを取得または設定します。
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// マッピングされた列インデックスを取得または設定します。
        /// </summary>
        public int Column { get; set; }

        /// <summary>
        /// マッピングされた読み込みまたは書き込み時、数式を利用するかどうかを取得します。
        /// </summary>
        public bool IsFormula { get; }

        /// <summary>
        /// 新しい CellItemInfo インスタンスを生成します。
        /// </summary>
        /// <param name="pi">マッピングされたプロパティ。</param>
        /// <param name="isFormula">数式を利用するかどうか。</param>
        public CellItemInfo(PropertyInfo pi, bool isFormula)
        {
            MapProperty = pi;
            IsFormula = isFormula;
        }
    }
}
