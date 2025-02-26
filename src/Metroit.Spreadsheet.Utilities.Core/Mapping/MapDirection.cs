namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// マップするインデックスの方向を定義します。
    /// </summary>
    public enum MapDirection
    {
        /// <summary>
        /// 行へマップすることを定義します。
        /// </summary>
        Row,

        /// <summary>
        /// 列へマップすることを定義します。
        /// </summary>
        Column,

        /// <summary>
        /// マップする行、列の制御を行わないことを定義します。
        /// </summary>
        None
    }
}
