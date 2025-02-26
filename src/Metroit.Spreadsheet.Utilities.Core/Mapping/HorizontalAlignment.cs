namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// セル配置を行う横位置を提供します。
    /// </summary>
    public enum HorizontalAlignment
    {
        /// <summary>
        /// 標準 を示します。
        /// </summary>
        General,

        /// <summary>
        /// 左詰め を示します。
        /// </summary>
        Left,

        /// <summary>
        /// 中央揃え を示します。
        /// </summary>
        Center,

        /// <summary>
        /// 右詰め を示します。
        /// </summary>
        Right,

        /// <summary>
        /// 繰り返し を示します。
        /// </summary>
        Fill,

        /// <summary>
        /// 両端揃え を示します。
        /// </summary>
        Justify,

        /// <summary>
        /// 選択範囲内で中央 を示します。
        /// </summary>
        CenterContinuous,

        /// <summary>
        /// 均等割り付け を示します。
        /// </summary>
        Distributed
    }
}
