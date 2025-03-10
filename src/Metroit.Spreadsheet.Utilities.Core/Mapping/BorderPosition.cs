namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// 罫線を引く位置を提供します。
    /// </summary>
    public enum BorderPosition
    {
        /// <summary>
        /// 指定なし を示します。
        /// </summary>
        None,

        /// <summary>
        /// 右下がり斜線 を示します。
        /// </summary>
        DiagonalDown,

        /// <summary>
        /// 右上がり斜線 を示します。
        /// </summary>
        DiagonalUp,

        /// <summary>
        /// 下線　を示します。
        /// </summary>
        EdgeBottom,

        /// <summary>
        /// 左線 を示します。
        /// </summary>
        EdgeLeft,

        /// <summary>
        /// 右線 を示します。
        /// </summary>
        EdgeRight,

        /// <summary>
        /// 上線 を示します。
        /// </summary>
        EdgeTop,

        /// <summary>
        /// 内側横線 を示します。
        /// </summary>
        InsideHorizontal,
        
        /// <summary>
        /// 内側縦線 を示します。
        /// </summary>
        InsideVertical
    }
}
