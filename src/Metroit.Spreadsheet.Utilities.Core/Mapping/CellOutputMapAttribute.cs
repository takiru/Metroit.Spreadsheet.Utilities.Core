using System;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// Excelセルの出力セルインデックス情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CellOutputMapAttribute : Attribute
    {
        /// <summary>
        /// 行インデックスを指定します。既定は-1です。
        /// </summary>
        public int Row { get; set; } = -1;

        /// <summary>
        /// 列インデックスを指定します。既定は-1です。
        /// </summary>
        public int Column { get; set; } = -1;

        /// <summary>
        /// 数式かどうかを指定します。既定はfalseです。
        /// </summary>
        public bool Formula { get; set; } = false;

        /// <summary>
        /// 新しい OutputMapCellAttribute インスタンスを生成します。
        /// </summary>
        public CellOutputMapAttribute() { }
    }
}
