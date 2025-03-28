using System;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// マッピングする位置をシフトする情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class MapShiftAttribute : Attribute
    {
        /// <summary>
        /// 現在のマッピング位置からシフトする値を取得します。既定は 1 です。
        /// </summary>
        public int ShiftCount { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        public MapShiftAttribute() : this(1) { }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="shiftCount">現在のマッピング位置からシフトする値。</param>
        public MapShiftAttribute(int shiftCount)
        {
            ShiftCount = shiftCount;
        }
    }
}
