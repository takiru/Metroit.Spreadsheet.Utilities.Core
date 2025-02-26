using System;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// Excelセルの出力セル書式情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CellFormatAttribute : Attribute
    {
        /// <summary>
        /// 書式を指定します。
        /// </summary>
        public string Format { get; }

        /// <summary>
        /// 新しい CellFormatAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="format">書式。</param>
        public CellFormatAttribute(string format)
        {
            Format = format;
        }
    }
}
