using System;
using System.Drawing;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// Excelセルの背景情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CellBackgroundAttribute : Attribute
    {
        /// <summary>
        /// 背景色 を指定します。既定は null です。
        /// </summary>
        public string Color { get; set; } = null;

        /// <summary>
        /// 新しい CellBackgroundAttribute インスタンスを生成します。
        /// </summary>
        public CellBackgroundAttribute(string color)
        {
            Color = color;
        }

        /// <summary>
        /// フォント色 を取得します。
        /// </summary>
        /// <returns>フォント色。</returns>
        public Color GetColor()
        {
            try
            {
                return ColorTranslator.FromHtml(Color);
            }
            catch
            {
                return System.Drawing.Color.Empty;
            }
        }
    }
}
