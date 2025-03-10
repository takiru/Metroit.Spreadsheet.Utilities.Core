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
        /// 背景色 を指定します。
        /// </summary>
        public Color Color { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="color">Red や #FF0000 などのHTML色表現で指定する色。</param>
        public CellBackgroundAttribute(string color)
        {
            Color = ColorTranslator.FromHtml(color);
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="red">RGB形式で指定する赤色。</param>
        /// <param name="green">RGB形式で指定する緑色。</param>
        /// <param name="blue">RGB形式で指定する青色。</param>
        public CellBackgroundAttribute(int red, int green, int blue)
        {
            Color = Color.FromArgb(red, green, blue);
        }
    }
}
