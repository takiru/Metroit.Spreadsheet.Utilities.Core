using System;
using System.Linq;

namespace Metroit.Spreadsheet.Utilities.Core
{
    /// <summary>
    /// セル参照の操作を提供します。
    /// </summary>
    public static class CellUtils
    {
        /// <summary>
        /// R1C1形式の値からA1形式の文字列に変換します。
        /// </summary>
        /// <param name="row">行インデックス。</param>
        /// <param name="column">列インデックス。</param>
        /// <returns>A1形式の文字列。</returns>
        public static string RcToA1(int row, int column)
        {
            return ToColumnName(column + 1) + (row + 1).ToString();
        }

        /// <summary>
        /// 数値形式の列インデックスを、A1形式の列インデックスに変換します。
        /// </summary>
        /// <param name="index">列インデックス。</param>
        /// <returns>A1形式の列インデックス。</returns>
        public static string ToColumnName(int index)
        {
            if (index < 1)
            {
                return string.Empty;
            }
            return ToColumnName((index - 1) / 26) + (char)('A' + ((index - 1) % 26));
        }

        /// <summary>
        /// Excel A1形式の列インデックスを、数値形式の列インデックスに変換します。
        /// </summary>
        /// <param name="value">Excel A1形式の列インデックス。</param>
        /// <returns>列インデックス。</returns>
        public static int ToColumnIndex(string value)
        {
            // 値が未設定の場合は実行不可
            if (value is null || value.Trim() == "")
            {
                throw new ArgumentException("Argument format is only alphabet");
            }

            // 英字以外が含まれる場合は実行不可
            string upperValue = value.ToUpper();
            if (!System.Text.RegularExpressions.Regex.IsMatch(upperValue, @"^[A-Z]+$"))
            {
                throw new FormatException("Argument format is only alphabet");
            }

            // 変換後がint.MaxValueに収まらない場合は実行不可
            if (upperValue.Length == "FXSHRXW".Length && upperValue.CompareTo("FXSHRXW") >= 1)
            {
                throw new ArgumentOutOfRangeException("Argument range max \"FXSHRXW\"");
            }

            return ToColumnNumber(upperValue, 0);
        }

        /// <summary>
        /// Excel A1形式の文字列を数値に変換する。
        /// </summary>
        /// <param name="value">Excel A1形式の文字列。</param>
        /// <param name="callNum">呼び出し回数</param>
        private static int ToColumnNumber(string value, int callNum)
        {
            if (value == "")
            {
                return 0;
            }

            int digit = (int)Math.Pow(26, callNum);
            return ((value.Last() - 'A' + 1) * digit) + ToColumnNumber(value.Substring(0, value.Length - 1), ++callNum);
        }
    }
}
