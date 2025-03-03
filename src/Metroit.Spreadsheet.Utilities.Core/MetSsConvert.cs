using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace Metroit.Spreadsheet.Utilities.Core
{
    /// <summary>
    /// セル参照の変換操作を提供します。
    /// </summary>
    public static class MetSsConvert
    {
        /// <summary>
        /// R1C1形式の値からA1形式の文字列に変換します。<br/>
        /// 行インデックスおよび列インデックスは、ソフトウェアに依存する最大インデックスを考慮しません。
        /// </summary>
        /// <param name="row">行インデックス。</param>
        /// <param name="column">列インデックス。</param>
        /// <returns>A1形式の文字列。</returns>
        /// <exception cref="ArgumentOutOfRangeException">row または column が 0 未満の場合に発生します。</exception>
        /// <remarks>
        /// ex.) row = 0, column = 0 の場合、A1 を返却します。<br/>
        /// row および column は 0 以上でなければなりません。
        /// </remarks>
        public static string RcToA1(int row, int column)
        {
            if (row < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(row));
            }
            if (column < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(column));
            }

            return $"{ToColumnA1(column)}{row + 1}";
        }

        /// <summary>
        /// 列インデックスからA1形式の列文字列に変換します。<br/>
        /// 列インデックスは、ソフトウェアに依存する最大インデックスを考慮しません。
        /// </summary>
        /// <param name="index">列インデックス。</param>
        /// <returns>A1形式の列文字列。</returns>
        /// <exception cref="ArgumentOutOfRangeException">index が 0 未満の場合に発生します。</exception>
        /// <remarks>
        /// ex.) index = 0 の場合、A を返却し、 index = 26 の場合、AA を返却します。<br/>
        /// index は 0 以上でなければなりません。
        /// </remarks>
        public static string ToColumnA1(int index)
        {
            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            if (index - 26 < 0)
            {
                return char.ToString((char)('A' + (index % 26)));
            }

            return ToColumnA1((index / 26) - 1) + (char)('A' + (index % 26));
        }

        /// <summary>
        /// A1形式の列文字列から列インデックスに変換します。<br/>
        /// 列文字列は、A から FXSHRXX までの範囲を受け付けます。
        /// </summary>
        /// <param name="value">A1形式の列文字列。</param>
        /// <returns>列インデックス。</returns>
        /// <exception cref="ArgumentNullException">value が null の場合に発生します。</exception>
        /// <exception cref="ArgumentException">value が空の場合に発生します。</exception>
        /// <exception cref="FormatException">value に英字以外が含まれる場合に発生します。</exception>
        /// <exception cref="ArgumentOutOfRangeException">value が <see cref="int.MaxValue"/> で表現可能な FXSHRXX を超過している場合に発生します。</exception>
        /// <remarks>
        /// ex.) value = A の場合、0 を返却し、 value = AA の場合、26 を返却します。<br/>
        /// value は A から FXSHRXX までの範囲の英字のみで構成されなければなりません。
        /// </remarks>
        public static int ToColumnIndex(string value)
        {
            // int.MaxValue に適合する列文字列
            const string MaxValue = "FXSHRXX";

            // 値が未設定の場合は実行不可
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("value is empty.");
            }

            // 英字以外が含まれる場合は実行不可
            string upperValue = value.ToUpper();
            if (!Regex.IsMatch(upperValue, @"^[A-Z]+$"))
            {
                throw new FormatException("Argument format is only alphabet");
            }

            // 変換後がint.MaxValueに収まらない場合は実行不可
            if (upperValue.Length > MaxValue.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(value), $"Argument range max \"{MaxValue}\"");
            }
            if (upperValue.Length == MaxValue.Length && upperValue.CompareTo(MaxValue) >= 1)
            {
                throw new ArgumentOutOfRangeException(nameof(value), $"Argument range max \"{MaxValue}\"");
            }

            return CalculateColumnIndex(upperValue, 0);
        }

        /// <summary>
        /// A1形式の列文字列から列インデックスを算出する。
        /// 下位にある文字列から順に計算を行い、インデックス値を求める。
        /// </summary>
        /// <param name="value">A1形式の列文字列。</param>
        /// <param name="callNum">0 から始まる呼び出し回数。</param>
        private static int CalculateColumnIndex(string value, int callNum)
        {
            var remain = value.Last() - 'A';
            var charIndex = 0;
            if (callNum > 0)
            {
                var pow = (int)Math.Pow(26, callNum);
                charIndex = pow + remain * pow;
            }
            else
            {
                charIndex += remain;
            }

            if (value.Length == 1)
            {
                // 最上位文字のインデックスを算出したら処理を終了する
                return charIndex;
            }

            return charIndex + CalculateColumnIndex(value.Substring(0, value.Length - 1), ++callNum);
        }
    }
}
