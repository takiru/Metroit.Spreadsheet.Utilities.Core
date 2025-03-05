using Metroit.Spreadsheet.Utilities.Core.Mapping;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Metroit.Spreadsheet.Utilities.Core
{
    /// <summary>
    /// セル参照の変換操作を提供します。
    /// </summary>
    public static class MetSpreadsheetConvert
    {
        /// <summary>
        /// セルインデックス値からA1形式の文字列に変換します。<br/>
        /// 行の指定、列の指定がそれぞれない場合、列または行のみを求めます。<br />
        /// 行インデックスおよび列インデックスは、ソフトウェアに依存する最大インデックスを考慮しません。
        /// </summary>
        /// <param name="row">行インデックス。行の指定がない場合は -1 を指定します。</param>
        /// <param name="column">列インデックス。列の指定がない場合は -1 を指定します。</param>
        /// <returns>A1形式の文字列。</returns>
        /// <exception cref="ArgumentException">row および column が 0 未満の場合に発生します。</exception>
        /// <remarks>
        /// ex.)
        /// <code>
        ///     var value = MetSpreadsheetConvert.RcToA1(0, 0);     // A1
        ///     var value = MetSpreadsheetConvert.RcToA1(-1, 0);    // A
        ///     var value = MetSpreadsheetConvert.RcToA1(0, -1);    // 1
        /// </code>
        /// row または column は 0 以上でなければなりません。
        /// </remarks>
        public static string ToA1(int row, int column)
        {
            if (row < 0 && column < 0)
            {
                throw new ArgumentException("Either row or column must be 0 or greater.");
            }

            if (row < 0)
            {
                return $"{ToColumnA1(column)}";
            }
            if (column < 0)
            {
                return $"{row + 1}";
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
        /// ex.)
        /// <code>
        ///     var value = MetSpreadsheetConvert.ToColumnA1(0);    // A
        ///     var value = MetSpreadsheetConvert.ToColumnA1(26);   // AA
        /// </code>
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
        /// ex.)
        /// <code>
        ///     var value = MetSpreadsheetConvert.ToColumnIndex("A");   // 0
        ///     var value = MetSpreadsheetConvert.ToColumnIndex("AA");  // 26
        /// </code>
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
        /// 指定したセル範囲をA1形式に変換します。<br/>
        /// 行の指定、列の指定がそれぞれない場合、列または行のみを求めます。<br />
        /// 範囲が同一の場合、対象セルをA1形式で求めます。<br />
        /// row1, column1 のセル位置が row2, column2 のセル位置より後ろに位置する場合、求められる範囲文字列は、前に位置するセル位置から始まります。
        /// </summary>
        /// <param name="row1">0 から始まる開始行インデックス。行の指定がない場合は -1 を指定します。</param>
        /// <param name="column1">0 から始まる開始列インデックス。列の指定がない場合は -1 を指定します。</param>
        /// <param name="row2">0 から始まる終了行インデックス。行の指定がない場合は -1 を指定します。</param>
        /// <param name="column2">0 から始まる終了列インデックス。列の指定がない場合は -1 を指定します。</param>
        /// <returns>A1形式の範囲文字列。</returns>
        /// <exception cref="ArgumentException">row1 および column1, または row2 および column2  が 0 未満の場合に発生します。</exception>
        /// <remarks>
        /// ex.)
        /// <code>
        ///     var value = MetSpreadsheetConvert.Range(0, 0, 1, 1);    // A1:B2
        ///     var value = MetSpreadsheetConvert.Range(1, 0, 0, 0);    // A1:A2
        ///     var value = MetSpreadsheetConvert.Range(0, 1, 0, 0);    // A1:B1
        ///     var value = MetSpreadsheetConvert.Range(1, 1, 0, 0);    // A1:B2
        ///     var value = MetSpreadsheetConvert.Range(0, -1, 0, -1);  // 1
        ///     var value = MetSpreadsheetConvert.Range(0, -1, 1, -1);  // 1:2
        ///     var value = MetSpreadsheetConvert.Range(-1, 0, -1, 0);  // A
        ///     var value = MetSpreadsheetConvert.Range(-1, 0, -1, 1);  // A:B
        ///     var value = MetSpreadsheetConvert.Range(0, 0, 0, 0);    // A1
        /// </code>
        /// row1 または column1 は 0 以上でなければなりません。<br/>
        /// row2 または column2 は 0 以上でなければなりません。<br/>
        /// row1 が 0 以上の時、 row2 は 0 以上でなければなりません。<br/>
        /// row2 が 0 以上の時、 row1 は 0 以上でなければなりません。<br/>
        /// column1 が 0 以上の時、 column2 は 0 以上でなければなりません。<br/>
        /// column2 が 0 以上の時、 column1 は 0 以上でなければなりません。
        /// </remarks>
        public static string ToRange(int row1, int column1, int row2, int column2)
        {
            if ((row1 >= 0 && row2 < 0) || (row2 >= 0 && row1 < 0))
            {
                throw new ArgumentException("row must be specified.");
            }
            if ((column1 >= 0 && column2 < 0) || (column2 >= 0 && column1 < 0))
            {
                throw new ArgumentException("column must be specified.");
            }

            var cell1 = ToA1(row1, column1);
            var cell2 = ToA1(row2, column2);

            if (cell1 == cell2)
            {
                return cell1;
            }
            if (row1 > row2)
            {
                return $"{cell2}:{cell1}";
            }
            if (column1 > column2)
            {
                return $"{cell2}:{cell1}";
            }

            return $"{cell1}:{cell2}";
        }

        /// <summary>
        /// A1:A2 で表現されるA1形式の検証。
        /// </summary>
        private static readonly Regex FullRangeRegex = new Regex(@"^([A-Z]+)([0-9+]):([A-Z]+)([0-9+])$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// A1 で表現されるA1形式の検証。
        /// </summary>
        private static readonly Regex SingleRangeRegex = new Regex(@"^([A-Z]+)([0-9+])$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// A で表現されるA1形式の検証。
        /// </summary>
        private static readonly Regex SingleColumnRangeRegex = new Regex(@"^([A-Z]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// A:B で表現されるA1形式の検証。
        /// </summary>
        private static readonly Regex ColumnRangeRegex = new Regex(@"^([A-Z]+):([A-Z]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// 1 で表現されるA1形式の検証。
        /// </summary>
        private static readonly Regex SingleRowRangeRegex = new Regex(@"^([0-9]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// 1:2 で表現されるA1形式の検証。
        /// </summary>
        private static readonly Regex RowRangeRegex = new Regex(@"^([0-9]+):([0-9]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// A1形式のセル範囲文字列からセル範囲のインデックス値に変換します。<br/>
        /// </summary>
        /// <param name="range">A1形式のセル範囲文字列。</param>
        /// <returns>セル範囲のインデックス値。行の指定、列の指定がそれぞれない場合、それぞれのフィールドは -1 を返却します。</returns>
        /// <exception cref="FormatException">range がA1形式ではありません。</exception>
        /// <remarks>
        /// ex.)
        /// <code>
        ///     var value = MetSpreadsheetConvert.ToCells("A1:A2"); // Row1=0, Column1=0, Row2=1, Column2=0
        ///     var value = MetSpreadsheetConvert.ToCells("A1");    // Row1=0, Column1=0, Row2=0, Column2=0
        ///     var value = MetSpreadsheetConvert.ToCells("A");     // Row1=-1, Column1=0, Row2=-1, Column2=0
        ///     var value = MetSpreadsheetConvert.ToCells("A:B");   // Row1=-1, Column1=0, Row2=-1, Column2=1
        ///     var value = MetSpreadsheetConvert.ToCells("1");     // Row1=0, Column1=-1, Row2=0, Column2=-1
        ///     var value = MetSpreadsheetConvert.ToCells("1:2");   // Row1=0, Column1=-1, Row2=1, Column2=-1
        /// </code>
        /// range は A1形式 でなければなりません。
        /// </remarks>
        public static (int Row1, int Column1, int Row2, int Column2) ToCells(string range)
        {
            Match match;

            // A1:A2 形式
            match = FullRangeRegex.Match(range);
            if (match.Success)
            {
                return (
                    int.Parse(match.Groups[2].Value) - 1,
                    ToColumnIndex(match.Groups[1].Value),
                    int.Parse(match.Groups[4].Value) - 1,
                    ToColumnIndex(match.Groups[3].Value)
                );
            }

            // A1 形式
            match = SingleRangeRegex.Match(range);
            if (match.Success)
            {
                return (
                    int.Parse(match.Groups[2].Value) - 1,
                    ToColumnIndex(match.Groups[1].Value),
                    int.Parse(match.Groups[2].Value) - 1,
                    ToColumnIndex(match.Groups[1].Value)
                );
            }

            // A 形式
            match = SingleColumnRangeRegex.Match(range);
            if (match.Success)
            {
                return (
                    -1,
                    ToColumnIndex(match.Groups[1].Value),
                    -1,
                    ToColumnIndex(match.Groups[1].Value)
                );
            }

            // A:B 形式
            match = ColumnRangeRegex.Match(range);
            if (match.Success)
            {
                return (
                    -1,
                    ToColumnIndex(match.Groups[1].Value),
                    -1,
                    ToColumnIndex(match.Groups[2].Value)
                );
            }

            // 1 形式
            match = SingleRowRangeRegex.Match(range);
            if (match.Success)
            {
                return (
                    int.Parse(match.Groups[1].Value) - 1,
                    -1,
                    int.Parse(match.Groups[1].Value) - 1,
                    -1
                );
            }

            // 1:2 形式
            match = RowRangeRegex.Match(range);
            if (match.Success)
            {
                return (
                    int.Parse(match.Groups[1].Value) - 1,
                    -1,
                    int.Parse(match.Groups[2].Value) - 1,
                    -1
                );
            }

            throw new FormatException("The cell range format is incorrect.");
        }

        ///// <summary>
        ///// シートに値の書き出しを行います。
        ///// 行、列のインデックスのいずれかが設定されていない項目は書き出しを行いません。
        ///// </summary>
        ///// <param name="value">書き出しを行うオブジェクト。</param>
        //public static void Write<T>(object value) where T : MetSpreadsheetOperatorBase
        //{
        //    Write<T>(value, null);
        //}

        ///// <summary>
        ///// シートに値の書き出しを行います。
        ///// 行、列のインデックスのいずれかが設定されていない項目は書き出しを行いません。
        ///// </summary>
        ///// <param name="value">書き出しを行うオブジェクト。</param>
        ///// <param name="param">実行パラメーター。</param>
        //public static void Write<T>(object value, object param) where T : MetSpreadsheetOperatorBase
        //{

        //}

        ///// <summary>
        ///// シートに値の書き出しを行います。
        ///// 行、列のインデックスのいずれかが設定されていない項目は、開始インデックスからマッピング方向へ書き出されます。
        ///// </summary>
        ///// <param name="value">書き出しを行うオブジェクト。</param>
        ///// <param name="mapIndex">マッピングする項目の書き出し開始インデックス。</param>
        ///// <param name="mapDirection">マッピング方向。</param>
        ///// <exception cref="ArgumentOutOfRangeException">mapIndex は 0 以上でなければなりません。</exception>
        ///// <remarks>
        ///// mapIndex は 0 以上でなければなりません。
        ///// </remarks>
        //public static void Write<T>(object value, int mapIndex, MapDirection mapDirection) where T : MetSpreadsheetOperatorBase
        //{
        //    Write<T>(value, mapIndex, mapDirection, 1, null);
        //}

        ///// <summary>
        ///// シートに値の書き出しを行います。
        ///// 行、列のインデックスのいずれかが設定されていない項目は、開始インデックスからマッピング方向へ書き出されます。
        ///// </summary>
        ///// <param name="value">書き出しを行うオブジェクト。</param>
        ///// <param name="mapIndex">マッピングする項目の書き出し開始インデックス。</param>
        ///// <param name="mapDirection">マッピング方向。</param>
        ///// <param name="param">実行パラメーター。</param>
        ///// <exception cref="ArgumentOutOfRangeException">mapIndex は 0 以上でなければなりません。</exception>
        ///// <remarks>
        ///// mapIndex は 0 以上でなければなりません。
        ///// </remarks>
        //public static void Write<T>(object value, int mapIndex, MapDirection mapDirection, object param) where T : MetSpreadsheetOperatorBase
        //{
        //    Write<T>(value, mapIndex, mapDirection, 1, param);
        //}

        ///// <summary>
        ///// シートに値の書き出しを行います。
        ///// 行、列のインデックスのいずれかが設定されていない項目は、開始インデックスからマッピング方向へ、スキップするセル数分をスキップしながら書き出されます。
        ///// </summary>
        ///// <param name="value">書き出しを行うオブジェクト。</param>
        ///// <param name="mapIndex">マッピングする項目の書き出し開始インデックス。</param>
        ///// <param name="mapDirection">マッピング方向。</param>
        ///// <param name="skip">マッピング方向へのスキップするセル数。</param>
        ///// <exception cref="ArgumentOutOfRangeException">mapIndex は 0 以上でなければなりません。</exception>
        ///// <remarks>
        ///// mapIndex は 0 以上でなければなりません。
        ///// </remarks>
        //public static void Write<T>(object value, int mapIndex, MapDirection mapDirection, int skip) where T : MetSpreadsheetOperatorBase
        //{
        //    Write<T>(value, mapIndex, mapDirection, skip, null);
        //}

        ///// <summary>
        ///// シートに値の書き出しを行います。
        ///// 行、列のインデックスのいずれかが設定されていない項目は、開始インデックスからマッピング方向へ、スキップするセル数分をスキップしながら書き出されます。
        ///// </summary>
        ///// <param name="value">書き出しを行うオブジェクト。</param>
        ///// <param name="mapIndex">0 から始まるマッピングする項目の書き出し開始インデックス。</param>
        ///// <param name="mapDirection">マッピング方向。</param>
        ///// <param name="skip">マッピング方向へのスキップするセル数。</param>
        ///// <param name="param">実行パラメーター。</param>
        ///// <exception cref="ArgumentOutOfRangeException">mapIndex は 0 以上でなければなりません。</exception>
        ///// <remarks>
        ///// mapIndex は 0 以上でなければなりません。
        ///// </remarks>
        //public static void Write<T>(object value, int mapIndex, MapDirection mapDirection, int skip, object param) where T : MetSpreadsheetOperatorBase
        //{
        //    if (mapIndex < 1)
        //    {
        //        throw new ArgumentOutOfRangeException(nameof(mapIndex));
        //    }

        //}

        ///// <summary>
        ///// シートから値の読み込みを行います。
        ///// </summary>
        ///// <typeparam name="T1"></typeparam>
        ///// <typeparam name="T2">読み込み結果を扱うクラス。</typeparam>
        ///// <returns>読み込みを行ったオブジェクト。</returns>
        //public static T2 Read<T1, T2>() where T1 : MetSpreadsheetOperatorBase where T2 : class, new()
        //{
        //    var result = new T2();

        //    return result;
        //}

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
