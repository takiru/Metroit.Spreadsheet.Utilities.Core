using Metroit.Spreadsheet.Utilities.Core.Mapping;
using System;

namespace Metroit.Spreadsheet.Utilities.Core
{
    /// <summary>
    /// シートからデータを書き込む操作を提供します。
    /// </summary>
    /// <typeparam name="T">実際のスプレッドシートを操作するクラス。</typeparam>
    public class MetSpreadsheetDataWriter<T> where T : MetSpreadsheetOperatorBase, new()
    {
        /// <summary>
        /// 実際のスプレッドシートを操作するオブジェクト。
        /// </summary>
        private readonly MetSpreadsheetOperatorBase Operator = null;

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        public MetSpreadsheetDataWriter()
        {
            Operator = new T();
        }

        /// <summary>
        /// シートへセルの書き出しを行います。
        /// </summary>
        /// <param name="value">書き出しを行うオブジェクト。</param>
        /// <exception cref="ArgumentNullException">value が null です。</exception>
        /// <remarks>
        /// value は null 以外でなければなりません。
        /// </remarks>
        public void Write(object value)
        {
            Operator.Write(value, null);
        }

        /// <summary>
        /// シートへセルの書き出しを行います。
        /// </summary>
        /// <param name="value">書き出しを行うオブジェクト。</param>
        /// <param name="param">実行パラメーター。</param>
        /// <exception cref="ArgumentNullException">value が null です。</exception>
        /// <remarks>
        /// value は null 以外でなければなりません。
        /// </remarks>
        public void Write(object value, object param)
        {
            Operator.Write(value, param);
        }

        /// <summary>
        /// シートへセルの書き出しを行います。
        /// 行、列のインデックスのいずれかが設定されていない項目は、開始インデックスからマッピング方向へ書き出されます。
        /// </summary>
        /// <param name="value">書き出しを行うオブジェクト。</param>
        /// <param name="mapIndex">マッピングする項目の書き出し開始インデックス。</param>
        /// <param name="mapDirection">マッピング方向。</param>
        /// <exception cref="ArgumentNullException">value が null です。</exception>
        /// <exception cref="ArgumentOutOfRangeException">mapIndex は 0 以上でなければなりません。</exception>
        /// <remarks>
        /// value は null 以外でなければなりません。
        /// mapIndex は 0 以上でなければなりません。
        /// </remarks>
        public void Write(object value, int mapIndex, MapDirection mapDirection)
        {
            Operator.Write(value, mapIndex, mapDirection, 1, null);
        }

        /// <summary>
        /// シートへセルの書き出しを行います。
        /// 行、列のインデックスのいずれかが設定されていない項目は、開始インデックスからマッピング方向へ書き出されます。
        /// </summary>
        /// <param name="value">書き出しを行うオブジェクト。</param>
        /// <param name="mapIndex">マッピングする項目の書き出し開始インデックス。</param>
        /// <param name="mapDirection">マッピング方向。</param>
        /// <param name="param">実行パラメーター。</param>
        /// <exception cref="ArgumentNullException">value が null です。</exception>
        /// <exception cref="ArgumentOutOfRangeException">mapIndex は 0 以上でなければなりません。</exception>
        /// <remarks>
        /// value は null 以外でなければなりません。
        /// mapIndex は 0 以上でなければなりません。
        /// </remarks>
        public void Write(object value, int mapIndex, MapDirection mapDirection, object param)
        {
            Operator.Write(value, mapIndex, mapDirection, 1, param);
        }

        /// <summary>
        /// シートへセルの書き出しを行います。
        /// 行、列のインデックスのいずれかが設定されていない項目は、開始インデックスからマッピング方向へ、スキップするセル数分をスキップしながら書き出されます。
        /// </summary>
        /// <param name="value">書き出しを行うオブジェクト。</param>
        /// <param name="mapIndex">マッピングする項目の書き出し開始インデックス。</param>
        /// <param name="mapDirection">マッピング方向。</param>
        /// <param name="skip">マッピング方向へのスキップするセル数。</param>
        /// <exception cref="ArgumentNullException">value が null です。</exception>
        /// <exception cref="ArgumentOutOfRangeException">mapIndex は 0 以上でなければなりません。</exception>
        /// <remarks>
        /// value は null 以外でなければなりません。
        /// mapIndex は 0 以上でなければなりません。
        /// </remarks>
        public void Write(object value, int mapIndex, MapDirection mapDirection, int skip)
        {
            Operator.Write(value, mapIndex, mapDirection, skip, null);
        }

        /// <summary>
        /// シートへセルの書き出しを行います。
        /// 行、列のインデックスのいずれかが設定されていない項目は、開始インデックスからマッピング方向へ、スキップするセル数分をスキップしながら書き出されます。
        /// </summary>
        /// <param name="value">書き出しを行うオブジェクト。</param>
        /// <param name="mapIndex">マッピングする項目の書き出し開始インデックス。</param>
        /// <param name="mapDirection">マッピング方向。</param>
        /// <param name="skip">マッピング方向へのスキップするセル数。</param>
        /// <param name="param">実行パラメーター。</param>
        /// <exception cref="ArgumentNullException">value が null です。</exception>
        /// <exception cref="ArgumentOutOfRangeException">mapIndex は 0 以上でなければなりません。</exception>
        /// <remarks>
        /// value は null 以外でなければなりません。
        /// mapIndex は 0 以上でなければなりません。
        /// </remarks>
        public void Write(object value, int mapIndex, MapDirection mapDirection, int skip, object param)
        {
            Operator.Write(value, mapIndex, mapDirection, skip, param);
        }
    }
}
