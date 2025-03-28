using System;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// スプレッドシートのセルの出力セルインデックス情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CellMapAttribute : Attribute
    {
        /// <summary>
        /// 未指定の場合のインデックス値を示します。
        /// </summary>
        internal static readonly int UnspecifiedIndex = -1;

        private int _row = UnspecifiedIndex;

        /// <summary>
        /// 0 から始まる行インデックスを取得または設定します。
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">行は -1 以上である必要があります。</exception>
        public int Row
        {
            get => _row;
            set
            {
                if (value < UnspecifiedIndex)
                {
                    throw new ArgumentOutOfRangeException(nameof(Row), value, $"Row must be greater than or equal to {UnspecifiedIndex}.");
                }
                _row = value;
            }
        }

        private int _column = UnspecifiedIndex;

        /// <summary>
        /// 0 から始まる列インデックスを取得または設定します。
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">列は -1 以上である必要があります。</exception>
        public int Column
        {
            get => _column;
            set
            {
                if (value < UnspecifiedIndex)
                {
                    throw new ArgumentOutOfRangeException(nameof(Column), value, $"Column must be greater than or equal to {UnspecifiedIndex}.");
                }
                _column = value;
            }
        }

        /// <summary>
        /// 数式のセルかどうかを取得または設定します。既定は false です。
        /// </summary>
        public bool Formula { get; set; } = false;

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        public CellMapAttribute() { }
    }
}
