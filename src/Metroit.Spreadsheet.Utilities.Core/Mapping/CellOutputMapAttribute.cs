using System;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// スプレッドシートのセルの出力セルインデックス情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CellOutputMapAttribute : Attribute
    {
        private int _row = -1;

        /// <summary>
        /// 0 から始まる行インデックス、または未指定を表す - 1 を取得または設定します。既定は -1 です。
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">行は -1 以上である必要があります。</exception>
        public int Row
        {
            get => _row;
            set
            {
                if (value < -1)
                {
                    throw new ArgumentOutOfRangeException(nameof(Row), value, "Row must be greater than or equal to -1.");
                }
                _row = value;
            }
        }

        private int _column = -1;

        /// <summary>
        /// 0 から始まる列インデックス、または未指定を表す - 1 を取得または設定します。既定は -1 です。
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">列は -1 以上である必要があります。</exception>
        public int Column
        {
            get => _column;
            set
            {
                if (value < -1)
                {
                    throw new ArgumentOutOfRangeException(nameof(Column), value, "Column must be greater than or equal to -1.");
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
        public CellOutputMapAttribute() { }
    }
}
