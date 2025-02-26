using System;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// Excelセルの文字列の修飾情報を提供します。
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CellCharacterDecorationAttribute : Attribute
    {
        /// <summary>
        /// 下線 を取得します。既定は null です。
        /// </summary>
        public UnderlineStyle? UnderlineStyle { get; } = null;

        /// <summary>
        /// 文字付き位置 を取得します。既定は null　です。
        /// </summary>
        public CharacterPosition? CharacterPosition { get; } = null;

        /// <summary>
        /// 新しい CellCharacterDecorationAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="underlineStyle">下線。</param>
        public CellCharacterDecorationAttribute(UnderlineStyle underlineStyle)
        {
            UnderlineStyle = underlineStyle;
        }

        /// <summary>
        /// 新しい CellCharacterDecorationAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="characterPosition">文字付き位置。</param>
        public CellCharacterDecorationAttribute(CharacterPosition characterPosition)
        {
            CharacterPosition = characterPosition;
        }

        /// <summary>
        /// 新しい CellCharacterDecorationAttribute インスタンスを生成します。
        /// </summary>
        /// <param name="underlineStyle">下線。</param>
        /// <param name="characterPosition">文字付き位置。</param>
        public CellCharacterDecorationAttribute(UnderlineStyle underlineStyle, CharacterPosition characterPosition) : this(underlineStyle)
        {
            CharacterPosition = characterPosition;
        }
    }
}
