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
        /// 下線スタイルを取得します。既定は <see cref="UnderlineStyle.None"/> です。
        /// <see cref="UnderlineStyle.None"/> 以外が指定された時、<see cref="CellFontAttribute.FontStyle"/> で指定した下線情報は失われます。
        /// </summary>
        public UnderlineStyle UnderlineStyle { get; } = UnderlineStyle.None;

        /// <summary>
        /// 文字付き位置を取得します。既定は <see cref="CharacterPosition.Normal"/> です。
        /// </summary>
        public CharacterPosition CharacterPosition { get; } = CharacterPosition.Normal;

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="underlineStyle">下線。</param>
        /// <param name="characterPosition">文字付き位置。</param>
        public CellCharacterDecorationAttribute(UnderlineStyle underlineStyle, CharacterPosition characterPosition)
        {
            UnderlineStyle = underlineStyle;
            CharacterPosition = characterPosition;
        }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="underlineStyle">下線。</param>
        public CellCharacterDecorationAttribute(UnderlineStyle underlineStyle) : this(underlineStyle, CharacterPosition.Normal) { }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="characterPosition">文字付き位置。</param>
        public CellCharacterDecorationAttribute(CharacterPosition characterPosition) : this(UnderlineStyle.None, characterPosition) { }
    }
}
