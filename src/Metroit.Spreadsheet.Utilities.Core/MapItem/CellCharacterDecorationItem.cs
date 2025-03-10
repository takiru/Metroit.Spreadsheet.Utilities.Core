using Metroit.Spreadsheet.Utilities.Core.Mapping;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルの文字列の修飾情報を提供します。
    /// </summary>
    public class CellCharacterDecorationItem
    {
        /// <summary>
        /// 下線を取得します。
        /// </summary>
        public UnderlineStyle UnderlineStyle { get; }

        /// <summary>
        /// 文字付き位置を取得します。
        /// </summary>
        public CharacterPosition CharacterPosition { get; }

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="underlineStyle">下線。</param>
        /// <param name="characterPosition">文字付き位置。</param>
        public CellCharacterDecorationItem(UnderlineStyle underlineStyle, CharacterPosition characterPosition)
        {
            UnderlineStyle = underlineStyle;
            CharacterPosition = characterPosition;
        }
    }
}
