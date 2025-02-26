using Metroit.Spreadsheet.Utilities.Core.Mapping;

namespace Metroit.Spreadsheet.Utilities.Core.MapItem
{
    /// <summary>
    /// Excelセルの文字列の修飾情報を提供します。
    /// </summary>
    public class CellCharacterDecorationItem
    {
        /// <summary>
        /// 下線 を取得します。既定は null です。
        /// </summary>
        public UnderlineStyle? UnderlineStyle { get; set; }

        /// <summary>
        /// 文字付き位置 を取得します。既定は null　です。
        /// </summary>
        public CharacterPosition? CharacterPosition { get; set; }

        /// <summary>
        /// 新しい CellCharacterDecorationItem インスタンスを生成します。
        /// </summary>
        /// <param name="underlineStyle">下線。</param>
        /// <param name="characterPosition">文字付き位置。</param>
        public CellCharacterDecorationItem(UnderlineStyle? underlineStyle, CharacterPosition? characterPosition)
        {
            UnderlineStyle = underlineStyle;
            CharacterPosition = characterPosition;
        }
    }
}
