using Metroit.Spreadsheet.Utilities.Core.MapItem;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// セルの入力を行う時の設定を提供します。
    /// </summary>
    public interface IInputConfiguration
    {
        /// <summary>
        /// 入力「を行うセル情報を意図に応じて設定します。
        /// </summary>
        /// <param name="map">スプレッドシートのセルから読み込みを行う情報。</param>
        /// <param name="param">実行パラメーター。</param>
        void ConfigureCell(CellInputMapItem map, object param);
    }
}
