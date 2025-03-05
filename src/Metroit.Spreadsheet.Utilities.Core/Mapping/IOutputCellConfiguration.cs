using Metroit.Spreadsheet.Utilities.Core.MapItem;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// セルの出力を行う時の出力セル設定を提供します。
    /// </summary>
    public interface IOutputCellConfiguration
    {
        /// <summary>
        /// 出力を行うセル情報を意図に応じて設定します。
        /// </summary>
        /// <param name="map">スプレッドシートのセルに書き出しを行う情報。</param>
        /// <param name="param">実行パラメーター。</param>
        void ConfigureCell(CellOutputMapItem map, object param);
    }
}
