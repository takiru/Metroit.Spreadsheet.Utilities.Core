using Metroit.Spreadsheet.Utilities.Core.MapItem;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// セルの出力を行う時の無視設定を提供します。
    /// </summary>
    public interface IOutputIgnoreConfiguration
    {
        /// <summary>
        /// 出力を行うセルを無視するかどうかを取得します。
        /// </summary>
        /// <param name="map">スプレッドシートのセルに書き出しを行う情報。
        /// <see cref="IOutputCellConfiguration"/>によってセル情報が変更された場合、変更のセル情報が得られます。
        /// </param>
        /// <param name="param">実行パラメーター。</param>
        /// <returns>出力しない場合は true, それ以外は false。</returns>
        bool IgnoreOutput(IReadOnlyCellMapItem map, object param);
    }
}
