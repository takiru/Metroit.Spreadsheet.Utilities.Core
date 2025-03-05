namespace Metroit.Spreadsheet.Utilities.Core
{
    /// <summary>
    /// シートからデータを読み込む操作を提供します。
    /// </summary>
    /// <typeparam name="T">実際のスプレッドシートを操作するクラス。</typeparam>
    public class MetSpreadsheetDataReader<T> where T : MetSpreadsheetOperatorBase, new()
    {
        /// <summary>
        /// 実際のスプレッドシートを操作するオブジェクト。
        /// </summary>
        private readonly MetSpreadsheetOperatorBase Operator = null;

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        public MetSpreadsheetDataReader()
        {
            Operator = new T();
        }

        /// <summary>
        /// シートからセルの読み込みを行います。
        /// </summary>
        /// <typeparam name="T1">読み込み結果を扱うクラス。</typeparam>
        /// <returns>読み込みを行ったオブジェクト。</returns>
        public T1 Read<T1>() where T1 : class, new()
        {
            var result = new T1();

            return result;
        }
    }
}
