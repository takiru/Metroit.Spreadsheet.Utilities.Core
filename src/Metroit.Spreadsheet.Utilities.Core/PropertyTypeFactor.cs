namespace Metroit.Spreadsheet.Utilities.Core
{
    /// <summary>
    /// プロパティのタイプ要素を定義します。
    /// </summary>
    enum PropertyTypeFactor
    {
        /// <summary>
        /// プリミティブ型を示します。
        /// </summary>
        Primitive,

        /// <summary>
        /// プリミティブ型に近い型を示します。
        /// </summary>
        NearPrimitive,

        /// <summary>
        /// 反復処理を持たないことを示します。
        /// </summary>
        NotEnumerable,

        /// <summary>
        /// 配列を示す。
        /// </summary>
        Array,

        /// <summary>
        /// ILIst&lt;&gt;を示します。
        /// </summary>
        IList,

        /// <summary>
        /// IDictionary&lt;,&gt;を示します。
        /// </summary>
        IDictionary,

        Unknown
    }
}
