using System.Collections.Generic;
using System.Reflection;

namespace Metroit.Spreadsheet.Utilities.Core
{
    /// <summary>
    /// 書き出しを行うオブジェクトを解析した結果、書き出し対象となるプロパティ情報を提供します。
    /// </summary>
    class TargetItem
    {
        /// <summary>
        /// プロパティを保有しているオブジェクトを取得します。
        /// </summary>
        public object Value { get; }

        private List<PropertyInfo> _properties = new List<PropertyInfo>();

        /// <summary>
        /// 書き出し対象となるプロパティを取得します。
        /// </summary>
        public IReadOnlyList<PropertyInfo> Properties => _properties;

        private List<TargetItem> _children = new List<TargetItem>();

        /// <summary>
        /// 書き出し対象となる <see cref="Value"/> 配下に存在するプロパティ情報を取得します。
        /// </summary>
        public IReadOnlyList<TargetItem> Children => _children;

        /// <summary>
        /// 新しいんスタンスを生成します。
        /// </summary>
        /// <param name="value">プロパティを保有しているオブジェクト。</param>
        public TargetItem(object value)
        {
            Value = value;
        }

        /// <summary>
        /// 書き出し対象となるプロパティを追加します。
        /// </summary>
        /// <param name="pi">プロパティ。</param>
        public void AddProperty(PropertyInfo pi)
        {
            _properties.Add(pi);
        }

        /// <summary>
        /// 書き出し対象となる <see cref="Value"/> 配下に存在するプロパティ情報を追加します。
        /// </summary>
        /// <param name="item"></param>
        public void AddChild(TargetItem item)
        {
            _children.Add(item);
        }
    }
}
