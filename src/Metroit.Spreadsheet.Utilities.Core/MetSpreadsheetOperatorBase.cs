using Metroit.Spreadsheet.Utilities.Core.Mapping;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace Metroit.Spreadsheet.Utilities.Core
{
    /// <summary>
    /// シートへ情報を読み込み／書き出しを行うための基底操作を提供します。
    /// </summary>
    public abstract class MetSpreadsheetOperatorBase
    {
        //public object Param { get; } = null;



        //protected internal MetSpreadsheetOperatorBase(object param)
        //{
        //    Param = param;
        //}

        protected MetSpreadsheetOperatorBase()
        {

        }

        /// <summary>
        /// 書き出し前の制御。
        /// </summary>
        protected virtual void OnPreWriting()
        {

        }

        /// <summary>
        /// 書き出し前の制御。
        /// </summary>
        protected virtual void OnPreReading()
        {

        }

        /// <summary>
        /// シートへセルの書き出しを行います。
        /// </summary>
        /// <param name="value">書き出しを行うオブジェクト。</param>
        /// <param name="param">実行パラメーター。</param>
        /// <exception cref="ArgumentNullException">value が null です。</exception>
        /// <remarks>
        /// value は null 以外でなければなりません。
        /// </remarks>
        public void Write(object value, object param)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            OnPreWriting();

            //_parsedProperties = new List<PropertyInfo>();
            //Parse(value, typeof(CellOutputMapAttribute));
            _parsedProperties = new TargetItem();
            Parse(value, typeof(CellOutputMapAttribute), _parsedProperties);

            TestResult(_parsedProperties);
            Console.WriteLine("Finish");
        }

        // TODO: 解析した結果を試しに出力してみる
        private void TestResult(TargetItem item)
        {
            foreach (var pi in item.TargetProperties)
            {
                // 対象の値自体が IEnumerable なら、要素が複数ある(List<クラス>とか)
                if (item.Value is IEnumerable enumItems)
                {
                    foreach (var enumItem in enumItems)
                    {
                        var a = pi.GetValue(enumItem);
                        Console.WriteLine(a);
                    }
                }
                else
                {
                    var a = pi.GetValue(item.Value);

                    // プロパティの値が配列なら
                    if (a is Array arys)
                    {
                        foreach (var ary in arys)
                        {
                            Console.WriteLine(ary);
                        }
                    }
                    else
                    {
                        if (a is string || a.GetType().IsPrimitive)
                        {
                            Console.WriteLine(a);
                        }
                        else
                        {
                            // プロパティの値が IEnumerable なら
                            if (a is IEnumerable enumItems2)
                            {
                                foreach (var enumItem in enumItems2)
                                {
                                    Console.WriteLine(enumItem);
                                }
                            }
                        }
                    }
                }
            }

            foreach (var c in item.Child)
            {
                TestResult(c);
            }
        }

        /// <summary>
        /// シートへセルの書き出しを行います。
        /// 行、列のインデックスのいずれかが設定されていない項目は、開始インデックスからマッピング方向へ、スキップするセル数分をスキップしながら書き出されます。
        /// </summary>
        /// <param name="value">書き出しを行うオブジェクト。</param>
        /// <param name="mapIndex">マッピングする項目の書き出し開始インデックス。</param>
        /// <param name="mapDirection">マッピング方向。</param>
        /// <param name="skip">マッピング方向へのスキップするセル数。</param>
        /// <param name="param">実行パラメーター。</param>
        /// <exception cref="ArgumentNullException">value が null です。</exception>
        /// <exception cref="ArgumentOutOfRangeException">mapIndex は 0 以上でなければなりません。</exception>
        /// <remarks>
        /// value は null 以外でなければなりません。
        /// mapIndex は 0 以上でなければなりません。
        /// </remarks>
        protected internal void Write(object value, int mapIndex, MapDirection mapDirection, int skip, object param)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }
            if (mapIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(mapIndex));
            }

            OnPreWriting();

            //_parsedProperties = new List<PropertyInfo>();
            _parsedProperties = new TargetItem();
            Parse(value, typeof(CellOutputMapAttribute), _parsedProperties);
        }


        private TargetItem _parsedProperties = new TargetItem();

        private bool IsPrimitiveOrNearPrimitive(Type safeType)
        {
            // プリミティブ型もしくは特定の型の場合、判定する
            if (safeType.IsPrimitive)
            {
                return true;
            }
            if (safeType == typeof(string) || safeType == typeof(DateTime) || safeType == typeof(decimal) || safeType == typeof(TimeSpan))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// オブジェクトを解析し、処理対象となるプロパティ要素を認識する。
        /// </summary>
        /// <param name="value">対象オブジェクト。</param>
        /// <param name="attribute">処理対象属性。</param>
        private void Parse(object value, Type attribute, TargetItem targetItem, object childValue = null)
        {
            Type t;
            if (value is Type)
            {
                t = (Type)value;
            }
            else
            {
                t = value.GetType();
            }

            var pis = t.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);
            if (childValue == null)
            {
                targetItem.Value = value;
            }

            foreach (var pi in pis)
            {
                // プリミティブ型もしくは特定の型の場合、判定する
                var safeType = Nullable.GetUnderlyingType(pi.PropertyType) ?? pi.PropertyType;

                if (IsPrimitiveOrNearPrimitive(safeType))
                {
                    // 処理対象となるプロパティでなければ認識対象としない
                    var mapAttr2 = Attribute.GetCustomAttribute(pi, attribute);
                    if (mapAttr2 == null)
                    {
                        continue;
                    }

                    targetItem.TargetProperties.Add(pi);
                    continue;
                }

                // 反復処理のある型の場合は再帰的に解析する
                if (safeType.GetInterfaces().Where(x => x == typeof(IEnumerable)).FirstOrDefault() != null)
                {
                    // IENumerable を実装したプロパティで、プリミティブ型だった時、処理対象となるプロパティであれば認識対象にする
                    // プリミティブ型もしくは特定の型の場合、判定する
                    // 配列なら受け入れる
                    if (safeType.IsArray)
                    {
                        var safeType2 = Nullable.GetUnderlyingType(safeType.GetElementType()) ?? safeType.GetElementType();
                        if (IsPrimitiveOrNearPrimitive(safeType2))
                        {
                            // 配列でも、処理対象となるプロパティでなければ認識対象としない
                            var mapAttr2 = Attribute.GetCustomAttribute(pi, attribute);
                            if (mapAttr2 == null)
                            {
                                continue;
                            }
                            targetItem.TargetProperties.Add(pi);
                            continue;
                        }


                        var item = new TargetItem();
                        if (childValue == null)
                        {
                            item.Value = pi.GetValue(value);
                        }
                        else
                        {
                            item.Value = pi.GetValue(childValue);
                        }

                        if (item.Value != null)
                        {
                            Parse(item.Value.GetType().GetElementType(), attribute, item, item.Value);
                            if (item.TargetProperties.Count > 0 || item.Child.Count > 0)
                            {
                                targetItem.Child.Add(item);
                            }
                        }
                        continue;
                    }

                    // IListなら受け入れる
                    if (safeType.GetInterfaces().Where(x => x == typeof(IList)).FirstOrDefault() != null)
                    {
                        var safeType2 = Nullable.GetUnderlyingType(safeType.GenericTypeArguments[0]) ?? safeType.GenericTypeArguments[0];
                        if (IsPrimitiveOrNearPrimitive(safeType2))
                        {
                            targetItem.TargetProperties.Add(pi);
                            continue;
                        }

                        //// 反復処理のあるジェネリックの場合は再帰的に解析する
                        var item = new TargetItem();
                        if (childValue == null)
                        {
                            item.Value = pi.GetValue(value);
                        }
                        else
                        {
                            item.Value = pi.GetValue(childValue);
                        }

                        if (item.Value != null)
                        {
                            Parse(item.Value.GetType().GenericTypeArguments[0], attribute, item, item.Value);
                            if (item.TargetProperties.Count > 0 || item.Child.Count > 0)
                            {
                                targetItem.Child.Add(item);
                            }
                        }
                        continue;
                    }

                    // 他の型は受け入れない
                    continue;
                }


                // 自前クラスの場合
                var item2 = new TargetItem();
                if (value is Type)
                {
                    item2.Value = pi.GetValue(childValue);
                }
                else
                {
                    item2.Value = pi.GetValue(value);
                }

                if (item2.Value != null)
                {
                    Parse(item2.Value.GetType(), attribute, item2, item2.Value);
                    if (item2.TargetProperties.Count > 0 || item2.Child.Count > 0)
                    {
                        targetItem.Child.Add(item2);
                    }
                }
            }
        }
    }


    class TargetItem
    {
        public object Value { get; set; }

        public List<PropertyInfo> TargetProperties { get; set; } = new List<PropertyInfo>();

        public List<TargetItem> Child { get; set; } = new List<TargetItem>();
    }
}
