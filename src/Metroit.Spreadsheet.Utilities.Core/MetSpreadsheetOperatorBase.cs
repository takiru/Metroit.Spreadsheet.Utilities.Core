using Metroit.Spreadsheet.Utilities.Core.Mapping;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace Metroit.Spreadsheet.Utilities.Core
{
    /// <summary>
    /// シートへ情報を読み込み／書き出しを行うための基底操作を提供します。
    /// </summary>
    public abstract class MetSpreadsheetOperatorBase
    {
        /// <summary>
        /// 追加パラメーター情報を取得します。
        /// </summary>
        public object Param { get; private set; } = null;

        /// <summary>
        /// 読み込み前に発生します。
        /// </summary>
        protected event CancelEventHandler PreReading;

        /// <summary>
        /// 読み込み前に行います。
        /// </summary>
        /// <param name="e">キャンセルできるイベントのデータ。</param>
        protected virtual void OnPreReading(CancelEventArgs e)
        {
            PreReading?.Invoke(this, e);
        }

        protected abstract void OnRead<T>();

        /// <summary>
        /// 読み込み後に行います。
        /// </summary>
        protected virtual void OnRead() { }

        /// <summary>
        /// 書き出し前に発生します。
        /// </summary>
        protected event CancelEventHandler PreWriting;

        /// <summary>
        /// 書き出し前に行います。
        /// </summary>
        /// <param name="e">キャンセルできるイベントのデータ。</param>
        protected virtual void OnPreWriting(CancelEventArgs e)
        {
            PreWriting?.Invoke(this, e);
        }

        /// <summary>
        /// 書き出しを行います。
        /// </summary>
        protected abstract void OnWrite();

        /// <summary>
        /// 書き出し後に行います。
        /// </summary>
        protected virtual void OnWrote() { }

        /// <summary>
        /// 処理対象となったプロパティ情報。
        /// </summary>
        private TargetItem _parsedProperties;

        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        protected MetSpreadsheetOperatorBase()
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
        internal void Write(object value, object param)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            Param = param;
            var e = new CancelEventArgs(false);
            OnPreWriting(e);
            if (e.Cancel)
            {
                return;
            }

            _parsedProperties = new TargetItem(value);
            Parse(_parsedProperties, typeof(CellOutputMapAttribute));

            TestResult(_parsedProperties);

            OnWrote();
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
        /// value は null 以外でなければなりません。<br/>
        /// mapIndex は 0 以上でなければなりません。
        /// </remarks>
        internal void Write(object value, int mapIndex, MapDirection mapDirection, int skip, object param)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }
            if (mapIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(mapIndex));
            }

            Param = param;
            var e = new CancelEventArgs(false);
            OnPreWriting(e);
            if (e.Cancel)
            {
                return;
            }

            _parsedProperties = new TargetItem(value);
            Parse(_parsedProperties, typeof(CellOutputMapAttribute));

            TestResult(_parsedProperties);

            OnWrote();
        }

        /// <summary>
        /// オブジェクトを解析し、処理対象となるプロパティ要素を認識する。
        /// </summary>
        /// <param name="targetItem">対象オブジェクト。</param>
        /// <param name="attribute">処理対象属性。</param>
        /// <param name="innerType">再帰処理によって解析を行う配列要素やジェネリックの型。</param>
        private void Parse(TargetItem targetItem, Type attribute, Type innerType = null)
        {
            Type t = innerType == null ? targetItem.Value.GetType() : innerType;
            var pis = t.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);

            foreach (var pi in pis)
            {
                Type safeType;
                switch (GetPropertyTypeFactor(pi.PropertyType, out safeType))
                {
                    case PropertyTypeFactor.Primitive:
                    case PropertyTypeFactor.NearPrimitive:
                        AddPrimitiveProperty(pi, attribute, targetItem);
                        break;

                    case PropertyTypeFactor.NotEnumerable:
                        AddNotEnumerableProperty(pi, attribute, pi.GetValue(targetItem.Value), targetItem);
                        break;

                    case PropertyTypeFactor.Array:
                        AddArrayProperty(pi, safeType, attribute, pi.GetValue(targetItem.Value), targetItem);
                        break;

                    case PropertyTypeFactor.IList:
                        AddIListProperty(pi, safeType, attribute, pi.GetValue(targetItem.Value), targetItem);
                        break;

                    case PropertyTypeFactor.IDictionary:
                        AddIDictionaryProperty(pi, safeType, attribute, pi.GetValue(targetItem.Value), targetItem);
                        break;
                }
            }
        }

        /// <summary>
        /// 解析に必要な、型が有する要素を求める。
        /// 複数の要素を保有していても、下記の順で判断および取得を行う。
        ///   1. プリミティブ型かどうか。
        ///   2. プリミティブ型に近い型かどうか。
        ///   3. 反復処理を持たない参照型かどうか。
        ///   4. 配列かどうか。
        ///   5. IList を実装しているかどうか。
        ///   6. IDictionary を実装しているかどうか。
        /// </summary>
        /// <param name="type">検査する型。</param>
        /// <param name="safeType">
        /// Nullable を除いた安全な型。
        /// 配列の場合は要素の型、IListの場合はジェネリックの型、IDictionaryの場合は第二ジェネリックの型を返却する。
        /// </param>
        /// <returns>検査する型が有している要素。</returns>
        /// <remarks>
        /// プリミティブ型に近い型とは、string, decimal, DateTime, TimeSpan を示す。
        /// </remarks>
        private PropertyTypeFactor GetPropertyTypeFactor(Type type, out Type safeType)
        {
            var challengeType = Nullable.GetUnderlyingType(type) ?? type;

            // プリミティブ型
            if (challengeType.IsPrimitive)
            {
                safeType = challengeType;
                return PropertyTypeFactor.Primitive;
            }

            // プリミティブ型に近い型
            var nearPrimitiveTypes = new Type[] {
                typeof(string),
                typeof(decimal),
                typeof(DateTime),
                typeof(TimeSpan)
            };
            if (nearPrimitiveTypes.Contains(challengeType))
            {
                safeType = challengeType;
                return PropertyTypeFactor.NearPrimitive;
            }

            // 反復処理を持たない参照型
            if (!challengeType.GetInterfaces().Any(x => x == typeof(IEnumerable)))
            {
                safeType = challengeType;
                return PropertyTypeFactor.NotEnumerable;
            }

            // 配列
            if (challengeType.IsArray)
            {
                safeType = Nullable.GetUnderlyingType(challengeType.GetElementType()) ?? challengeType.GetElementType();
                return PropertyTypeFactor.Array;
            }

            // ILIst<>
            if (challengeType.GetInterfaces().Any(x => x.IsGenericType && x.GetGenericTypeDefinition() == typeof(IList<>)))
            {
                safeType = Nullable.GetUnderlyingType(challengeType.GenericTypeArguments[0]) ?? challengeType.GenericTypeArguments[0];
                return PropertyTypeFactor.IList;
            }

            // IDictionary<,>
            if (challengeType.GetInterfaces().Any(x => x.IsGenericType && x.GetGenericTypeDefinition() == typeof(IDictionary<,>)))
            {
                safeType = Nullable.GetUnderlyingType(challengeType.GenericTypeArguments[1]) ?? challengeType.GenericTypeArguments[1];
                return PropertyTypeFactor.IDictionary;
            }

            safeType = null;
            return PropertyTypeFactor.Unknown;
        }

        /// <summary>
        /// プロパティに処理対象とする属性が含まれるかどうかを取得する。
        /// </summary>
        /// <param name="pi">検証を行うプロパティ。</param>
        /// <param name="attribute">処理対象とする、プロパティに保有が必要な属性。</param>
        /// <returns>処理対象となるプロパティの場合は true, それ以外は false。</returns>
        private bool HasTargetAttribute(PropertyInfo pi, Type attribute)
        {
            if (Attribute.GetCustomAttribute(pi, attribute) == null)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// プリミティブ型のプロパティを対象プロパティとして追加する。
        /// </summary>
        /// <param name="pi">検証を行うプロパティ。</param>
        /// <param name="attribute">処理対象とする、プロパティに保有が必要な属性。</param>
        /// <param name="targetItem">対象プロパティとして追加する対象アイテム。</param>
        private void AddPrimitiveProperty(PropertyInfo pi, Type attribute, TargetItem targetItem)
        {
            if (!HasTargetAttribute(pi, attribute))
            {
                return;
            }

            targetItem.AddProperty(pi);
        }

        /// <summary>
        /// 反復処理を持たないプロパティを対象プロパティとして追加する。
        /// </summary>
        /// <param name="pi">検証を行うプロパティ。</param>
        /// <param name="attribute">処理対象とする、プロパティに保有が必要な属性。</param>
        /// <param name="value">追加を行うオブジェクト。</param>
        /// <param name="targetItem">対象プロパティとして追加する対象アイテム。</param>
        private void AddNotEnumerableProperty(PropertyInfo pi, Type attribute, object value, TargetItem targetItem)
        {
            if (value == null)
            {
                return;
            }

            // クラスを再帰的に解析する
            var item = new TargetItem(value);
            Parse(item, attribute);
            if (item.Properties.Count > 0 || item.Children.Count > 0)
            {
                targetItem.AddChild(item);
            }
        }

        /// <summary>
        /// 配列のプロパティを対象プロパティとして追加する。
        /// </summary>
        /// <param name="pi">検証を行うプロパティ。</param>
        /// <param name="elementSafeType">Nullable を除いた配列要素の安全な型。</param>
        /// <param name="attribute">処理対象とする、プロパティに保有が必要な属性。</param>
        /// <param name="value">追加を行うオブジェクト。</param>
        /// <param name="targetItem">対象プロパティとして追加する対象アイテム。</param>
        private void AddArrayProperty(PropertyInfo pi, Type elementSafeType, Type attribute, object value, TargetItem targetItem)
        {
            if (value == null)
            {
                return;
            }

            // 配列要素がプリミティブ型もしくは近い型は単純追加する
            switch (GetPropertyTypeFactor(elementSafeType, out _))
            {
                case PropertyTypeFactor.Primitive:
                case PropertyTypeFactor.NearPrimitive:
                    AddPrimitiveProperty(pi, attribute, targetItem);
                    return;
            }

            // 配列要素がクラスの場合は再帰的に解析する
            var item = new TargetItem(value);
            Parse(item, attribute, item.Value.GetType().GetElementType());
            if (item.Properties.Count > 0 || item.Children.Count > 0)
            {
                targetItem.AddChild(item);
            }
        }

        /// <summary>
        /// IList&lt;&gt; を持つプロパティを対象プロパティとして追加する。
        /// </summary>
        /// <param name="pi">検証を行うプロパティ。</param>
        /// <param name="genericSafeType">Nullable を除いたジェネリックの安全な型。</param>
        /// <param name="attribute">処理対象とする、プロパティに保有が必要な属性。</param>
        /// <param name="value">追加を行うオブジェクト。</param>
        /// <param name="targetItem">対象プロパティとして追加する対象アイテム。</param>
        private void AddIListProperty(PropertyInfo pi, Type genericSafeType, Type attribute, object value, TargetItem targetItem)
        {
            if (value == null)
            {
                return;
            }

            // List<string> などのジェネリックがプリミティブ型もしくは近い型は単純追加する
            switch (GetPropertyTypeFactor(genericSafeType, out _))
            {
                case PropertyTypeFactor.Primitive:
                case PropertyTypeFactor.NearPrimitive:
                    AddPrimitiveProperty(pi, attribute, targetItem);
                    return;
            }

            // ジェネリックがクラスの場合は再帰的に解析する
            var item = new TargetItem(value);
            Parse(item, attribute, item.Value.GetType().GenericTypeArguments[0]);
            if (item.Properties.Count > 0 || item.Children.Count > 0)
            {
                targetItem.AddChild(item);
            }
        }

        /// <summary>
        /// IDictionary&lt;,&gt; を持つプロパティを対象プロパティとして追加する。
        /// </summary>
        /// <param name="pi">検証を行うプロパティ。</param>
        /// <param name="genericSafeType">Nullable を除いたジェネリックの安全な型。</param>
        /// <param name="attribute">処理対象とする、プロパティに保有が必要な属性。</param>
        /// <param name="value">追加を行うオブジェクト。</param>
        /// <param name="targetItem">対象プロパティとして追加する対象アイテム。</param>
        private void AddIDictionaryProperty(PropertyInfo pi, Type genericSafeType, Type attribute, object value, TargetItem targetItem)
        {
            if (value == null)
            {
                return;
            }

            // Dictionary<int, string> などの第二ジェネリックがプリミティブ型もしくは近い型は単純追加する
            switch (GetPropertyTypeFactor(genericSafeType, out _))
            {
                case PropertyTypeFactor.Primitive:
                case PropertyTypeFactor.NearPrimitive:
                    AddPrimitiveProperty(pi, attribute, targetItem);
                    return;
            }

            // 第二ジェネリックがクラスの場合は再帰的に解析する
            var item = new TargetItem(value);
            Parse(item, attribute, item.Value.GetType().GenericTypeArguments[1]);
            if (item.Properties.Count > 0 || item.Children.Count > 0)
            {
                targetItem.AddChild(item);
            }
        }






        // TODO: 解析した結果を試しに出力してみる
        private void TestResult(TargetItem item)
        {
            foreach (var pi in item.Properties)
            {
                if (pi.PropertyType.GetInterfaces().Any(x => x.IsGenericType && x.GetGenericTypeDefinition() == typeof(IDictionary<,>)))
                {
                    var values = (IDictionary)pi.GetValue(item.Value);
                    foreach (var value in values.Values)
                    {
                        Console.WriteLine(value);
                    }
                    continue;
                }

                if (item.Value.GetType().GetInterfaces().Any(x => x.IsGenericType && x.GetGenericTypeDefinition() == typeof(IDictionary<,>)))
                {
                    var dictionary = (IDictionary)item.Value;
                    object[] listAry = new object[dictionary.Values.Count];
                    dictionary.Values.CopyTo(listAry, 0);
                    var list = listAry.ToList();
                    foreach (var listItem in list)
                    {
                        var value = pi.GetValue(listItem);
                        Console.WriteLine(value);
                    }
                }
                else
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
            }

            foreach (var c in item.Children)
            {
                TestResult(c);
            }
        }
    }
}
