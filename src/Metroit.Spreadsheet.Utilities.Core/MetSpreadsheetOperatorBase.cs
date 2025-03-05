﻿using Metroit.Spreadsheet.Utilities.Core.Mapping;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Security.AccessControl;

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

            _parsedProperties = new TargetItem(value);
            Parse(_parsedProperties, typeof(CellOutputMapAttribute));

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
            //_parsedProperties = new TargetItem();
            //Parse(value, typeof(CellOutputMapAttribute), _parsedProperties);
        }


        private TargetItem _parsedProperties;

        /// <summary>
        /// 型がプリミティブ型もしくはプリミティブ型に近い型かどうかを取得する。
        /// プリミティブ型に近い型は string, DateTime, decimal, Timespan 。
        /// </summary>
        /// <param name="safeType">Nullable を除いた安全な型。</param>
        /// <returns>プリミティブ型もしくはプリミティブ型に近い型の場合は true, それ以外は false。</returns>
        private bool IsPrimitiveOrNearPrimitive(Type safeType)
        {
            // プリミティブ型もしくは特定の型の場合、判定する
            if (safeType.IsPrimitive)
            {
                return true;
            }
            if (safeType == typeof(string) || safeType == typeof(DateTime) ||
                safeType == typeof(decimal) || safeType == typeof(TimeSpan))
            {
                return true;
            }

            return false;
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
        /// プリミティブ型のプロパティを対象プロパティとして追加を検証する。
        /// </summary>
        /// <param name="pi">検証を行うプロパティ。</param>
        /// <param name="safeType">Nullable を除いた安全な型。</param>
        /// <param name="attribute">処理対象とする、プロパティに保有が必要な属性。</param>
        /// <param name="targetItem">対象プロパティとして追加する対象アイテム。</param>
        /// <returns>対象プロパティとして追加された場合は true, それ以外は false。</returns>
        private bool TryAddPrimitiveProperty(PropertyInfo pi, Type safeType, Type attribute, TargetItem targetItem)
        {
            if (!IsPrimitiveOrNearPrimitive(safeType))
            {
                return false;
            }

            if (!HasTargetAttribute(pi, attribute))
            {
                return false;
            }

            targetItem.TargetProperties.Add(pi);
            return true;
        }

        /// <summary>
        /// 配列のプロパティを対象プロパティとして追加する。
        /// </summary>
        /// <param name="pi">検証を行うプロパティ。</param>
        /// <param name="elementSafeType">Nullable を除いた配列要素の安全な型。</param>
        /// <param name="attribute">処理対象とする、プロパティに保有が必要な属性。</param>
        /// <param name="value"></param>
        /// <param name="targetItem">対象プロパティとして追加する対象アイテム。</param>
        private void AddArrayProperty(PropertyInfo pi, Type elementSafeType, Type attribute, object value, TargetItem targetItem)
        {
            if (value == null)
            {
                return;
            }

            // 配列要素がプリミティブ型
            if (IsPrimitiveOrNearPrimitive(elementSafeType))
            {
                if (!HasTargetAttribute(pi, attribute))
                {
                    return;
                }

                targetItem.TargetProperties.Add(pi);
                return;
            }

            // 配列要素がクラス
            var item = new TargetItem(value);
            Parse(item, attribute, item.Value.GetType().GetElementType());
            if (item.TargetProperties.Count > 0 || item.Child.Count > 0)
            {
                targetItem.Child.Add(item);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pi"></param>
        /// <param name="genericSafeType"></param>
        /// <param name="attribute"></param>
        /// <param name="value"></param>
        /// <param name="targetItem"></param>
        private void AddIListProperty(PropertyInfo pi, Type genericSafeType, Type attribute, object value, TargetItem targetItem)
        {
            // List<string> などのプリミティブの場合
            if (IsPrimitiveOrNearPrimitive(genericSafeType))
            {
                if (!HasTargetAttribute(pi, attribute))
                {
                    return;
                }
                targetItem.TargetProperties.Add(pi);
                return;
            }

            //// 反復処理のあるジェネリックの場合は再帰的に解析する
            var item = new TargetItem(value);

            if (item.Value != null)
            {
                Parse(item, attribute, item.Value.GetType().GenericTypeArguments[0]);
                if (item.TargetProperties.Count > 0 || item.Child.Count > 0)
                {
                    targetItem.Child.Add(item);
                }
            }
        }

        /// <summary>
        /// オブジェクトを解析し、処理対象となるプロパティ要素を認識する。
        /// </summary>
        /// <param name="value">対象オブジェクト。</param>
        /// <param name="attribute">処理対象属性。</param>
        private void Parse(TargetItem targetItem, Type attribute, Type childType = null)
        {
            Type t = childType == null ? targetItem.Value.GetType() : childType;
            var pis = t.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);

            foreach (var pi in pis)
            {
                var safeType = Nullable.GetUnderlyingType(pi.PropertyType) ?? pi.PropertyType;

                // プリミティブ型を想定して追加する
                if (TryAddPrimitiveProperty(pi, safeType, attribute, targetItem))
                {
                    continue;
                }

                // 反復処理のある型
                if (safeType.GetInterfaces().Where(x => x == typeof(IEnumerable)).FirstOrDefault() != null)
                {
                    // 配列なら受け入れる
                    if (safeType.IsArray)
                    {
                        var elementSafeType = Nullable.GetUnderlyingType(safeType.GetElementType()) ?? safeType.GetElementType();
                        AddArrayProperty(pi, elementSafeType, attribute, pi.GetValue(targetItem.Value), targetItem);
                        continue;
                    }

                    // IList 以外の IEnumerable は受け付けない
                    if (!safeType.GetInterfaces().Where(x => x == typeof(IList)).Any())
                    {
                        continue;
                    }

                    //// IListなら受け入れる
                    var genericSafeType = Nullable.GetUnderlyingType(safeType.GenericTypeArguments[0]) ?? safeType.GenericTypeArguments[0];
                    AddIListProperty(pi, genericSafeType, attribute, pi.GetValue(targetItem.Value), targetItem);
                    continue;
                }

                // 自前クラスの場合
                var item2 = new TargetItem(pi.GetValue(targetItem.Value));
                if (item2.Value != null)
                {
                    Parse(item2, attribute);
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
        public TargetItem(object value)
        {
            Value = value;
        }

        public object Value { get; }

        public List<PropertyInfo> TargetProperties { get; set; } = new List<PropertyInfo>();

        public List<TargetItem> Child { get; set; } = new List<TargetItem>();
    }
}
