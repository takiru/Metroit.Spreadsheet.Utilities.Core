using Metroit.Spreadsheet.Utilities.Core.MapItem;
using Metroit.Spreadsheet.Utilities.Core.Mapping;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Xml.Schema;

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
            Parse(_parsedProperties, typeof(CellMapAttribute));
            Write(_parsedProperties, 0, MapDirection.None, 0, -1);

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
            Parse(_parsedProperties, typeof(CellMapAttribute));

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

        /// <summary>
        /// 書き出しを行う。
        /// </summary>
        /// <param name="item">書き出しを行うプロパティ情報。</param>
        /// <param name="mapIndex">書き出しインデックス。</param>
        /// <param name="mapDirection">マッピング方向。</param>
        /// <param name="skip">マッピング方向へのスキップするセル数。</param>
        /// <param name="currentMapIndex">現在のマッピングする項目の書き出し開始インデックス。</param>
        private int Write(TargetItem item, int mapIndex, MapDirection mapDirection, int skip, int currentMapIndex)
        {
            var internalMapIndex = mapIndex;
            foreach (var pi in item.Properties)
            {
                // TODO: piが配列やIList, IDirectoryの時、MapShiftAttributeを取得して、その値分、mapIndexへ追加する。
                //       それがないなら、mapIndexのところに出力するって。
                //       skipなんて概念は、いらない。どこから書き出したいのか、だけでいい。
                //       んで、1クラスすべての組み合わせの出力が完了して、次のリストにいった時、強制的にmapIndexは直前の+1されたものにする

                // 書き出し位置の生成
                var mapItem = CreateOutMapItem(item.Value, pi, internalMapIndex, mapDirection);

                // 結合
                mapItem = MergeCell(pi, mapItem);

                // 書式
                SetFormat(pi, mapItem);

                // フォント
                SetFont(pi, mapItem);

                // 修飾
                SetDecoration(pi, mapItem);

                // 配置
                SetAlignment(pi, mapItem);

                // 背景
                SetBackground(pi, mapItem);

                // 罫線
                SetBorders(pi, mapItem);

                // ユーザー制御に伴う書き出しの無視(特定条件の時だけ書き出ししたくないとか)
                if (IgnoreOutputCell(item.Value, mapItem))
                {
                    continue;
                }

                // 値の書き出し
                var value = pi.GetValue(item.Value);
                SetValue(mapItem, value);

                if (mapDirection == MapDirection.Row)
                {
                    if (currentMapIndex < mapItem.EndRow)
                    {
                        currentMapIndex = mapItem.EndRow;
                    }
                }
                if (mapDirection == MapDirection.Column)
                {
                    if (currentMapIndex < mapItem.EndColumn)
                    {
                        currentMapIndex = mapItem.EndColumn;
                    }
                }
            }

            // 子要素の書き込み
            if (item.Children.Count() > 0)
            {
                var childMapIndex = internalMapIndex;
                foreach (var child in item.Children)
                {
                    Write(child, internalMapIndex, mapDirection, skip, currentMapIndex);
                    childMapIndex += skip;
                }
                internalMapIndex = childMapIndex;
            }

            //// マッピングするインデックスを加算
            //internalMapIndex = currentMapIndex;
            //if (mapDirection != MapDirection.None)
            //{
            //    internalMapIndex += skip;
            //}

            //currentMapIndex = internalMapIndex;

            return internalMapIndex;
        }

        /// <summary>
        /// 行インデックスを自動マッピングする。
        /// 未指定、もしくはマッピング方向が行でない時はマッピングしない。
        /// </summary>
        /// <param name="mapItem">書き出し情報。</param>
        /// <param name="mapIndex">マッピング位置。</param>
        /// <param name="mapDirection">マッピング方向。</param>
        private void MapRowIndex(CellOutputMapItem mapItem, int mapIndex, MapDirection mapDirection)
        {
            if (mapItem.OriginalRow != CellMapAttribute.UnspecifiedIndex)
            {
                return;
            }

            if (mapDirection == MapDirection.Row)
            {
                return;
            }

            mapItem.ChangeRow(mapIndex);
        }

        /// <summary>
        /// 列インデックスを自動マッピングする。
        /// 未指定、もしくはマッピング方向が行でない時はマッピングしない。
        /// </summary>
        /// <param name="mapItem">書き出し情報。</param>
        /// <param name="mapIndex">マッピング位置。</param>
        /// <param name="mapDirection">マッピング方向。</param>
        private void MapColumnIndex(CellOutputMapItem mapItem, int mapIndex, MapDirection mapDirection)
        {
            if (mapItem.OriginalColumn != CellMapAttribute.UnspecifiedIndex)
            {
                return;
            }

            if (mapDirection == MapDirection.Column)
            {
                return;
            }

            mapItem.ChangeColumn(mapIndex);
        }

        /// <summary>
        /// 書き出しマップ情報を生成する。
        /// </summary>
        /// <param name="obj">プロパティが含まれるオブジェクト。</param>
        /// <param name="pi">プロパティ。</param>
        /// <param name="mapIndex">マッピングする項目の書き出しインデックス。</param>
        /// <param name="mapDirection">マッピング方向。</param>
        /// <returns>書き出しマップ情報。</returns>
        private CellOutputMapItem CreateOutMapItem(object obj, PropertyInfo pi, int mapIndex, MapDirection mapDirection)
        {
            var mapAttr = pi.GetCustomAttribute<CellMapAttribute>();
            var mapItem = new CellOutputMapItem(pi.Name, mapAttr.Row, mapAttr.Column, mapAttr.Formula);
            MapRowIndex(mapItem, mapIndex, mapDirection);
            MapColumnIndex(mapItem, mapIndex, mapDirection);

            // ユーザー制御に伴うセル位置の変更
            ConfigureOutputCell(obj, mapItem);

            return mapItem;
        }

        /// <summary>
        /// ユーザー制御に伴うセル位置の変更を行う。
        /// </summary>
        /// <param name="obj">プロパティが含まれるオブジェクト。</param>
        /// <param name="mapItem">書き出しマップ情報。</param>
        private void ConfigureOutputCell(object obj, CellOutputMapItem mapItem)
        {
            // ユーザー制御に伴うセル位置の変更
            var outputConfig = obj as IOutputCellConfiguration;
            if (outputConfig != null)
            {
                outputConfig.ConfigureCell(mapItem, Param);
            }
        }

        /// <summary>
        /// ユーザー制御に伴う書き出しの無視を行う。
        /// 主に、特定条件の時だけ書き出しを行いたくないなどのユーザー要件に応じた制御。
        /// </summary>
        /// <param name="obj">プロパティが含まれるオブジェクト。</param>
        /// <param name="mapItem">書き出しマップ情報。</param>
        /// <returns>書き出しを無視する場合は true, それ以外は false。</returns>
        private bool IgnoreOutputCell(object obj, IReadOnlyCellMapItem mapItem)
        {
            // ユーザー制御に伴う書き出しの無視
            var ignoreConfig = obj as IOutputIgnoreConfiguration;
            if (ignoreConfig != null)
            {
                return ignoreConfig.IgnoreOutput(mapItem, Param);
            }

            return false;
        }




        /// <summary>
        /// 実際のセル結合を行う処理。
        /// </summary>
        /// <param name="mapItem">結合情報が反映された書き出しマップ情報。</param>
        protected abstract void MergeCell(IReadOnlyCellMapItem mapItem);


        /// <summary>
        /// 結合を行う。
        /// </summary>
        /// <param name="pi">プロパティ。</param>
        /// <param name="mapItem">書き出しマップ情報。</param>
        /// <returns>書き出しマップ情報。</returns>
        private CellOutputMapItem MergeCell(PropertyInfo pi, CellOutputMapItem mapItem)
        {
            var mergeAttr = pi.GetCustomAttribute<CellMergeAttribute>();
            if (mergeAttr == null)
            {
                return mapItem;
            }

            // 結合するセル位置を求める
            var endRow = mergeAttr.Row;
            var endColumn = mergeAttr.Column;
            if (mergeAttr.Position == CellMergePosition.Relative)
            {
                endRow = mapItem.EndRow + mergeAttr.Row;
                endColumn = mapItem.EndColumn + mergeAttr.Column;
            }

            // マージされたセル範囲を把握する
            var startRow = mapItem.StartRow;
            var startColumn = mapItem.StartColumn;

            if (endRow < startRow)
            {
                var tempRow = endRow;
                endRow = startRow;
                startRow = tempRow;
            }
            if (endColumn < startColumn)
            {
                var tempColumn = endColumn;
                endColumn = startColumn;
                startColumn = tempColumn;
            }
            mapItem.ChangeCell(startRow, startColumn, endRow, endColumn);
            MergeCell(mapItem);

            return mapItem;
        }


        /// <summary>
        /// 実際の書式設定を行う処理。
        /// </summary>
        /// <param name="mapItem"></param>
        /// <param name="formatItem">書式設定。</param>
        protected abstract void SetFormat(IReadOnlyCellMapItem mapItem, CellFormatItem formatItem);

        /// <summary>
        /// 書式設定を行う。
        /// </summary>
        /// <param name="pi"></param>
        /// <param name="mapItem"></param>
        private void SetFormat(PropertyInfo pi, IReadOnlyCellMapItem mapItem)
        {
            var formatAttr = pi.GetCustomAttribute<CellFormatAttribute>();
            if (formatAttr == null)
            {
                return;
            }

            var formatItem = new CellFormatItem(formatAttr.Format);
            SetFormat(mapItem, formatItem);
        }


        /// <summary>
        /// 実際のフォント設定を行う処理。
        /// </summary>
        /// <param name="mapItem"></param>
        /// <param name="fontItem">フォント設定。</param>
        protected abstract void SetFont(IReadOnlyCellMapItem mapItem, CellFontItem fontItem);

        /// <summary>
        /// フォント設定を行う。
        /// </summary>
        /// <param name="pi"></param>
        /// <param name="mapItem"></param>
        private void SetFont(PropertyInfo pi, CellOutputMapItem mapItem)
        {
            var fontAttr = pi.GetCustomAttribute<CellFontAttribute>();
            if (fontAttr == null)
            {
                return;
            }

            var fontItem = new CellFontItem(fontAttr.FamilyName, fontAttr.Size, fontAttr.FontStyle, fontAttr.Color);
            SetFont(mapItem, fontItem);
        }


        /// <summary>
        /// 実際の修飾設定を行う処理。
        /// </summary>
        /// <param name="mapItem"></param>
        /// <param name="characterDecorationItem">文字修飾設定。</param>
        protected abstract void SetDecoration(IReadOnlyCellMapItem mapItem, CellCharacterDecorationItem characterDecorationItem);

        /// <summary>
        /// 文字修飾設定を行う。
        /// </summary>
        /// <param name="pi"></param>
        /// <param name="mapItem"></param>
        private void SetDecoration(PropertyInfo pi, CellOutputMapItem mapItem)
        {
            var decorationAttr = pi.GetCustomAttribute<CellCharacterDecorationAttribute>();
            if (decorationAttr == null)
            {
                return;
            }

            var decorationItem = new CellCharacterDecorationItem(decorationAttr.UnderlineStyle, decorationAttr.CharacterPosition);
            SetDecoration(mapItem, decorationItem);
        }


        /// <summary>
        /// 実際の配置設定を行う処理。
        /// </summary>
        /// <param name="mapItem"></param>
        /// <param name="alignmentItem">配置設定。</param>
        protected abstract void SetAlignment(IReadOnlyCellMapItem mapItem, CellAlignmentItem alignmentItem);

        /// <summary>
        /// 配置設定を行う。
        /// </summary>
        /// <param name="pi"></param>
        /// <param name="mapItem"></param>
        private void SetAlignment(PropertyInfo pi, CellOutputMapItem mapItem)
        {
            var alignmentAttr = pi.GetCustomAttribute<CellAlignmentAttribute>();
            if (alignmentAttr == null)
            {
                return;
            }

            var alignmentItem = new CellAlignmentItem(alignmentAttr.Horizontal, alignmentAttr.Vertical);
            SetAlignment(mapItem, alignmentItem);
        }


        /// <summary>
        /// 実際の背景設定を行う処理。
        /// </summary>
        /// <param name="mapItem"></param>
        /// <param name="backgroundItem">背景設定。</param>
        protected abstract void SetBackground(IReadOnlyCellMapItem mapItem, CellBackgroundItem backgroundItem);

        /// <summary>
        /// 背景設定を行う。
        /// </summary>
        /// <param name="pi"></param>
        /// <param name="mapItem"></param>
        private void SetBackground(PropertyInfo pi, IReadOnlyCellMapItem mapItem)
        {
            var backgroundAttr = pi.GetCustomAttribute<CellBackgroundAttribute>();
            if (backgroundAttr == null)
            {
                return;
            }

            var backgroundItem = new CellBackgroundItem(backgroundAttr.Color);
            SetBackground(mapItem, backgroundItem);
        }


        /// <summary>
        /// 実際の罫線設定を行う処理。
        /// </summary>
        /// <param name="mapItem"></param>
        /// <param name="borderItem">罫線設定。</param>
        protected abstract void SetBorder(IReadOnlyCellMapItem mapItem, CellBorderItem borderItem);

        /// <summary>
        /// 罫線設定を行う。
        /// </summary>
        /// <param name="pi"></param>
        /// <param name="mapItem"></param>
        private void SetBorders(PropertyInfo pi, IReadOnlyCellMapItem mapItem)
        {
            var borderAttrs = pi.GetCustomAttributes<CellBorderAttribute>();

            foreach (var borderAttr in borderAttrs)
            {
                var borderItem = new CellBorderItem(borderAttr.Position, borderAttr.Style, borderAttr.Weight, borderAttr.Color);
                SetBorder(mapItem, borderItem);
            }
        }

        /// <summary>
        /// 実際の値設定を行う処理。
        /// </summary>
        /// <param name="mapItem"></param>
        /// <param name="value">書き出す値。</param>
        protected abstract void SetValue(IReadOnlyCellMapItem mapItem, object value);







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
