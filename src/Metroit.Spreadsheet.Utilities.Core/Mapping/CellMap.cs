using Metroit.Spreadsheet.Utilities.Core.MapItem;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Metroit.Spreadsheet.Utilities.Core.Mapping
{
    /// <summary>
    /// Excelとのセルマップを提供します。
    /// Excel操作オブジェクトを利用した実際の読み込み、書き込みは呼出コードに依存します。
    /// </summary>
    public class CellMap
    {
        /// <summary>
        /// Read() によって読み込みマップを編集するメソッド名。
        /// </summary>
        private const string InputMapConfigureMethodName = "ConfigureInputMap";

        /// <summary>
        /// Write() によって書き込みマップを編集するメソッド名。
        /// </summary>
        private const string OutputMapConfigureMethodName = "ConfigureOutputMap";

        /// <summary>
        /// 入力セルマップを有するプロパティに値を読み込みます。
        /// </summary>
        /// <typeparam name="T">インスタンス化可能なクラス。</typeparam>
        /// <param name="itemRead">Excelオブジェクトを利用したデータ読み込みコード。</param>
        /// <param name="param">実行パラメーター。</param>
        /// <returns>T オブジェクト。</returns>
        public T Read<T>(Func<CellInputItemInfo, object, object> itemRead, object param = null) where T : new()
        {
            var entity = new T();
            ReadValues(entity, itemRead, -1, MapDirection.None, param);
            return entity;
        }

        /// <summary>
        /// 入力セルマップを有するプロパティに値を読み込みます。
        /// </summary>
        /// <typeparam name="T">インスタンス化可能なクラス。</typeparam>
        /// <param name="itemRead">Excelオブジェクトを利用したデータ読み込みコード。</param>
        /// <param name="mapIndex">マップするインデックス。</param>
        /// <param name="mapDirection">マップする方向が行か列か。</param>
        /// <param name="param">実行パラメーター。</param>
        /// <returns>T オブジェクト。</returns>
        public T Read<T>(Func<CellInputItemInfo, object, object> itemRead, int mapIndex, MapDirection mapDirection = MapDirection.Row, object param = null) where T : new()
        {
            var entity = new T();
            ReadValues(entity, itemRead, mapIndex, mapDirection, param);
            return entity;
        }

        /// <summary>
        /// 入力セルマップを有するプロパティに値を読み込みます。
        /// </summary>
        /// <param name="entity">入力オブジェクト。</param>
        /// <param name="itemRead">Excelオブジェクトを利用したデータ読み込みコード。</param>
        /// <param name="param">実行パラメーター。</param>
        public void Read(object entity, Func<CellInputItemInfo, object, object> itemRead, object param = null)
        {
            ReadValues(entity, itemRead, -1, MapDirection.None, param);
        }

        /// <summary>
        /// 入力セルマップを有するプロパティに値を読み込みます。
        /// </summary>
        /// <param name="entity">入力オブジェクト。</param>
        /// <param name="itemRead">Excelオブジェクトを利用したデータ読み込みコード。</param>
        /// <param name="mapIndex">マップするインデックス。</param>
        /// <param name="mapDirection">マップする方向が行か列か。</param>
        /// <param name="param">実行パラメーター。</param>
        public void Read(object entity, Func<CellInputItemInfo, object, object> itemRead, int mapIndex, MapDirection mapDirection = MapDirection.Row, object param = null)
        {
            ReadValues(entity, itemRead, mapIndex, mapDirection, param);
        }

        /// <summary>
        /// 入力セルマップを有するプロパティに値を読み込む。
        /// </summary>
        /// <param name="entity">入力オブジェクト。</param>
        /// <param name="itemRead">Excelオブジェクトを利用したデータ読み込みコード。</param>
        /// <param name="mapIndex">マップするインデックス。</param>
        /// <param name="mapDirection">マップする方向が行か列か。</param>
        /// <param name="param">実行パラメーター。</param>
        private void ReadValues(object entity, Func<CellInputItemInfo, object, object> itemRead, int mapIndex, MapDirection mapDirection, object param)
        {
            // オブジェクトからプロパティを取得
            var t = entity.GetType();
            var pis = t.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);

            foreach (var pi in pis)
            {
                // 入力セルマップを有していないプロパティは実施しない
                var mapAttr = Attribute.GetCustomAttribute(pi, typeof(CellInputMapAttribute)) as CellInputMapAttribute;
                if (mapAttr == null)
                {
                    continue;
                }

                var itemInfo = new CellInputItemInfo(pi,
                    new CellInputMapItem(pi.Name, mapAttr.Row, mapAttr.Column, mapAttr.Formula)
                    );

                if (mapDirection == MapDirection.Row)
                {
                    itemInfo.Cell.ChangeRow(mapIndex);
                }
                if (mapDirection == MapDirection.Column)
                {
                    itemInfo.Cell.ChangeColumn(mapIndex);
                }

                // 編集メソッドがある場合は呼出を行う
                var mi = t.GetMethod(InputMapConfigureMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
                if (mi != null)
                {
                    mi.Invoke(entity, new object[] { itemInfo, param });
                }

                // シートから値の取得を行う
                object cellValue = itemRead(itemInfo, param);

                // 対象セルの値がない場合は空とする(型がstringの場合は空文字とする)
                if (cellValue == null || cellValue.ToString() == "")
                {
                    if (itemInfo.MapProperty.PropertyType == typeof(string))
                    {
                        pi.SetValue(entity, "", null);
                    }
                    else
                    {
                        pi.SetValue(entity, null, null);
                    }
                    continue;
                }

                // Nullable<T>に対応した元の型を取得し、プロパティ値を設定する
                var safeType = Nullable.GetUnderlyingType(itemInfo.MapProperty.PropertyType) ?? itemInfo.MapProperty.PropertyType;
                if (safeType == typeof(DateTime))
                {
                    pi.SetValue(entity, DateTime.FromOADate(Convert.ToDouble(cellValue)), null);
                }
                else
                {
                    pi.SetValue(entity, Convert.ChangeType(cellValue, safeType), null);
                }
            }
        }

        /// <summary>
        /// 出力セルマップを有するプロパティの値をExcelに書き込みます。
        /// </summary>
        /// <param name="entity">出力オブジェクト。</param>
        /// <param name="itemWrite">Excelオブジェクトを利用したデータ書き込みコード。</param>
        /// <param name="param">実行パラメーター。</param>
        public void Write(object entity, Action<CellOutputItemInfo, object, object> itemWrite, object param = null)
        {
            WriteValues(entity, itemWrite, -1, MapDirection.None, param);
        }

        /// <summary>
        /// 出力セルマップを有するプロパティの値をExcelに書き込みます。
        /// </summary>
        /// <param name="entity">出力オブジェクト。</param>
        /// <param name="itemWrite">Excelオブジェクトを利用したデータ書き込みコード。</param>
        /// <param name="mapIndex">マップするインデックス。</param>
        /// <param name="mapDirection">マップする方向が行か列か。</param>
        /// <param name="param">実行パラメーター。</param>
        public void Write(object entity, Action<CellOutputItemInfo, object, object> itemWrite, int mapIndex, MapDirection mapDirection = MapDirection.Row, object param = null)
        {
            WriteValues(entity, itemWrite, mapIndex, mapDirection, param);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="mapItem"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        private CellOutputMapItem ConfigureOutputCell(object entity, CellOutputMapItem mapItem, object param)
        {
            // ユーザー制御に伴うセル位置の変更
            var outputConfig = entity as IOutputCellConfiguration;
            if (outputConfig != null)
            {
                outputConfig.ConfigureCell(mapItem, param);
            }

            return mapItem;
        }

        private bool IgnoreOutputCell(object entity, IReadOnlyCellMapItem mapItem, object param)
        {
            // ユーザー制御に伴う書き出しの無視
            var ignoreConfig = entity as IOutputIgnoreConfiguration;
            if (ignoreConfig.IgnoreOutput(mapItem, param))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 出力セルマップを有するプロパティの値をExcelに書き込みます。
        /// マップのカスタム編集は行われません。
        /// </summary>
        /// <param name="entity">出力オブジェクト。</param>
        /// <param name="itemWrite">Excelオブジェクトを利用したデータ書き込みコード。</param>
        /// <param name="mapIndex">マップするインデックス。</param>
        /// <param name="mapDirection">マップする方向が行か列か。</param>
        /// <param name="param">実行パラメーター。</param>
        private void WriteValues(object entity, Action<CellOutputItemInfo, object, object> itemWrite,
            int mapIndex, MapDirection mapDirection, object param)
        {
            // オブジェクトからプロパティを取得
            var t = entity.GetType();
            var pis = t.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);

            foreach (var pi in pis)
            {
                // 出力セルマップを有していないプロパティは実施しない
                var mapAttr = Attribute.GetCustomAttribute(pi, typeof(CellMapAttribute)) as CellMapAttribute;
                if (mapAttr == null)
                {
                    continue;
                }

                // 属性から得られたセル情報
                var mapItem = new CellOutputMapItem(pi.Name, mapAttr.Row, mapAttr.Column, mapAttr.Formula);

                // ユーザー制御に伴うセル位置の変更
                mapItem = ConfigureOutputCell(entity, mapItem, param);

                // ユーザー制御に伴う書き出しの無視
                if (IgnoreOutputCell(entity, mapItem, param))
                {
                    continue;
                }

                // 結合
                var mergeAttr = Attribute.GetCustomAttribute(pi, typeof(CellMergeAttribute)) as CellMergeAttribute;
                CellMergeItem mergeItem = null;
                if (mergeAttr != null)
                {
                    var endRow = mergeAttr.Row;
                    var endColumn = mergeAttr.Column;
                    if (mergeAttr.Position == CellMergePosition.Relative)
                    {
                        //endRow = mapItem.Row + mergeAttr.Row;
                        //endColumn = mapItem.Column + mergeAttr.Column;
                    }

                    mergeItem = new CellMergeItem(mergeAttr.Row, mergeAttr.Column, mergeAttr.Position, endRow, endColumn);
                }







                // 書式
                var formatAttr = Attribute.GetCustomAttribute(pi, typeof(CellFormatAttribute)) as CellFormatAttribute;
                CellFormatItem formatItem = null;
                if (formatAttr != null)
                {
                    formatItem = new CellFormatItem(formatAttr.Format);
                }

                // フォント
                var fontAttr = Attribute.GetCustomAttribute(pi, typeof(CellFontAttribute)) as CellFontAttribute;
                CellFontItem fontItem = null;
                if (fontAttr != null)
                {
                    //fontItem = new CellFontItem(fontAttr.GetFont(), fontAttr.FontStyle, fontAttr.Size, fontAttr.GetColor());
                }

                var decorationAttr = Attribute.GetCustomAttribute(pi, typeof(CellCharacterDecorationAttribute)) as CellCharacterDecorationAttribute;
                CellCharacterDecorationItem decorationItem = null;
                if (decorationAttr != null)
                {
                    decorationItem = new CellCharacterDecorationItem(decorationAttr.UnderlineStyle, decorationAttr.CharacterPosition);
                }

                // 配置
                var alignAttr = Attribute.GetCustomAttribute(pi, typeof(CellAlignmentAttribute)) as CellAlignmentAttribute;
                CellAlignmentItem alignItem = null;
                if (alignAttr != null)
                {
                    alignItem = new CellAlignmentItem(alignAttr.Horizontal, alignAttr.Vertical);
                }

                // 背景
                var backgroundAttr = Attribute.GetCustomAttribute(pi, typeof(CellBackgroundAttribute)) as CellBackgroundAttribute;
                CellBackgroundItem backgroundItem = null;
                if (backgroundAttr != null)
                {
                    //backgroundItem = new CellBackgroundItem(backgroundAttr.GetColor());
                }

                // 罫線
                var borderAttrs = Attribute.GetCustomAttributes(pi, typeof(CellBorderAttribute));
                var borderItems = new List<CellBorderItem>();
                if (borderAttrs.Length > 0)
                {
                    foreach (var borderAttr in borderAttrs.Cast<CellBorderAttribute>())
                    {
                        //borderItems.Add(new CellBorderItem(borderAttr.Position, borderAttr.Style, borderAttr.Weight, borderAttr.GetColor()));
                    }
                }

                var itemInfo = new CellOutputItemInfo(pi.Name,
                    mapItem,
                    mergeItem,
                    formatItem,
                    fontItem,
                    decorationItem,
                    alignItem,
                    backgroundItem,
                    borderItems
                    );

                if (mapDirection == MapDirection.Row)
                {
                    itemInfo.Cell.ChangeRow(mapIndex);
                }
                if (mapDirection == MapDirection.Column)
                {
                    itemInfo.Cell.ChangeColumn(mapIndex);
                }

                // プロパティ指定でない時、マップのカスタム編集メソッドがある場合は呼出を行う
                //if (ignoreProperties == null)
                //{
                //    var mi = t.GetMethod(OutputMapConfigureMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
                //    if (mi != null)
                //    {
                //        mi.Invoke(entity, new object[] { itemInfo, param });
                //    }
                //}


                var value = pi.GetValue(entity, null);
                itemWrite.Invoke(itemInfo, param, value);
            }
        }
    }
}
