using Metroit.Spreadsheet.Utilities.Core.MapItem;
using Metroit.Spreadsheet.Utilities.Core.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Metroit.Spreadsheet.Utilities.Core.Test
{
    public class MetSpreadsheetDataWriterTest
    {
        [Fact(DisplayName = "Test")]
        public void Test()
        {
            var writer = new MetSpreadsheetDataWriter<DummyOperator>();

            var value = new Item();
            value.CharValue = 'c';
            value.CharValues = new[] { 'c', 'c' };
            value.StringValue = "StringValue";
            value.IntValue = -123;
            value.ArrayStringValues = new string[] { "ArrayString1", "ArrayString2", "ArrayString3" };
            value.DictionaryValues = new Dictionary<int, string>()
            {
                { 1, "Dic1" },
                { 2, "Dic2" },
                { 3, "Dic3" }
            };
            value.DictionaryItems = new Dictionary<int, ChildItem>()
            {
                {1, new ChildItem(){ ChildHoge ="DictionaryItem1"} },
                {2, new ChildItem(){ ChildHoge ="DictionaryItem2"} },
                {3, new ChildItem(){ ChildHoge ="DictionaryItem3"} }
            };
            value.ListStringValues = new List<string>()
            {
                "ListString1",
                "ListString2",
                "ListString3"
            };
            value.ListChildItems = new List<ChildItem>()
            {
                new ChildItem(){ ChildHoge ="ListChildItems1" },
                new ChildItem(){ ChildHoge ="ListChildItems2" }
            };
            value.ArrayChildItems = new ChildItem[] { new ChildItem() { ChildHoge = "ArrayChildItems1" }, new ChildItem() { ChildHoge = "ArrayChildItems2" } };
            value.ChildItem = new ChildItem() { ChildHoge = "ChildItem1" };
            //value.ChildItem = null;

            value.GrandChildItem = new GrandChildItem() { GreatGrandChildItem = new GreatGrandChildItem() { GreatGrandChildHoge = "GreatGrandChildHoge1" } };
            value.NonGrandChildItem = new NonGrandChildItem() { NonGreatGrandChildItem = new NonGreatGrandChildItem() { NonGreatGrandChildHoge = "NonGreatGrandChildHoge1" } };
            value.GrandChildItem2 = new GrandChildItem2()
            {
                GreatGrandChildItem2 = new GreatGrandChildItem2()
                {
                    GreatGrandChildItem2ChildItems = new List<ChildItem>()
                    {
                        new ChildItem() { ChildHoge = "GreatGrandChildItem2ChildItems1" },
                        new ChildItem() { ChildHoge = "GreatGrandChildItem2ChildItems2" }
                    }
                }
            };

            // これは通用する。取得するところのインデックスがちがくても、ちゃんと取ってこれる。同じクラス情報という認識を持てている
            //var pi = value.ListChildItems[0].GetType().GetProperty("ChildHoge");
            //pi.GetValue(value.ListChildItems[1]);
;
;
            writer.Write(value);
        }
    }

    class DummyOperator : MetSpreadsheetOperatorBase
    {
        public DummyOperator() : base()
        {

        }

        protected override void MergeCell(IReadOnlyCellMapItem mapItem)
        {
            
        }

        protected override void OnRead<T>()
        {
            
        }

        protected override void OnWrite()
        {
            
        }

        protected override void SetAlignment(IReadOnlyCellMapItem mapItem, CellAlignmentItem alignmentItem)
        {
            
        }

        protected override void SetBackground(IReadOnlyCellMapItem mapItem, CellBackgroundItem backgroundItem)
        {
            
        }

        protected override void SetBorder(IReadOnlyCellMapItem mapItem, CellBorderItem borderItem)
        {
            
        }

        protected override void SetDecoration(IReadOnlyCellMapItem mapItem, CellCharacterDecorationItem characterDecorationItem)
        {
            
        }

        protected override void SetFont(IReadOnlyCellMapItem mapItem, CellFontItem fontItem)
        {
            
        }

        protected override void SetFormat(IReadOnlyCellMapItem mapItem, CellFormatItem formatItem)
        {
            
        }

        protected override void SetValue(IReadOnlyCellMapItem mapItem, object value)
        {
            
        }
    }

    class Item : IOutputCellConfiguration, IOutputIgnoreConfiguration
    {
        [CellMap(Column = 0)]
        public char CharValue { get; set; }

        [CellMap(Column = 0)]
        [MapShift()]
        public char[] CharValues { get; set; }

        [CellMap(Column = 0)]
        public string StringValue { get; set; }

        [CellMap(Column = 2)]
        public int? IntValue { get; set; }

        [CellMap(Column = 4)]
        [MapShift()]
        public string[] ArrayStringValues { get; set; }

        [CellMap(Column = 3)]
        [MapShift()]
        public Dictionary<int, string> DictionaryValues { get; set; }

        public Dictionary<int, ChildItem> DictionaryItems { get; set; }

        [CellMap(Column = 1)]
        [MapShift()]
        public List<string> ListStringValues { get; set; }

        [MapShift()]
        public List<ChildItem> ListChildItems { get; set; }

        [MapShift()]
        public ChildItem[] ArrayChildItems { get; set; }

        public ChildItem ChildItem { get; set; }

        public GrandChildItem GrandChildItem { get; set; }

        public NonGrandChildItem NonGrandChildItem { get; set; }

        public GrandChildItem2 GrandChildItem2 { get; set; }

        public void ConfigureCell(CellOutputMapItem map, object param)
        {
            Console.WriteLine($"ConfigureCell!: {map.Name}");
        }

        public bool IgnoreOutput(IReadOnlyCellMapItem map, object param)
        {
            Console.WriteLine($"IgnoreOutput!: {map.Name}");
            return false;
        }
    }

    class ChildItem
    {
        [CellMap(Column = 10)]
        public string ChildHoge { get; set; }
    }

    class GrandChildItem
    {
        public GreatGrandChildItem GreatGrandChildItem { get; set; }
    }

    class GreatGrandChildItem
    {
        [CellMap(Column = 11)]
        public string GreatGrandChildHoge { get; set; }
    }


    class NonGrandChildItem
    {
        public NonGreatGrandChildItem NonGreatGrandChildItem { get; set; }
    }

    class NonGreatGrandChildItem
    {
        public string NonGreatGrandChildHoge { get; set; }
    }

    class GrandChildItem2
    {
        public GreatGrandChildItem2 GreatGrandChildItem2 { get; set; }
    }

    class GreatGrandChildItem2
    {
        public List<ChildItem> GreatGrandChildItem2ChildItems { get; set; }
    }
}
