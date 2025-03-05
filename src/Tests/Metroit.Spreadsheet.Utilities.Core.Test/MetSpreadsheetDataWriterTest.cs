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

            writer.Write(value);
        }
    }

    class DummyOperator : MetSpreadsheetOperatorBase
    {
        public DummyOperator() : base()
        {

        }
    }

    class Item
    {
        [CellOutputMap(Column = 0)]
        public char CharValue { get; set; }

        [CellOutputMap(Column = 0)]
        public char[] CharValues { get; set; }

        [CellOutputMap(Column = 0)]
        public string StringValue { get; set; }

        [CellOutputMap(Column = 2)]
        public int? IntValue { get; set; }

        [CellOutputMap(Column = 4)]
        public string[] ArrayStringValues { get; set; }

        [CellOutputMap(Column = 3)]
        public Dictionary<int, string> DictionaryValues { get; set; }

        [CellOutputMap(Column = 1)]
        public List<string> ListStringValues { get; set; }

        public List<ChildItem> ListChildItems { get; set; }

        public ChildItem[] ArrayChildItems { get; set; }

        public ChildItem ChildItem { get; set; }

        public GrandChildItem GrandChildItem { get; set; }

        public NonGrandChildItem NonGrandChildItem { get; set; }

        public GrandChildItem2 GrandChildItem2 { get; set; }
    }

    class ChildItem
    {
        [CellOutputMap(Column = 10)]
        public string ChildHoge { get; set; }
    }

    class GrandChildItem
    {
        public GreatGrandChildItem GreatGrandChildItem { get; set; }
    }

    class GreatGrandChildItem
    {
        [CellOutputMap(Column = 11)]
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
