using System;
using System.Collections.Generic;
using System.Runtime.Remoting;
using Xunit;

namespace Metroit.Spreadsheet.Utilities.Core.Tests
{
    [Trait("Category", "R1C1形式の値からA1形式の文字列に変換する")]
    public class MetSsConvertTest
    {
        [Fact(DisplayName = "実行可能なパラメーターかどうか")]
        [Trait("Category", "R1C1形式の値からA1形式の文字列に変換する")]
        public void TestCase001()
        {
            Assert
                .Throws<ArgumentOutOfRangeException>(() => MetSsConvert.RcToA1(-1, 0))
                .Message
                .Is("指定された引数は、有効な値の範囲内にありません。\r\nパラメーター名:row");

            Assert
                .Throws<ArgumentOutOfRangeException>(() => MetSsConvert.RcToA1(0, -1))
                .Message
                .Is("指定された引数は、有効な値の範囲内にありません。\r\nパラメーター名:column");
        }

        [Fact(DisplayName = "A - Z の変換を行う")]
        public void TestCase002()
        {
            string a1;

            a1 = MetSsConvert.RcToA1(0, 0);
            a1.Is("A1");

            a1 = MetSsConvert.RcToA1(1048575, 25);
            a1.Is("Z1048576");
        }

        [Fact(DisplayName = "BA - BZ の変換を行う")]
        public void TestCase003()
        {
            string a1;

            a1 = MetSsConvert.RcToA1(1048575, 52);
            a1.Is("BA1048576");

            a1 = MetSsConvert.RcToA1(1048575, 77);
            a1.Is("BZ1048576");
        }

        [Fact(DisplayName = "xlsx の最大セル XFD1048576の変換を行う")]
        public void TestCase004()
        {
            string a1;

            a1 = MetSsConvert.RcToA1(1048575, 16383);
            a1.Is("XFD1048576");
        }

        [Fact(DisplayName = "機能が許容する最大セル XFD1048576の変換を行う")]
        public void TestCase005()
        {
            string a1;

            a1 = MetSsConvert.RcToA1(1048575, int.MaxValue);
            a1.Is("FXSHRXX1048576");
        }
    }
}
