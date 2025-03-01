using System;
using Xunit;

namespace Metroit.Spreadsheet.Utilities.Core.Test
{
    [Trait("TestCategory", "MetSsConvertTest")]
    public class MetSsConvertTest
    {
        [Theory(DisplayName = "RcToA1()が実行可能かどうか")]
        [InlineData("row", -1, 0)]
        [InlineData("column", 0, -1)]
        public void TestRcToA1Executable(string expected, int row, int column)
        {
            Assert
                .Throws<ArgumentOutOfRangeException>(() => MetSsConvert.RcToA1(row, column))
                .Message
                .Is($"指定された引数は、有効な値の範囲内にありません。\r\nパラメーター名:{expected}");
        }

        [Theory(DisplayName = "RcToA1()の実行")]
        [InlineData("A1", 0, 0)]
        [InlineData("Z1048576", 1048575, 25)]
        [InlineData("BA1048576", 1048575, 52)]
        [InlineData("BZ1048576", 1048575, 77)]
        [InlineData("XFD1048576", 1048575, 16383)]
        [InlineData("FXSHRXX1048576", 1048575, int.MaxValue)]
        public void TestRcToA1(string expected, int row, int column)
        {
            MetSsConvert.RcToA1(row, column).Is(expected);
        }

        [Fact(DisplayName = "ToColumnA1()が実行可能かどうか")]
        public void TestToColumnA1Executable()
        {
            Assert
                .Throws<ArgumentOutOfRangeException>(() => MetSsConvert.ToColumnA1(-1))
                .Message
                .Is("指定された引数は、有効な値の範囲内にありません。\r\nパラメーター名:index");
        }

        [Theory(DisplayName = "ToColumnA1()の実行")]
        [InlineData("A", 0)]
        [InlineData("B", 1)]
        [InlineData("Z", 25)]
        [InlineData("AA", 26)]
        [InlineData("AB", 27)]
        [InlineData("AZ", 51)]
        [InlineData("BA", 52)]
        [InlineData("BZ", 77)]
        [InlineData("CA", 78)]
        [InlineData("XFD", 16383)]
        [InlineData("FXSHRXX", int.MaxValue)]
        public void TestToColumnA1(string expect, int index)
        {
            MetSsConvert.ToColumnA1(index).Is(expect);
        }

        [Theory(DisplayName = "ToColumnIndex()が実行可能かどうか")]
        [InlineData(typeof(ArgumentNullException), "値を Null にすることはできません。\r\nパラメーター名:value", null)]
        [InlineData(typeof(ArgumentException), "value is empty.", "")]
        [InlineData(typeof(FormatException), "Argument format is only alphabet", "@A")]
        [InlineData(typeof(ArgumentOutOfRangeException), "Argument range max \"FXSHRXX\"\r\nパラメーター名:value", "FXSHRXY")]
        [InlineData(typeof(ArgumentOutOfRangeException), "Argument range max \"FXSHRXX\"\r\nパラメーター名:value", "AAAAAAAA")]
        public void TestToColumnIndexExecutable(Type exceptionType, string expected, string value)
        {
            Assert.Throws(exceptionType, 
                () => MetSsConvert.ToColumnIndex(value))
                .Message
                .Is(expected);
        }

        [Theory(DisplayName = "ToColumnIndex()の実行")]
        [InlineData(0, "A")]
        [InlineData(25, "Z")]
        [InlineData(26, "AA")]
        [InlineData(27, "AB")]
        [InlineData(51, "AZ")]
        [InlineData(52, "BA")]
        [InlineData(77, "BZ")]
        [InlineData(78, "CA")]
        [InlineData(702, "AAA")]
        [InlineData(4341, "FJZ")]
        [InlineData(16383, "XFD")]
        [InlineData(int.MaxValue, "FXSHRXX")]
        public void TestToColumnIndex(int expected, string value)
        {
            MetSsConvert.ToColumnIndex(value).Is(expected);
        }
    }
}
