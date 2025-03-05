using System;
using Xunit;

namespace Metroit.Spreadsheet.Utilities.Core.Test
{
    public class MetSpreadsheetConvertTest
    {
        [Fact(DisplayName = "ToA1()が実行可能ではない")]
        public void TestToA1Executable()
        {
            Assert
                .Throws<ArgumentException>(() => MetSpreadsheetConvert.ToA1(-1, -1))
                .Message
                .Is($"Either row or column must be 0 or greater.");
        }

        [Theory(DisplayName = "ToA1()の実行")]
        [InlineData("A1", 0, 0)]
        [InlineData("Z1048576", 1048575, 25)]
        [InlineData("BA1048576", 1048575, 52)]
        [InlineData("BZ1048576", 1048575, 77)]
        [InlineData("XFD1048576", 1048575, 16383)]
        [InlineData("FXSHRXX1048576", 1048575, int.MaxValue)]
        [InlineData("A", -1, 0)]
        [InlineData("1", 0, -1)]
        public void TestToA1(string expected, int row, int column)
        {
            MetSpreadsheetConvert.ToA1(row, column).Is(expected);
        }

        [Fact(DisplayName = "ToColumnA1()が実行可能ではない")]
        public void TestToColumnA1Executable()
        {
            Assert
                .Throws<ArgumentOutOfRangeException>(() => MetSpreadsheetConvert.ToColumnA1(-1))
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
            MetSpreadsheetConvert.ToColumnA1(index).Is(expect);
        }

        [Theory(DisplayName = "ToColumnIndex()が実行可能ではない")]
        [InlineData(typeof(ArgumentNullException), "値を Null にすることはできません。\r\nパラメーター名:value", null)]
        [InlineData(typeof(ArgumentException), "value is empty.", "")]
        [InlineData(typeof(FormatException), "Argument format is only alphabet", "@A")]
        [InlineData(typeof(ArgumentOutOfRangeException), "Argument range max \"FXSHRXX\"\r\nパラメーター名:value", "FXSHRXY")]
        [InlineData(typeof(ArgumentOutOfRangeException), "Argument range max \"FXSHRXX\"\r\nパラメーター名:value", "AAAAAAAA")]
        public void TestToColumnIndexExecutable(Type exceptionType, string expected, string value)
        {
            Assert.Throws(exceptionType,
                () => MetSpreadsheetConvert.ToColumnIndex(value))
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
            MetSpreadsheetConvert.ToColumnIndex(value).Is(expected);
        }

        [Theory(DisplayName = "ToRange()が実行可能ではない")]
        [InlineData("Either row or column must be 0 or greater.", -1, -1, -1, -1)]
        [InlineData("row must be specified.", -1, 0, 0, 0)]
        [InlineData("column must be specified.", 0, -1, 0, 0)]
        public void TestToRangeExecutable(string expected, int row1, int column1, int row2, int column2)
        {
            Assert
                .Throws<ArgumentException>(() => MetSpreadsheetConvert.ToRange(row1, column1, row2, column2))
                .Message
                .Is(expected);
        }

        [Theory(DisplayName = "ToRange()の実行")]
        [InlineData("A1:A2", 1, 0, 0, 0)]
        [InlineData("A1:B1", 0, 1, 0, 0)]
        [InlineData("A1:B2", 1, 1, 0, 0)]
        [InlineData("1", 0, -1, 0, -1)]
        [InlineData("1:2", 0, -1, 1, -1)]
        [InlineData("A", -1, 0, -1, 0)]
        [InlineData("A:B", -1, 0, -1, 1)]
        [InlineData("A1", 0, 0, 0, 0)]
        public void TestToRange(string expected, int row1, int column1, int row2, int column2)
        {
            MetSpreadsheetConvert.ToRange(row1, column1, row2, column2).Is(expected);
        }

        [Theory(DisplayName = "ToCells()が実行可能ではない")]
        [InlineData("@1:A2")]
        [InlineData("@1")]
        [InlineData("1A")]
        [InlineData("1A:1A")]
        [InlineData("A1:A")]
        [InlineData("1:A")]
        public void TestToCellsExecutable(string range)
        {
            Assert
                .Throws<FormatException>(() => MetSpreadsheetConvert.ToCells(range))
                .Message
                .Is("The cell range format is incorrect.");
        }

        [Theory(DisplayName = "ToCells()の実行")]
        [InlineData(0, 0, 1, 0, "A1:A2")]
        [InlineData(0, 0, 0, 0, "A1")]
        [InlineData(-1, 0, -1, 0, "A")]
        [InlineData(-1, 0, -1, 1, "A:B")]
        [InlineData(0, -1, 0, -1, "1")]
        [InlineData(0, -1, 1, -1, "1:2")]
        public void TestToCells(int expectedRow1, int expectedColumn1, int expectedRow2, int expectedColumn2, string range)
        {
            MetSpreadsheetConvert.ToCells(range).Is((expectedRow1, expectedColumn1, expectedRow2, expectedColumn2));
        }
    }
}
