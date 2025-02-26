using System;
using Xunit;

namespace Metroit.Spreadsheet.Utilities.Core.XUnitTests
{
    public class UnitTest1
    {
        [Fact]
        public void R1C1�`���̒l����A1�`���̕�����ɕϊ�����()
        {
            Assert
                .Throws<ArgumentOutOfRangeException>(() => MetSsConvert.RcToA1(-1, 0))
                .Message
                .Is("�w�肳�ꂽ�����́A�L���Ȓl�͈͓̔��ɂ���܂���B\r\n�p�����[�^�[��:row");

            Assert
                .Throws<ArgumentOutOfRangeException>(() => MetSsConvert.RcToA1(0, -1))
                .Message
                .Is("�w�肳�ꂽ�����́A�L���Ȓl�͈͓̔��ɂ���܂���B\r\n�p�����[�^�[��:column");

            string a1;

            a1 = MetSsConvert.RcToA1(0, 0);
            a1.Is("A1");

            a1 = MetSsConvert.RcToA1(1048575, 25);
            a1.Is("Z1048576");

            a1 = MetSsConvert.RcToA1(1048575, 52);
            a1.Is("BA1048576");

            a1 = MetSsConvert.RcToA1(1048575, 16383);
            a1.Is("XFD1048576");

            a1 = MetSsConvert.RcToA1(1048575, int.MaxValue);
            a1.Is("FXSHRXX1048576");
        }

        [Fact]
        public void ��C���f�b�N�X����A1�`���̗񕶎���ɕϊ�����()
        {
            Assert
                .Throws<ArgumentOutOfRangeException>(() => MetSsConvert.ToColumnA1(-1))
                .Message
                .Is("�w�肳�ꂽ�����́A�L���Ȓl�͈͓̔��ɂ���܂���B\r\n�p�����[�^�[��:index");

            string name;

            name = MetSsConvert.ToColumnA1(0);
            name.Is("A");

            name = MetSsConvert.ToColumnA1(1);
            name.Is("B");

            name = MetSsConvert.ToColumnA1(25);
            name.Is("Z");

            name = MetSsConvert.ToColumnA1(26);
            name.Is("AA");

            name = MetSsConvert.ToColumnA1(27);
            name.Is("AB");

            name = MetSsConvert.ToColumnA1(51);
            name.Is("AZ");

            name = MetSsConvert.ToColumnA1(52);
            name.Is("BA");

            name = MetSsConvert.ToColumnA1(77);
            name.Is("BZ");

            name = MetSsConvert.ToColumnA1(78);
            name.Is("CA");

            name = MetSsConvert.ToColumnA1(16383);
            name.Is("XFD");

            name = MetSsConvert.ToColumnA1(int.MaxValue);
            name.Is("FXSHRXX");
        }

        [Fact]
        public void A1�`���̗񕶎��񂩂��C���f�b�N�X�ɕϊ�����()
        {
            Assert
                .Throws<ArgumentNullException>(() => MetSsConvert.ToColumnIndex(null))
                .Message
                .Is("�l�� Null �ɂ��邱�Ƃ͂ł��܂���B\r\n�p�����[�^�[��:value");

            Assert
                .Throws<ArgumentNullException>(() => MetSsConvert.ToColumnIndex(""))
                .Message
                .Is("�l�� Null �ɂ��邱�Ƃ͂ł��܂���B\r\n�p�����[�^�[��:value");

            Assert
                .Throws<FormatException>(() => MetSsConvert.ToColumnIndex("@A"))
                .Message
                .Is("Argument format is only alphabet");

            Assert
                .Throws<ArgumentOutOfRangeException>(() => MetSsConvert.ToColumnIndex("FXSHRXY"))
                .Message
                .Is("Argument range max \"FXSHRXX\"\r\n�p�����[�^�[��:value");

            Assert
                .Throws<ArgumentOutOfRangeException>(() => MetSsConvert.ToColumnIndex("AAAAAAAA"))
                .Message
                .Is("Argument range max \"FXSHRXX\"\r\n�p�����[�^�[��:value");

            int index;

            index = MetSsConvert.ToColumnIndex("A");
            index.Is(0);

            index = MetSsConvert.ToColumnIndex("Z");
            index.Is(25);

            index = MetSsConvert.ToColumnIndex("AA");
            index.Is(26);

            index = MetSsConvert.ToColumnIndex("AB");
            index.Is(27);

            index = MetSsConvert.ToColumnIndex("AZ");
            index.Is(51);

            index = MetSsConvert.ToColumnIndex("BA");
            index.Is(52);

            index = MetSsConvert.ToColumnIndex("BZ");
            index.Is(77);

            index = MetSsConvert.ToColumnIndex("CA");
            index.Is(78);

            index = MetSsConvert.ToColumnIndex("AAA");
            index.Is(702);

            index = MetSsConvert.ToColumnIndex("FJZ");
            index.Is(4341);

            index = MetSsConvert.ToColumnIndex("XFD");
            index.Is(16383);

            index = MetSsConvert.ToColumnIndex("FXSHRXX");
            index.Is(int.MaxValue);
        }
    }
}