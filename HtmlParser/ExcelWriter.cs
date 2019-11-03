using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlParser
{
    public class ExcelWriter
    {
        private Stream stream;
        BinaryWriter writer;
        readonly ushort[] clBegin = { 0x0809, 8, 0, 0x10, 0, 0 };
        readonly ushort[] clEnd = { 0x0A, 00 };


        public ExcelWriter(Stream stream)
        {
            this.stream = stream;
            writer = new BinaryWriter(stream);
        }
        private void WriteUshortArray(ushort[] value)
        {
            for (int i = 0; i < value.Length; i++)
                writer.Write(value[i]);
        }
        public void WriteCell(int row, int col, string value)
        {
            ushort[] clData = { 0x0204, 0, 0, 0, 0, 0 };
            int iLen = value.Length;
            byte[] plainText = Encoding.Default.GetBytes(value);
            clData[1] = (ushort)(8 + iLen);
            clData[2] = (ushort)row;
            clData[3] = (ushort)col;
            clData[5] = (ushort)iLen;
            WriteUshortArray(clData);
            writer.Write(plainText);
        }
        public void WriteCell(int row, int col)
        {
            ushort[] clData = { 0x0201, 6, 0, 0, 0x17 };
            clData[2] = (ushort)row;
            clData[3] = (ushort)col;
            WriteUshortArray(clData);
        }
        public void BeginWrite()
        {
            WriteUshortArray(clBegin);
        }
        public void EndWrite()
        {
            WriteUshortArray(clEnd);
            writer.Flush();
        }
    }
}
