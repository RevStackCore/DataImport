using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace RevStackCore.DataImport
{
    public static class Extensions
    {
        private static string SEPARATOR = ",";
        public static string ToQuotedString(this string source)
        {
            if (string.IsNullOrEmpty(source))
                source = "";
            string QUOTE = "\"";
            source = source.Replace("\"", "\"\"");
            return QUOTE + source + QUOTE;
        }

        public static string ToCsvString(this ExcelWorksheet workSheet, bool ignoreHeader=false, bool matchCase=false)
        {
            StringBuilder sb = new StringBuilder();
            if(!ignoreHeader)
            {
                var columns = workSheet.ToColumnList();
                sb.AppendLine(columns.toCsvString());
            }
            int totalRows = workSheet.Dimension.Rows;
            for (int i = 2; i <= totalRows; i++)
            {
                StringBuilder s = new StringBuilder();
                for (int j = 1; j <= workSheet.Dimension.End.Column; j++)
                {
                    s.Append(workSheet.Cells[i, j].Value.ToNullableString().ToQuotedString() + SEPARATOR);
                }
                sb.AppendLine(s.ToString());
            }

            return sb.ToString();
        }

       
        public static byte[] ToCsvByteArray(this ExcelWorksheet workSheet, bool ignoreHeader = false, bool matchCase=false)
        {
            var cvsString = workSheet.ToCsvString(ignoreHeader,matchCase);
            byte[] byteArray = Encoding.ASCII.GetBytes(cvsString);
            return byteArray;

        }

        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="source"></param>
        /// <param name="startLength"></param>
        /// <returns></returns>
        public static string FirstChars(this string source, int startLength)
        {
            return source.Substring(0, startLength);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string TrimLastCharacter(this String str)
        {
            if (String.IsNullOrEmpty(str))
            {
                return str;
            }
            else
            {
                return str.TrimEnd(str[str.Length - 1]);
            }
        }

        public static string ToNullableString(this object src, bool lowerCase=false)
        {
            if(src==null)
            {
                return "";
            }
            else
            {
                string str= src.ToString();
                if(lowerCase)
                {
                    str = str.ToLower();
                }
                return str;
            }
        }

        private static List<string> ToColumnList(this ExcelWorksheet workSheet, bool matchCase=false)
        {
            List<string> columnNames = new List<string>();
            var w = workSheet;
            for (int i = 1; i <= workSheet.Dimension.End.Column; i++)
            {
                columnNames.Add(workSheet.Cells[1, i].Value.ToNullableString(matchCase)); // 1 = First Row, i = Column Number
            }
            return columnNames;
        }


        private static string toCsvString(this string[] source)
        {
            string separator = ",";
            StringBuilder sb = new StringBuilder();
            foreach (string s in source)
            {
                sb.Append(s.ToQuotedString() + separator);
            }
            string result = sb.ToString();
            return result.TrimLastCharacter();
        }

        private static string toCsvString(this List<string> source)
        { 
            StringBuilder sb = new StringBuilder();
            foreach (string s in source)
            {
                sb.Append(s.ToQuotedString() + SEPARATOR);
            }
            string result = sb.ToString();
            return result.TrimLastCharacter();
        }
    }
}
