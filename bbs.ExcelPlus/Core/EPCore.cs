using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace bbs.ExcelPlus
{
    internal static class EPCore
    {
        const int DEFAULT_COLUMN_WIDTH = 20;

        public static string GetExcelColumnName(int columnIndex)
        {
            int firstInt = columnIndex / 676;
            int secondInt = (columnIndex % 676) / 26;
            if (secondInt == 0)
            {
                secondInt = 26;
                firstInt--;
            }
            int thirdInt = (columnIndex % 26);

            char firstChar = (char)('A' + firstInt - 1);
            char secondChar = (char)('A' + secondInt - 1);
            char thirdChar = (char)('A' + thirdInt);

            if (columnIndex < 26)
                return thirdChar.ToString();

            if (columnIndex < 702)
                return string.Format("{0}{1}", secondChar, thirdChar);

            return string.Format("{0}{1}{2}", firstChar, secondChar, thirdChar);
        }

        public static string ReplaceHexadecimalSymbols(string txt)
        {
            string pattern = "[\x00-\x08\x0B\x0C\x0E-\x1F]";
            return Regex.Replace(txt, pattern, string.Empty, RegexOptions.Compiled);
        }

        public static Type GetNullableType(Type t)
        {
            Type returnType = t;
            if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
                returnType = Nullable.GetUnderlyingType(t);
            return returnType;
        }

        public static bool IsNullableType(Type type)
            => (type == typeof(string) || type.IsArray || (type.IsGenericType && type.GetGenericTypeDefinition().Equals(typeof(Nullable<>))));

        public static DataTable ListToDataTable<T>(IEnumerable<T> list, string tableName = null)
        {
            DataTable dt = new DataTable(tableName);

            foreach (PropertyInfo info in typeof(T).GetProperties())
                dt.Columns.Add(new DataColumn(info.Name, GetNullableType(info.PropertyType)));

            foreach (T t in list)
            {
                DataRow row = dt.NewRow();
                foreach (PropertyInfo info in typeof(T).GetProperties())
                {
                    if (!IsNullableType(info.PropertyType))
                        row[info.Name] = info.GetValue(t, null);
                    else
                        row[info.Name] = (info.GetValue(t, null) ?? DBNull.Value);
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        public static void WriteDataTableToExcelWorksheet(DataTable dt, WorksheetPart worksheetPart, DefinedNames definedNamesCol)
        {
            OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart, Encoding.ASCII);
            writer.WriteStartElement(new Worksheet());

            UInt32 inx = 1;
            writer.WriteStartElement(new Columns());
            foreach (DataColumn dc in dt.Columns)
            {
                writer.WriteElement(new Column { Min = inx, Max = inx, CustomWidth = true, Width = DEFAULT_COLUMN_WIDTH });
                inx++;
            }
            writer.WriteEndElement();


            writer.WriteStartElement(new SheetData());

            string cellValue = "";
            string cellReference = "";

            int numberOfColumns = dt.Columns.Count;
            bool[] IsIntegerColumn = new bool[numberOfColumns];
            bool[] IsFloatColumn = new bool[numberOfColumns];
            bool[] IsDateColumn = new bool[numberOfColumns];

            string[] excelColumnNames = new string[numberOfColumns];
            for (int n = 0; n < numberOfColumns; n++)
                excelColumnNames[n] = GetExcelColumnName(n);

            uint rowIndex = 1;

            writer.WriteStartElement(new Row { RowIndex = rowIndex, Height = 20, CustomHeight = true });
            for (int colInx = 0; colInx < numberOfColumns; colInx++)
            {
                DataColumn col = dt.Columns[colInx];
                writer.AppendHeaderTextCell(excelColumnNames[colInx] + "1", col.ColumnName);
                IsIntegerColumn[colInx] = (col.DataType.FullName.StartsWith("System.Int"));
                IsFloatColumn[colInx] = (col.DataType.FullName == typeof(decimal).FullName) || (col.DataType.FullName == typeof(double).FullName) || (col.DataType.FullName == typeof(Single).FullName);
                IsDateColumn[colInx] = (col.DataType.FullName == typeof(DateTime).FullName);
            }
            writer.WriteEndElement();   //  End of header "Row"

            double cellFloatValue = 0;
            CultureInfo ci = Thread.CurrentThread.CurrentCulture; //new CultureInfo("en-US");
            foreach (DataRow dr in dt.Rows)
            {
                ++rowIndex;

                writer.WriteStartElement(new Row { RowIndex = rowIndex });

                for (int colInx = 0; colInx < numberOfColumns; colInx++)
                {
                    cellValue = dr.ItemArray[colInx].ToString();
                    cellValue = ReplaceHexadecimalSymbols(cellValue);
                    cellReference = excelColumnNames[colInx] + rowIndex.ToString();

                    if (IsIntegerColumn[colInx] || IsFloatColumn[colInx])
                    {
                        cellFloatValue = 0;
                        bool bIncludeDecimalPlaces = IsFloatColumn[colInx];
                        if (double.TryParse(cellValue, out cellFloatValue))
                        {
                            cellValue = cellFloatValue.ToString(ci);
                            writer.AppendNumericCell(cellReference, cellValue, bIncludeDecimalPlaces ? NumericFormatDecimaPlaces.Two : NumericFormatDecimaPlaces.Zero);
                        }
                    }
                    else if (IsDateColumn[colInx])
                    {
                        DateTime dateValue;
                        if (DateTime.TryParse(cellValue, out dateValue))
                            writer.AppendDateCell(cellReference, dateValue);
                        else
                            writer.AppendTextCell(cellReference, cellValue);
                    }
                    else
                        writer.AppendTextCell(cellReference, cellValue);
                }
                writer.WriteEndElement(); //  End of Row
            }
            writer.WriteEndElement(); //  End of SheetData
            writer.WriteEndElement(); //  End of worksheet

            writer.Close();
        }

        public static bool CreateExcelDocument(DataSet ds, string excelFilename)
        {
            try
            {
                MemoryStream memoryStream = new MemoryStream();
                SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(excelFilename, SpreadsheetDocumentType.Workbook))
                    spreadsheet.WriteExcelFile(ds);
                
                Trace.WriteLine("Successfully created: " + excelFilename);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Failed, exception thrown: " + ex.Message);
                return false;
            }
        }
    }
}
