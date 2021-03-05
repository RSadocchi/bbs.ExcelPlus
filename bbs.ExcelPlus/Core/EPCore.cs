using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace bbs.ExcelPlus
{

    internal static class EPCore
    {
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

        public static void AppendNumericCell(this OpenXmlWriter writer, string cellReference, string cellStringValue, NumericFormatDecimaPlaces numericFormatDecimaPlaces = NumericFormatDecimaPlaces.Zero)
        {
            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                StyleIndex = (UInt32)numericFormatDecimaPlaces,
                DataType = CellValues.Number
            });
        }

        public static void AppendFormulaCell(this OpenXmlWriter writer, string cellReference, string cellStringValue)
        {
            writer.WriteElement(new Cell
            {
                CellFormula = new CellFormula(cellStringValue),
                CellReference = cellReference,
                DataType = CellValues.Number
            });
        }

        public static void AppendDateCell(this OpenXmlWriter writer, string cellReference, DateTime dateTimeValue)
        {
            string cellStringValue = dateTimeValue.ToOADate().ToString(CultureInfo.InvariantCulture);
            bool bHasBlankTime = (dateTimeValue.Date == dateTimeValue);

            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                StyleIndex = UInt32Value.FromUInt32(bHasBlankTime ? (uint)2 : (uint)1),
                DataType = CellValues.Number
            });
        }
    }
}
