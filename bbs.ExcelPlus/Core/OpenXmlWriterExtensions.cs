using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;

namespace bbs.ExcelPlus
{
    internal static class OpenXmlWriterExtensions
    {
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

        public static void AppendTextCell(this OpenXmlWriter writer, string cellReference, string cellStringValue)
        {
            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                DataType = CellValues.String
            });
        }

        public static void AppendHeaderTextCell(this OpenXmlWriter writer, string cellReference, string cellStringValue)
        {
            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                DataType = CellValues.String,
                StyleIndex = 3
            });
        }
    }
}
