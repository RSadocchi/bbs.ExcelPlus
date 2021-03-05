using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.IO;

namespace bbs.ExcelPlus
{
    public interface IExcelPlusService
    {
        Stylesheet Stylesheet { get; set; }

        ExcelFileInfo CreateExcelDocument<T>(Dictionary<string, List<T>> sheetsData, string filename);
        ExcelFileInfo CreateExcelDocument<T>(List<T> data, string filename);
        ExcelFileModel ReadExcelDocument(Stream fileStream, bool headerInFirstRow = true);
        ExcelFileModel ReadExcelDocument(string filePath, bool headerInFirstRow = true);
        void SetStylesheet(IEnumerable<NumberingFormat> numberingFormats = null, IEnumerable<Font> fonts = null, IEnumerable<Fill> fills = null, IEnumerable<Border> borders = null, IEnumerable<CellFormat> cellFormats = null);
        byte[] StreamExcelDocument<T>(Dictionary<string, List<T>> sheetsData);
        byte[] StreamExcelDocument<T>(List<T> data);
    }
}