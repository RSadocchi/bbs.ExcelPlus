using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace bbs.ExcelPlus
{
    public class ExcelPlusService
    {

        #region ///READ
        public ExcelFileModel ReadExcelDocument(string filePath, bool headerInFirstRow = true)
        {
            return null;
        }

        public ExcelFileModel ReadExcelDocument(Stream fileStream, bool headerInFirstRow = true)
        {
            return null;
        }
        #endregion

        #region ///WRITE
        public Stylesheet Stylesheet { get; set; } = null;
        public void SetStylesheet(
            IEnumerable<NumberingFormat> numberingFormats = null,
            IEnumerable<Font> fonts = null,
            IEnumerable<Fill> fills = null,
            IEnumerable<Border> borders = null,
            IEnumerable<CellFormat> cellFormats = null)
        {

        }

        public byte[] StreamExcelDocument<T>(List<T> data)
        {
            using (var stream = new MemoryStream())
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
            {
                var ds = new DataSet();
                ds.Tables.Add(EPCore.ListToDataTable(data));
                document.WriteExcelFile(ds, Stylesheet);
                return stream.ToArray();
            }
        }

        public byte[] StreamExcelDocument<T>(Dictionary<string, List<T>> sheetsData)
        {
            using (var stream = new MemoryStream())
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
            {
                var ds = new DataSet();
                foreach (var data in sheetsData)
                    ds.Tables.Add(EPCore.ListToDataTable(data.Value, data.Key));
                document.WriteExcelFile(ds, Stylesheet);
                return stream.ToArray();
            }
        }

        public ExcelFileInfo CreateExcelDocument<T>(List<T> data, string filename)
        {
            var efi = new ExcelFileInfo()
            {
                FileName = Path.GetFileName(filename),
                CompletePath = filename,
                Bytes = StreamExcelDocument(data)
            };

            File.WriteAllBytesAsync(filename, efi.Bytes);
            efi.FileInfo = new FileInfo(filename);

            return efi;
        }

        public ExcelFileInfo CreateExcelDocument<T>(Dictionary<string, List<T>> sheetsData, string filename)
        {
            var efi = new ExcelFileInfo()
            {
                FileName = Path.GetFileName(filename),
                CompletePath = filename,
                Bytes = StreamExcelDocument(sheetsData)
            };
            
            File.WriteAllBytesAsync(filename, efi.Bytes);
            efi.FileInfo = new FileInfo(filename);

            return efi;
        }
        #endregion
    }
}
