using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
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
            ExcelFileModel efm = null;
            using (var fs = new FileStream(filePath, FileMode.Open))
                efm = ReadExcelDocument(fs, headerInFirstRow);
            efm.Name = Path.GetFileName(filePath);
            return efm;
        }

        public ExcelFileModel ReadExcelDocument(Stream fileStream, bool headerInFirstRow = true)
        {
            if (fileStream == null) return null;

            var efm = new ExcelFileModel();
            using (var document = SpreadsheetDocument.Open(stream: fileStream, isEditable: false))
            {
                var wBookPart = document.WorkbookPart;
                var sheets = wBookPart.Workbook.GetFirstChild<Sheets>();

                foreach (Sheet sheet in sheets)
                {
                    Worksheet wSheet = ((WorksheetPart)wBookPart.GetPartById(sheet.Id)).Worksheet;
                    SheetData sheetData = (SheetData)wSheet.GetFirstChild<SheetData>();
                    var sheetModel = new SheetModel()
                    {
                        SheetId = Convert.ToInt32(sheet.Id),
                        SheetName = sheet.Name
                    };

                    int rowIdx = 0;
                    int cellIdx = 1;
                    var columnNames = new List<(int idx, string name)>();
                    var rowModels = new List<RowModel>();

                    foreach (Row row in sheetData)
                    {
                        RowModel rowModel = null;

                        if (headerInFirstRow && rowIdx == 0)
                        {
                            foreach (Cell cell in row)
                            {
                                if (cell.DataType != null)
                                {
                                    if (cell.DataType == CellValues.SharedString)
                                    {
                                        int id;
                                        if (int.TryParse(cell.InnerText, out id))
                                        {
                                            var item = wBookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            columnNames.Add((idx: cellIdx, name: item?.Text?.Text ?? item?.InnerText ?? item?.InnerXml));
                                        }
                                    }
                                }
                                else
                                {
                                    columnNames.Add((idx: cellIdx, name: cell.InnerText ?? cellIdx.ToString()));
                                }
                                cellIdx += 1;
                            }
                        }
                        else
                        {
                            rowModel = new RowModel()
                            {
                                RowId = rowIdx
                            };

                            foreach (Cell cell in row)
                            {
                                if (rowIdx == 0)
                                    columnNames.Add((idx: cellIdx, name: cellIdx.ToString()));

                                var res = new CellModel()
                                {
                                    RowId = rowIdx,
                                    ColumnName = columnNames.FirstOrDefault(t => t.idx == cellIdx).name,
                                    CellId = cellIdx
                                };

                                if (cell.CellFormula?.InnerText == null)
                                {
                                    if (cell.DataType != null)
                                    {
                                        switch (cell.DataType.Value)
                                        {
                                            case CellValues.SharedString:
                                                if (int.TryParse(cell.InnerText, out int id))
                                                {
                                                    var item = wBookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                                    res.CellValue = item?.Text?.Text ?? item?.InnerText ?? item?.InnerXml;
                                                }
                                                break;

                                            case CellValues.Number:
                                                if (decimal.TryParse(cell.InnerText, out decimal numberVal))
                                                {
                                                    res.CellValue = cell.InnerText;
                                                    res.CellValueNumber = numberVal;
                                                }
                                                break;

                                            case CellValues.Date:
                                                if (DateTime.TryParse(cell.InnerText, out DateTime dateVal))
                                                {
                                                    res.CellValue = cell.InnerText;
                                                    res.CellValueDateTime = dateVal;
                                                }
                                                break;

                                            case CellValues.Boolean:
                                                res.CellValue = cell.InnerText;
                                                switch (cell.InnerText)
                                                {
                                                    case "0":
                                                        res.CellValueBoolean = false;
                                                        break;
                                                    default:
                                                        res.CellValueBoolean = true;
                                                        break;
                                                }
                                                break;

                                            default:
                                                res.CellValue = cell.InnerText;
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            EPCore.SetFormattedValue(cell, ref res);
                                        }
                                        catch
                                        {
                                            res.CellValue = cell.InnerText;
                                        }
                                    }
                                }

                                rowModel.Cells.Add(res);
                                cellIdx += 1;
                            }
                        }

                        if (rowModel != null)
                            sheetModel.Rows.Add(rowModel);

                        rowIdx += 1;
                        cellIdx = 1;
                    }

                    efm.Sheets.Add(sheetModel);
                }
            }

            return efm;
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
            var stylesheet = new Stylesheet()
            {
                NumberingFormats = numberingFormats?.Count() > 0 ? new NumberingFormats(numberingFormats) : null,
                Fonts = fonts?.Count() > 0 ? new Fonts(fonts) : null,
                Fills = fills?.Count() > 0 ? new Fills(fills) : null,
                Borders = borders?.Count() > 0 ? new Borders(borders) : null,
                CellFormats = cellFormats?.Count() > 0 ? new CellFormats(cellFormats) : null
            };

            Stylesheet = stylesheet;
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
