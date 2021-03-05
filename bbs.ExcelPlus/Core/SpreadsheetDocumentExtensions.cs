using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Data;

namespace bbs.ExcelPlus
{
    internal static class SpreadsheetDocumentExtensions
    {
        public static Stylesheet GenerateDefaultStyleSheet(Stylesheet stylesheet = null)
            => GenerateDefaultStyleSheet(
                stylesheet?.NumberingFormats,
                stylesheet?.Fonts,
                stylesheet?.Fills,
                stylesheet?.Borders,
                stylesheet?.CellFormats);

        public static Stylesheet GenerateDefaultStyleSheet(
            NumberingFormats numberingFormats = null,
            Fonts fonts = null,
            Fills fills = null,
            Borders borders = null,
            CellFormats cellFormats = null)
        {
            uint iExcelIndex = 164;

            return new Stylesheet(
                numberingFormats ?? new NumberingFormats(
                    //  
                    new NumberingFormat()                                                   // Custom number format # 164: especially for date-times
                    {
                        NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                        FormatCode = StringValue.FromString("dd/MM/yyyy hh:mm:ss")
                    },
                    new NumberingFormat()                                                   // Custom number format # 165: especially for date times (with a blank time)
                    {
                        NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                        FormatCode = StringValue.FromString("dd/MM/yyyy")
                    }
               ),
                fonts ?? new Fonts(
                    new Font(                                                               // Index 0 - The default font.
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 1 - A 12px bold font
                        new Bold(),
                        new FontSize() { Val = 12 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 2 - An Italic font.
                        new Italic(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" })
                ),
                fills ?? new Fills(
                    new Fill(                                                           // Index 0 - The default fill.
                        new PatternFill() { PatternType = PatternValues.None }),
                    new Fill(                                                           // Index 1 - The default fill (required)
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue("B8B8B8") }
                        )
                        { PatternType = PatternValues.Solid }),
                    new Fill(                                                           // Index 2 - The blue fill.
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue("BDD7EE") }
                        )
                        { PatternType = PatternValues.Solid }),
                    new Fill(                                                           // Index 3 - Gray fill.
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue("BFBFBF") }
                        )
                        { PatternType = PatternValues.Solid })
                ),
                borders ?? new Borders(
                    new Border(                                                         // Index 0 - The default border.
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border(                                                         // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                ),
                cellFormats ?? new CellFormats(
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },                         // Style # 0 - The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                    new CellFormat() { NumberFormatId = 164 },                                         // Style # 1 - DateTimes
                    new CellFormat() { NumberFormatId = 165 },                                         // Style # 2 - Dates (with a blank time)
                    new CellFormat(
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    )
                    { FontId = 1, FillId = 2, BorderId = 0, ApplyFont = true, ApplyAlignment = true }, // Style # 3 - Header row 
                    new CellFormat() { NumberFormatId = 3 },                                           // Style # 4 - Number format: #,##0
                    new CellFormat() { NumberFormatId = 4 },                                           // Style # 5 - Number format: #,##0.00
                    new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },       // Style # 6 - Bold 
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Style # 7 - Italic
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Style # 8 - Times Roman
                    new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Style # 9 - Yellow Fill
                    new CellFormat(                                                                    // Style # 10 - Alignment
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    )
                    { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }      // Style # 11 - Border
                )
            );
        }

        public static void WriteExcelFile(this SpreadsheetDocument spreadsheet, DataSet ds, Stylesheet stylesheet = null)
        {
            spreadsheet.AddWorkbookPart();
            spreadsheet.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

            DefinedNames definedNamesCol = new DefinedNames();

            spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

            WorkbookStylesPart workbookStylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
            workbookStylesPart.Stylesheet = GenerateDefaultStyleSheet(stylesheet);
            workbookStylesPart.Stylesheet.Save();

            uint worksheetNumber = 1;
            Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            foreach (DataTable dt in ds.Tables)
            {
                string worksheetName = dt.TableName;

                WorksheetPart newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart), SheetId = worksheetNumber, Name = worksheetName };

                sheets.Append(sheet);

                EPCore.WriteDataTableToExcelWorksheet(dt, newWorksheetPart, definedNamesCol);

                worksheetNumber++;
            }
            spreadsheet.WorkbookPart.Workbook.Append(definedNamesCol);
            spreadsheet.WorkbookPart.Workbook.Save();
        }
    }
}
