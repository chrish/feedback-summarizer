using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Tilbakemelding_parser
{
    /// <summary>
    /// Åpne regnearket
    /// Les ut verdi i angitt rad/kolonne
    /// Skriv verdi til angitt rad/kolonne
    /// Sett bgfarge på angitt rad/kolonne
    /// Lagre regnearket
    /// </summary>
    public class WorkbookManager
    {
        SpreadsheetDocument Document;
        string OriginalFile;
        Worksheet FormResultsWkSheet;

        public static int REDSTYLEIDX;
        public static int YELLOWSTYLEIDX;
        public static int GREENSTYLEIDX;

        public static int MapStyleIdx(string color) {
            if (color.Equals("red", StringComparison.OrdinalIgnoreCase)) {
                return REDSTYLEIDX;
            }
            else if (color.Equals("Yellow", StringComparison.OrdinalIgnoreCase))
            {
                return YELLOWSTYLEIDX;
            }
            else if (color.Equals("Green", StringComparison.OrdinalIgnoreCase))
            {
                return GREENSTYLEIDX;
            }
            else
            {
                return 2; //Default grey idx
            }
        }

        public WorkbookManager(string wkbookFile) {
            OriginalFile = wkbookFile;
            Document = SpreadsheetDocument.Open(wkbookFile, true);
            WorkbookPart workbookPart = Document.WorkbookPart;
            var sheetId = (workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>()).FirstOrDefault().Id; // Antar at det alltid er det første arket vi skal prosessere              

            FormResultsWkSheet = ((WorksheetPart)workbookPart.GetPartById(sheetId)).Worksheet;

            // Sett opp styles. Hent ut exi styleparts, sjekk om vi har styles allerede/legg til og reg idx
            WorkbookStylesPart wbpst = workbookPart.GetPartsOfType<WorkbookStylesPart>().Single();

            int fills;

            if (wbpst.Stylesheet.Fills.Count <= 2)
            { // Det finnes bare to defaultstyles, har vi mer har vi lagt til allerede. 
                wbpst.Stylesheet.Fills.Append(
                    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FF00FF00" } }) { PatternType = PatternValues.Solid }),  //4; grønn
                    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }) { PatternType = PatternValues.Solid }), //3; gul
                    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFF0000" } }) { PatternType = PatternValues.Solid }) //2; rød
                );
               fills = wbpst.Stylesheet.Fills.Count();

                CellFormat cfGreen = new CellFormat() { FillId = Convert.ToUInt32(fills - 3) };
                CellFormat cfYellow = new CellFormat() { FillId = Convert.ToUInt32(fills - 2) };
                CellFormat cfRed = new CellFormat() { FillId = Convert.ToUInt32(fills - 1) };

                wbpst.Stylesheet.CellFormats.Append(cfGreen, cfYellow, cfRed);

            }

            wbpst.Stylesheet.Save();
            
            fills = Convert.ToInt32((UInt32)wbpst.Stylesheet.CellFormats.ChildElements.Count);

            REDSTYLEIDX = fills - 1;
            YELLOWSTYLEIDX = fills - 2;
            GREENSTYLEIDX = fills - 3;

        }

        public string GetStringFromSharedStringTable(int stringId) {
            SharedStringTable stringTable= Document.WorkbookPart.SharedStringTablePart.SharedStringTable;
            string sharedString = Document.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements[stringId].InnerText;
            return sharedString;
        }

        public int GetNumberOfRows() {
            SheetData sheetdata = (SheetData)FormResultsWkSheet.GetFirstChild<SheetData>();

            int rows = (from Row r in sheetdata select r).Count();

            return rows;
        }

        public Row GetRow(int rowNumber) {

            SheetData sheetdata = (SheetData)FormResultsWkSheet.GetFirstChild<SheetData>();

            var row = (from Row r in sheetdata where r.RowIndex == rowNumber select r).Single();

            return row;
        }

        public void UpdateEntireRow(int rowNumber, Row row) {

        }

        public void Save() {
            var dir = Path.GetDirectoryName(OriginalFile);
            Document.SaveAs(dir + "\\tempfile.xlsx");
        }

        public void CloseSpreadSheet() {
            Document.Close();
        }
    }
}
