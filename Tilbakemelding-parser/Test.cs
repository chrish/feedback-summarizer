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
    public class Test
    {
        SpreadsheetDocument document;

        public Test(string wkbookFile)
        {
            try
            {
                document = SpreadsheetDocument.Open(wkbookFile, false);
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheets formresults = workbookPart.Workbook.GetFirstChild<Sheets>(); // Antar at det alltid er det første arket vi skal prosessere


                //using for each loop to get the sheet from the sheetcollection  
                foreach (Sheet thesheet in formresults)
                {
                    Console.WriteLine("Excel Sheet Name : " + thesheet.Name);
                    Console.WriteLine("----------------------------------------------- ");
                    //statement to get the worksheet object by using the sheet id  
                    Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                    var wbpst = workbookPart.GetPartsOfType<WorkbookStylesPart>();

                    SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                    foreach (Row thecurrentrow in thesheetdata)
                    {
                        foreach (Cell thecurrentcell in thecurrentrow)
                        {
                            //statement to take the integer value  
                            string currentcellvalue = string.Empty;
                            if (thecurrentcell.DataType != null)
                            {
                                if (thecurrentcell.DataType == CellValues.SharedString)
                                {
                                    int id;
                                    if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                    {
                                        SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                        if (item.Text != null)
                                        {
                                            //code to take the string value  
                                            Console.WriteLine(item.Text.Text + " ");
                                        }
                                        else if (item.InnerText != null)
                                        {
                                            currentcellvalue = item.InnerText;
                                        }
                                        else if (item.InnerXml != null)
                                        {
                                            currentcellvalue = item.InnerXml;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine(Convert.ToInt16(thecurrentcell.InnerText) + " ");
                            }
                        }
                        Console.WriteLine();
                    }
                    Console.WriteLine("");

                    Console.ReadLine();
                }
            }
            catch (Exception)
            {

            }
        }
    }
}
