using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Tilbakemelding_parser
{
    public class Parser
    {
        List<Rule> Rules;
        WorkbookManager Wbm;

        public Parser(string rulesFile, WorkbookManager wbm) {
            Rules = JsonConvert.DeserializeObject<List<Rule>>(File.ReadAllText(rulesFile));
            Wbm = wbm;
        }

        /// <summary>
        /// Parse en rad ved å gå over alt av regler.
        /// 
        /// Reglene er definer som json:
        ///     Array med alle regler
        ///         Hver idx er et objekt med en gitt column=bokstav
        ///         Objektet inneholder også en array "rules" som inneholder reglene som skal parses
        ///             Hver regel er et objekt i arrayen med color:farge, regexp:regel som skal parses og en optional dependsOn:kolonne-Farge
        ///             Tanken er at vi kan ha en dependency, der en kolonne får en farge basert på eget innhold _samt_ fargen fra en annen kolonne. 
        ///             Eks: Delvis kontroll på lagret informasjon er gul (kol AL), dersom vi ikke vet om 
        ///             
        /// 
        /// </summary>
        /// <param name="rowToParse"></param>
        /// <returns></returns>
        public Row ParseRow(Row rowToParse) {
            Dictionary<string, string> flags = new Dictionary<string, string>();

            foreach (Rule r in Rules) {
                // do stuff, add flags to flags as we go along....
                var rxCol = new Regex("^(" + r.Column + ")[0-9]", RegexOptions.IgnoreCase);

                // Se på regelen og hent cella vi vil jobbe mot. 
                Cell cell = (from Cell c in rowToParse.Elements<Cell>() where rxCol.IsMatch(c.CellReference) select c).Single<Cell>();

                for (int i = 0; i < r.Rules.Count; i++) {
                    // sjekk om vi har en regel som matcher. Første vinner. Dersom mer enn en treffer flagger vi cella med grått. 
                    Regex rxCellVal = new Regex(r.Rules[i]["Regexp"], RegexOptions.IgnoreCase);
                    Regex refRxCellVal;
                    var cellValue = GetCellValue(cell);
                    var refCellValue = "";
                    
                    Cell refCell;

                    // Resultatet av innholdet i en celle kan vektes mot en annen:
                    if (r.RefColumn != null) {
                        int rowIdx = Convert.ToInt32((uint)rowToParse.RowIndex);
                        string refCellRef = r.RefColumn + rowIdx;

                        refCell = (from Cell c in rowToParse.Elements<Cell>() where c.CellReference.Equals(refCellRef) select c).Single<Cell>();
                        refCellValue = GetCellValue(refCell);
                    }

                    // Sjekk om regelen har en refregexp:
                    if (r.Rules[i].ContainsKey("RefRegexp"))
                    {
                        refRxCellVal = new Regex(r.Rules[i]["RefRegexp"], RegexOptions.IgnoreCase);
                        // Sjekk om vi har treff mot refcella
                        if (rxCellVal.IsMatch(cellValue) && refRxCellVal.IsMatch(refCellValue))
                        {
                            flags.Add(cell.CellReference, r.Rules[i]["Color"]);
                            // Dersom treff går vi ut av løkka, slik at vi kan ha catch-all regler.
                            break;
                        }
                    }
                    else if (rxCellVal.IsMatch(cellValue))
                    {
                        // Ingen refcelle, kjør vanlig compare
                        flags.Add(cell.CellReference, r.Rules[i]["Color"]);
                        break;
                    }
                }
            }

            // At the end, set bgcolors according to flags, summarize and add columns as needed.
            foreach (var kv in flags) {
                Cell c = (from Cell cc in rowToParse where cc.CellReference.Equals(kv.Key) select cc).Single<Cell>();
                c.StyleIndex = Convert.ToUInt32(WorkbookManager.MapStyleIdx(kv.Value));
                
            }

            // Legg til kolonner her.
            int numGr = (from KeyValuePair<string, string> f in flags where f.Value.Equals("green") select f).Count();
            int numYl = (from KeyValuePair<string, string> f in flags where f.Value.Equals("yellow") select f).Count();
            int numRd = (from KeyValuePair<string, string> f in flags where f.Value.Equals("red") select f).Count();

            // CC, CD, CE, CF for totalt antall grønn, gul og rød, samt en total score
            Cell grCell = new Cell() { CellReference = "CC" + rowToParse.RowIndex, DataType = CellValues.String, CellValue = new CellValue(numGr.ToString()) };
            Cell ylCell = new Cell() { CellReference = "CD" + rowToParse.RowIndex, DataType = CellValues.String, CellValue = new CellValue(numYl.ToString()) };
            Cell rdCell = new Cell() { CellReference = "CE" + rowToParse.RowIndex, DataType = CellValues.String, CellValue = new CellValue(numRd.ToString()) };

            grCell.StyleIndex = Convert.ToUInt32(WorkbookManager.MapStyleIdx("green"));
            ylCell.StyleIndex = Convert.ToUInt32(WorkbookManager.MapStyleIdx("yellow"));
            rdCell.StyleIndex = Convert.ToUInt32(WorkbookManager.MapStyleIdx("red"));

            rowToParse.Append(grCell, ylCell, rdCell);

            // Så må vi tenke litt... Hva er fornuftige grenser for endelig score? 
            // Flere grønne enn gule+røde => grønn
            // Flere grønne+gule enn røde => gul, eller røde+gule enn grønne => gul
            // Flere røde enn grønne+gule => rød
            var finalScore = "red";
            if (numGr > numYl + numRd) finalScore = "green";
            else if (numGr + numYl >= numRd || numGr <= numYl + numRd) finalScore = "yellow";
            else if (numGr + numYl <= numRd) finalScore = "red";

            Cell fnlCell = new Cell() { CellReference = "CF" + rowToParse.RowIndex, DataType = CellValues.String, CellValue = new CellValue("") };
            fnlCell.StyleIndex = Convert.ToUInt32(WorkbookManager.MapStyleIdx(finalScore));
            rowToParse.Append(fnlCell);

            return rowToParse;
        }

        /// <summary>
        /// Hjelpemetode for å hente en celleverdi
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        protected string GetCellValue(Cell cell) {
            var cellValue = "";
            
            if (cell.DataType != null && cell.DataType == "s")
            {
                var dt = cell.DataType;
                cellValue = Wbm.GetStringFromSharedStringTable(Convert.ToInt32(cell.CellValue.Text));
            }
            else if (cell.CellValue != null)
            {
                cellValue = cell.CellValue.Text;
            }
            else
            {
                cellValue = string.Empty;
            }

            return cellValue;
        }
    }
}
