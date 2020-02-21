using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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

                for (int i=0; i<r.Rules.Count; i++) {
                    // sjekk om vi har en regel som matcher. Første vinner. Dersom mer enn en treffer flagger vi cella med grått. 
                    var rxCellVal = new Regex(r.Rules[i]["Regexp"], RegexOptions.IgnoreCase);

                    //Todo: Cell Datatype null for kol K (tall)
                    //           Datatype s for kol O, denne linker til SharedStringTable. 
                    // Må lage en if som ser på datatype, dersom null brukes innhold direkte, 
                    // dersom s må WorkBookManager inneholde en metode som går mot stringtablen 
                    // og henter stringen for et gitt innhold. 

                    // Dersom treff går vi ut av løkka, slik at vi kan ha catch-all regler.
                    var cellValue = "";

                    if (cell.DataType != null && cell.DataType == "s") 
                    {
                        var dt = cell.DataType;
                        cellValue = Wbm.GetStringFromSharedStringTable(Convert.ToInt32(cell.CellValue.Text));
                    }
                    else {
                        cellValue = cell.CellValue.Text;
                    }
                    
                    if (rxCellVal.IsMatch(cellValue)) {
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

            return rowToParse;
        }
    }
}
