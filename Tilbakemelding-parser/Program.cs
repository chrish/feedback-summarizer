using System;

namespace Tilbakemelding_parser
{
    /// <summary>
    /// What to do... 
    /// 
    /// 1 - åpne regnearket. 
    /// 2 - Gå til kol cc, legg inn følgende headere: Antall grønne, antall gule, antall røde, totalvurdering, kommentar. 
    /// 3 - Åpne regelsettet, vurder kolonner ihht til regler og verdi. Reglene blir regexp eller noe i den gata. Sett bgcolor til utfallet, hold track på utfall i en eller annen struktur. 
    /// 4 - Print oppsummering av vurderingene
    /// 5 - Gjør opp en totalvurdering basert på oppsummeringen, med prio 1-5. 
    /// 6 - Lagre / lukk
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
//            Test wm = new Test(@"X:\Code\internt\kartlegging-vurdering\Sjekkliste risiko analyse_demo.xlsx");

            WorkbookManager wbm = new WorkbookManager(@"X:\Code\internt\kartlegging-vurdering\Sjekkliste risiko analyse_demo.xlsx");

            Parser p = new Parser("rules.json", wbm);

            int numRows = wbm.GetNumberOfRows();

            for (int i = 2; i < numRows; i++)
            {
                var row = wbm.GetRow(i);
                p.ParseRow(row);
            }
            wbm.Save();
        }
    }
}
