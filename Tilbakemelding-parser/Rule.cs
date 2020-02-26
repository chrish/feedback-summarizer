using System;
using System.Collections.Generic;
using System.Text;

namespace Tilbakemelding_parser
{
    public class Rule
    {
        public string Column { get; set; }
        public List<Dictionary<string, string>> Rules { get; set; }

        public string RefColumn { get; set; }

        public Rule()
        {
            Column = "";
            Rules = new List<Dictionary<string, string>>();
        }
    }
}
