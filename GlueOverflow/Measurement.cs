using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GlueOverflow
{
    class Measurement
    {
        public int Number { get; set; }
        public string Program { get; set; }
        public string Build
        { get; set; }
        public string CHS_Result
        { get; set; }
        public string Vendor { get; set; }
        public string Location { get; set; }
        public string UL { get; set; }
        public string Dimension { get; set; }
        public double Raw_value { get; set; }
    }
}
