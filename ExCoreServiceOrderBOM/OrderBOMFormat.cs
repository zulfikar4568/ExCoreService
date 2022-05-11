using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExCoreServiceOrderBOM
{
    public class OrderBOMFormat
    {
        [Index(0)]
        public string ProductionOrder { get; set; }
        [Index(2)]
        public string Material { get; set; }
        [Index(7)]
        public string Qty { get; set; }
        [Index(24)]
        public string MaterialGroup { get; set; }
        [Index(25)]
        public string Scanning { get; set; }
    }
}
