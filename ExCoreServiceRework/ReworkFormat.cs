using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExCoreServiceRework
{
    public class ReworkFormat
    {
        [Index(0)]
        public string ReworkPO { get; set; }
        [Index(1)]
        public string Container { get; set; }
    }
}
