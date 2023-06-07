using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExCoreServiceOrder
{
    public class OrderFormat
    {
        [Index(0)]
        public string WorkCenter { get; set; }
        [Index(1)]
        public string Order { get; set; }
        [Index(2)]
        public string Material { get; set; }
        [Index(5)]
        public string OrderType { get; set; }
        [Index(8)]
        public string TargetQty { get; set; }
        [Index(13)]
        public string StartTime { get; set; }
        [Index(14)]
        public string EndTime { get; set; }
        [Index(15)]
        public string SystemStatus { get; set; }
    }
}
