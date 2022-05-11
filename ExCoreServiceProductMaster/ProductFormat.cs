using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExCoreServiceProductMaster
{
    public class ProductFormat
    {  
        [Index(0)]
        public string Product { get; set; }
        [Index(1)]
        public string Description { get; set; }
        [Index(2)]
        public string Procurement { get; set; }
        [Index(3)]
        public string UOM { get; set; }
        [Index(4)]
        public string ProductType { get; set; }
        [Index(5)]
        public string ProductFamily { get; set; }
        [Index(6)]
        public string ProductFamilyDescription { get; set; }
    }
}
