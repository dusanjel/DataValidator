using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataValidator
{
    public class EpmInventory
    {
        public string Customer { get; set; }
        public string Device { get; set; }
        public string DetectedSoftware { get; set; }
        public string Version { get; set; }
        public string DetectionDate { get; set; }
        public string DetectionTime { get; set; }
    }
}
