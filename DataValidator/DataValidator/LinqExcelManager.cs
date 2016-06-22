using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataValidator
{
    public static class LinqExcelManager
    {
        public static IEnumerable<EpmInventory> ImportEpmReport(string filePath)
        {
            var excel = new ExcelQueryFactory(filePath);
            excel.AddMapping<EpmInventory>(x => x.DetectedSoftware, "Detected Software"); //maps the "DetectedSoftware" property to the "Detected Software" column
            excel.AddMapping("DetectionDate", "Detection Date");       //maps the "Employees" property to the "Employee Count" column
            excel.AddMapping("DetectionTime", "Detection Time");

            var epmAtms = from c in excel.Worksheet<EpmInventory>("Inventory Report") //worksheet name = 'US Companies'
                               where c.DetectedSoftware == "StandardBase-CD2-MUP"
                               select c;

            return epmAtms;
        }
    }
}
