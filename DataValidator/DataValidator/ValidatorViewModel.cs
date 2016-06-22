using Microsoft.Practices.Prism.Commands;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace DataValidator
{
    public class ValidatorViewModel
    {
        public ObservableCollection<ComparedAtms> DifferentAtms
        {
            get;
            private set;
        }

        public ObservableCollection<Atm> NotExistAtms
        {
            get;
            private set;
        }

        public ValidatorViewModel()
        {
            this.DifferentAtms = new ObservableCollection<ComparedAtms>();
            this.NotExistAtms = new ObservableCollection<Atm>();
        }

        #region ImportEpmReportCommand

        private DelegateCommand importEpmReportCommand;

        public ICommand ImportEpmReportCommand
        {
            get
            {
                if (importEpmReportCommand == null)
                    importEpmReportCommand = new DelegateCommand(ImportEpmReportExecuted, ImportEpmReportCanExecute);
                return importEpmReportCommand;
            }
        }

        public bool ImportEpmReportCanExecute()
        {
            return true;
        }

        public void ImportEpmReportExecuted()
        {
            var epmAtms = ExcelManager.ImprortAtmDataFromEpmReport(@"C:\EPM_Report_2016-06-04.xlsx");
            var siteAtms = ExcelManager.ImprortAtmDataFromSharepoint(@"C:\Atms.xlsx");

            
            bool exist = false;

            //compare
            foreach (var epm in epmAtms)
            {
                foreach (var site in siteAtms)
                {
                    if (String.Equals(epm.Name, site.Name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        exist = true; 
                        if (!String.Equals(epm.AptraCD2Version, site.AptraCD2Version, StringComparison.CurrentCultureIgnoreCase))
                        {
                            this.DifferentAtms.Add(new ComparedAtms() { Name = epm.Name, Customer = epm.Customer, EpmMUP = epm.AptraCD2Version, MUP = site.AptraCD2Version});
                        }
                        break;
                    }
                }

                if (exist == false)
                {
                    NotExistAtms.Add(epm);
                }
                else
                {
                    exist = false;
                }
                
            }

            var list = this.DifferentAtms.ToList<ComparedAtms>();
            var dataTable = ExcelManager.ConvertListToDataTable(list);
            ExcelManager.ExportDataToExcel("DifferentMUP", dataTable, new System.IO.DirectoryInfo(@"C:\"));
        }

        #endregion

        #region ExportReportCommand

        private DelegateCommand exportReportCommand;

        public ICommand ExportReportCommand
        {
            get
            {
                if (exportReportCommand == null)
                    exportReportCommand = new DelegateCommand(ExportReportExecuted, ExportReportCanExecute);
                return exportReportCommand;
            }
        }

        public bool ExportReportCanExecute()
        {
            return true;
        }

        public void ExportReportExecuted()
        {
            var list1 = this.DifferentAtms.ToList<ComparedAtms>();
            var dataTable1 = ExcelManager.ConvertListToDataTable(list1);
            ExcelManager.ExportDataToExcel("DifferentMUP", dataTable1, new System.IO.DirectoryInfo(@"C:\"));

            MessageBox.Show(@"DifferentMUP.xlsx exported to c:\");

            var list2 = this.NotExistAtms.ToList<Atm>();
            var dataTable2 = ExcelManager.ConvertListToDataTable(list2);
            ExcelManager.ExportDataToExcel("NotExistAtms", dataTable2, new System.IO.DirectoryInfo(@"C:\"));

            MessageBox.Show(@"NotExistAtms.xlsx exported to c:\");
        }

        #endregion
    }
}
