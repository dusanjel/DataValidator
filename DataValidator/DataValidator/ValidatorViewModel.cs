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
        private String SWDTool;

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

        #region ImportSCCMReportCommand

        private DelegateCommand importSCCMReportCommand;

        public ICommand ImportSCCMReportCommand
        {
            get
            {
                if (importSCCMReportCommand == null)
                    importSCCMReportCommand = new DelegateCommand(ImportSCCMReportExecuted, ImportSCCMReportCanExecute);
                return importSCCMReportCommand;
            }
        }

        public bool ImportSCCMReportCanExecute()
        {
            return true;
        }

        public void ImportSCCMReportExecuted()
        {
            try
            {
                this.DifferentAtms.Clear(); // Dusan obrisi kolekciju
                this.NotExistAtms.Clear(); // Clear to prevent duplicate records.
                var SCCMAtms = ExcelManager.ImprortAtmDataFromSCCMReport(@"C:\SCCM.xlsx");
                var siteAtms = ExcelManager.ImprortAtmDataFromSharepoint(@"C:\Atms.xlsx");
                SWDTool = "DifferentMUPSCCM";

                bool exist = false;

                //compare
                foreach (var SCCM in SCCMAtms)
                {
                    foreach (var site in siteAtms)
                    {
                        if (String.Equals(SCCM.Name, site.Name, StringComparison.CurrentCultureIgnoreCase))
                        {
                            exist = true;
                            if (!String.Equals(SCCM.AptraCD2Version, site.AptraCD2Version, StringComparison.CurrentCultureIgnoreCase))
                            {
                                this.DifferentAtms.Add(new ComparedAtms() { Name = SCCM.Name, Customer = SCCM.Customer, SWDMUP = SCCM.AptraCD2Version, MUP = site.AptraCD2Version });
                            }
                            break;
                        }
                    }

                    if (exist == false)
                    {
                        NotExistAtms.Add(SCCM);
                    }
                    else
                    {
                        exist = false;
                    }

                }

                var list = this.DifferentAtms.ToList<ComparedAtms>();
                var dataTable = ExcelManager.ConvertListToDataTable(list);
                ExcelManager.ExportDataToExcel(SWDTool, dataTable, new System.IO.DirectoryInfo(@"C:\"));
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        #endregion

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
            try
            {
                this.DifferentAtms.Clear(); // Dusan obrisi kolekciju
                this.NotExistAtms.Clear(); // Clear to prevent duplicate records.
                var epmAtms = ExcelManager.ImprortAtmDataFromEpmReport(@"C:\EPM.xlsx");
                var siteAtms = ExcelManager.ImprortAtmDataFromSharepoint(@"C:\Atms.xlsx");
                SWDTool = "DifferentMUPEpm";

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
                                this.DifferentAtms.Add(new ComparedAtms() { Name = epm.Name, Customer = epm.Customer, SWDMUP = epm.AptraCD2Version, MUP = site.AptraCD2Version });
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
                ExcelManager.ExportDataToExcel(SWDTool, dataTable, new System.IO.DirectoryInfo(@"C:\"));
            }
            catch (Exception ex)
            {

                throw ex;
            }
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
            try
            {
                var list1 = this.DifferentAtms.ToList<ComparedAtms>();
                var dataTable1 = ExcelManager.ConvertListToDataTable(list1);
                ExcelManager.ExportDataToExcel(SWDTool, dataTable1, new System.IO.DirectoryInfo(@"C:\"));

                MessageBox.Show(SWDTool + @".xlsx exported to c:\");

                var list2 = this.NotExistAtms.ToList<Atm>();
                var dataTable2 = ExcelManager.ConvertListToDataTable(list2);


                if (SWDTool == "DifferentMUPEpm")
                {
                    ExcelManager.ExportDataToExcel("NotExistAtmsEpm", dataTable2, new System.IO.DirectoryInfo(@"C:\"));
                    MessageBox.Show(@"NotExistAtmsEpm.xlsx exported to c:\");
                }
                else if (SWDTool == "DifferentMUPSCCM")
                {
                    ExcelManager.ExportDataToExcel("NotExistAtmsSCCM", dataTable2, new System.IO.DirectoryInfo(@"C:\"));
                    MessageBox.Show(@"NotExistAtmsSCCM.xlsx exported to c:\");
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        #endregion
    }
}
