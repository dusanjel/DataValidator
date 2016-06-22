using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace DataValidator
{
    public static class ExcelManager
    {
        public static List<Atm> ImprortAtmDataFromEpmReport(string filePath)
        {
            List<Atm> atms = new List<Atm>();

            // Get the file we are going to process
            var existingFile = new FileInfo(filePath);
            // Open and read the XlSX file.
            using (var package = new ExcelPackage(existingFile))
            {
                // Get the work book in the file
                var workBook = package.Workbook;
                if (workBook != null)
                {
                    if (workBook.Worksheets.Count > 0)
                    {
                        // Get the inventory worksheet
                        var worksheet = workBook.Worksheets["Inventory Report"];

                        // read some data
                        object col1Header = worksheet.Cells[1, 1].Value;

                        int rows = worksheet.Dimension.Rows;
                        for (int i = 2; i <= rows; i++)
                        {
                            var customer = worksheet.Cells[i, 1].Value;
                            var atmName = worksheet.Cells[i, 2].Value;
                            var software = worksheet.Cells[i, 3].Value;
                            var version = worksheet.Cells[i, 4].Value;
                            var date = worksheet.Cells[i, 5].Value;

                            if(software != null && software.ToString() == "StandardBase-CD2-MUP")
                            {
                                if (!atms.Exists(x => x.Name == atmName.ToString()))
                                {
                                    Atm atm = new Atm();
                                    if (customer != null) atm.Customer = customer.ToString();
                                    if (atmName != null) atm.Name = atmName.ToString();
                                    if (version != null) atm.AptraCD2Version = version.ToString().Substring(6);                                   
                                    atms.Add(atm);
                                }
                            }
                            
                        }
                    }
                }
            }
            return atms;
        }

        internal static List<Atm> ImprortAtmDataFromSCCMReport(string filePath)
        {
            List<Atm> atms = new List<Atm>();

            // Get the file we are going to process
            var existingFile = new FileInfo(filePath);
            // Open and read the XlSX file.
            using (var package = new ExcelPackage(existingFile))
            {
                // Get the work book in the file
                var workBook = package.Workbook;
                if (workBook != null)
                {
                    if (workBook.Worksheets.Count > 0)
                    {
                        // Get the inventory worksheet
                        var worksheet = workBook.Worksheets["SCCM"];

                        // read some data
                        object col1Header = worksheet.Cells[1, 1].Value;

                        int rows = worksheet.Dimension.Rows;
                        for (int i = 2; i <= rows; i++)
                        {
                            var customer = worksheet.Cells[i, 5].Value;
                            var atmName = worksheet.Cells[i, 1].Value;
                            var software = worksheet.Cells[i, 3].Value; // for SCCM it pulls InstallTime0 column
                            var version = worksheet.Cells[i, 4].Value;
                            var date = worksheet.Cells[i, 2].Value;

                            if (software != null)
                            {
                                if (!atms.Exists(x => x.Name == atmName.ToString()))
                                {
                                    Atm atm = new Atm();
                                    if (customer != null) atm.Customer = customer.ToString();
                                    if (atmName != null) atm.Name = atmName.ToString();
                                    if (version != null) atm.AptraCD2Version = version.ToString().Substring(6);                                    
                                    atms.Add(atm);
                                }
                            }

                        }
                    }
                }
            }
            return atms;
        }

        public static List<Atm> ImprortAtmDataFromSharepoint(string filePath)
        {
            List<Atm> atms = new List<Atm>();

            // Get the file we are going to process
            var existingFile = new FileInfo(filePath);
            // Open and read the XlSX file.
            using (var package = new ExcelPackage(existingFile))
            {
                // Get the work book in the file
                var workBook = package.Workbook;
                if (workBook != null)
                {
                    if (workBook.Worksheets.Count > 0)
                    {
                        // Get the inventory worksheet
                        var worksheet = workBook.Worksheets.First();

                        // read some data
                        object col1Header = worksheet.Cells[1, 1].Value;

                        int rows = worksheet.Dimension.Rows;
                        for (int i = 2; i <= rows; i++)
                        {
                            var customer = worksheet.Cells[i, 1].Value;
                            var atmName = worksheet.Cells[i, 3].Value;
                            var version = worksheet.Cells[i, 4].Value;

                            if (!atms.Exists(x => x.Name == atmName.ToString()))
                            {
                                Atm atm = new Atm();
                                if (customer != null) atm.Customer = customer.ToString();
                                if (atmName != null) atm.Name = atmName.ToString();
                                if (version != null) atm.AptraCD2Version = version.ToString(); // Dusan debug
                                atms.Add(atm);
                            }

                        }
                    }
                }
            }
            return atms;
        }

        public static string ExportDataToExcel(string fileName, DataTable dataTable, DirectoryInfo outputDir)
        {
            FileInfo newFile = new FileInfo(outputDir.FullName + string.Format(@"\{0}.xlsx", fileName));
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputDir.FullName + string.Format(@"\{0}.xlsx", fileName));
            }

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Output");

                DataTable dt = new DataTable();
                dt = dataTable;
                int dt_numberOfRows = dataTable.Rows.Count;
                int dt_numberOfColumns = dataTable.Columns.Count;
                int i = 1, j = 1;

                //Add column names to excel
                IEnumerable<string> columnNames = dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
                foreach (string columnName in columnNames)
                {
                    worksheet.Cells[1, i++].Value = columnName;
                }

                //populate excel with data
                for (int a = 0; a < dt_numberOfRows; a++)
                {
                    for (int b = 0; b < dt_numberOfColumns; b++)
                    {
                        worksheet.Cells[a + 2, b + 1].Value = dataTable.Rows[a][b].ToString();
                    }
                }

                //format excel sheet
                //Set a border around
                worksheet.Cells[1, 1, dt_numberOfRows + 1, dt_numberOfColumns].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, dt_numberOfRows + 1, dt_numberOfColumns].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, dt_numberOfRows + 1, dt_numberOfColumns].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, dt_numberOfRows + 1, dt_numberOfColumns].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, 1, dt_numberOfColumns].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, 1, 1, dt_numberOfColumns].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                //format document
                worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells

                // set some document properties
                package.Workbook.Properties.Title = fileName;

                // set some extended property values
                package.Workbook.Properties.Company = "NCR";

                // save our new workbook and we are done!
                package.Save();
            }

            return newFile.FullName;
        }

        /// <summary>
        /// Converts list to DataTable
        /// </summary>
        /// <param name="list">object list</param>
        /// <returns>Datatable, if list is empty return null</returns>
        public static DataTable ConvertListToDataTable(IList list)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //Get all the properties

                if (list.Count == 0) return null;

                object obj = list[0];
                PropertyDescriptorCollection props = TypeDescriptor.GetProperties(obj.GetType());

                //Create heeader
                for (int i = 0; i < props.Count; i++)
                {
                    PropertyDescriptor prop = props[i];
                    if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                        dataTable.Columns.Add(prop.Name, prop.PropertyType.GetGenericArguments()[0]);
                    else
                        dataTable.Columns.Add(prop.Name, prop.PropertyType);
                }

                //Create rows
                object[] values = new object[props.Count];
                foreach (object item in list)
                {
                    for (int i = 0; i < values.Length; i++)
                    {
                        values[i] = props[i].GetValue(item);
                    }
                    dataTable.Rows.Add(values);
                }

                return dataTable;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
