using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSharp_ReaderWriterHelper
{
    /// <summary>
    /// Excel Operations Class
    /// </summary>
    public class ExcelHelper
    {
        /// <summary>
        /// Excel Application object
        /// </summary>
        private static Excel.Application xlsApp;

        /// <summary>
        /// Excel Workbook object 
        /// </summary>
        private static Excel.Workbook xlsWorkbook;

        /// <summary>
        /// Excel Worksheet object
        /// </summary>
        private static Excel._Worksheet xlsWorksheet;

        /// <summary>
        /// Excel Range object
        /// </summary>
        private static Excel.Range xlsRange;

        /// <summary>
        /// Create an excel application object
        /// </summary>
        private static void CreateExcelAplication()
        {
            // Try to create new excel application object
            try
            {
                xlsApp = new Excel.Application();
            }
            catch (Exception ex)
            {
                // Throw the exception
                throw ex;
            }
        }

        /// <summary>
        /// Reads the excel file and returns a datatable
        /// </summary>
        /// <param name="filePath">Excel file's filepath</param>
        /// /// <param name="worksheetNumber">Worksheet number of excel</param>
        /// <returns>DataTable object</returns>
        public static DataTable ReadExcel(string filePath, int worksheetNumber)
        {
            // Create an excell application object
            CreateExcelAplication();

            // Create a datatable object
            DataTable dtExcel = new DataTable();

            try
            {
                // Read the excel file
                xlsWorkbook = xlsApp.Workbooks.Open(filePath);

                // Get the worksheet which number is worksheetNumber
                xlsWorksheet = xlsWorkbook.Sheets[worksheetNumber];

                // Get the range of xlsWorksheet object
                xlsRange = xlsWorksheet.UsedRange;

                // Row count of excel file
                int rowCount = xlsRange.Rows.Count;
                // Column count of excel file
                int colCount = xlsRange.Columns.Count;

                // Add column to dtExcel object for the number of excel file's column count
                for (int i = 0; i < colCount; i++)
                {
                    dtExcel.Columns.Add();
                }

                // Add row to dtExcel object for the number of excel file's row count
                for (int i = 0; i < rowCount; i++)
                {
                    dtExcel.Rows.Add();
                }

                // Add datas of excel to the dtExcel datatable object
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        dtExcel.Rows[i][j] = xlsRange.Cells[i + 1, j + 1].Value2;
                    }
                }

                // Close the workBook
                xlsWorkbook.Close();
            }
            catch (Exception ex)
            {
                // Throw the exception
                throw ex;
            }

            // Kill the excel operations if exist
            ExcelKill();

            // Doldurulan excel nesnesi geri donderilir
            return dtExcel;
        }

        /// <summary>
        /// Saves the datatable as an excel file which specifies with filepath
        /// </summary>
        /// <param name="dataTable">DataTable object</param>
        /// <param name="filePath">Filepath</param>
        public static void SaveExcel(DataTable dataTable, string filePath)
        {
            // Create an excell application object
            CreateExcelAplication();

            // Missing value object
            object misValue = System.Reflection.Missing.Value;

            try
            {
                Excel.Workbook xlsWorkbook2;

                // Read the excel file
                xlsWorkbook2 = xlsApp.Workbooks.Add(misValue);

                // Get the worksheet which number is 1
                xlsWorksheet = (Excel.Worksheet)xlsWorkbook2.Worksheets.get_Item(1);

                // Set the name of worksheet
                xlsWorksheet.Name = "Sheet_name";

                // Add the datas into worksheet
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        xlsWorksheet.Cells[i + 1, j + 1] = dataTable.Rows[i][j];
                    }
                }

                // Save the workbook object
                xlsWorkbook2.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                        Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                // Close the workBook
                xlsWorkbook2.Close();

            }
            catch (Exception ex)
            {
                // Throw the exception
                throw ex;
            }

            // Kill the excel operations if exist
            ExcelKill();
        }

        /// <summary>
        /// Kills excel operations in system
        /// </summary>
        public static void ExcelKill()
        {
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcesses())
            {
                if (p.ProcessName == "EXCEL")
                {
                    p.Kill();
                }
            }
        }
    }
}
