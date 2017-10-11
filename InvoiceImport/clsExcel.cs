using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace InvoiceImport
{
    class clsExcel
    {
        Type xlType;
        dynamic xlApp;
        dynamic xlWorkbook;
        dynamic xlWorksheet;
        dynamic xlRange;
        //public dynamic xlCell;
        //public dynamic xlRow;
        //public dynamic xlCol;

        private bool workbookOpen = false;

        public clsExcel(string workbookName, bool visible = true)
        {
            WorkbookName = workbookName;
            Visible = visible;
        }

        public clsExcel()
        {

        }

        ~clsExcel()
        {
            release_objects();
        }

        #region properties

        public bool Visible { get; set; }
        public string WorkbookName { get; set; }
        public string LastError { get; set; }
        public dynamic WorkBook { get { return xlWorkbook; } }
        public dynamic WorkSheet { get { return xlWorksheet; } }
        public dynamic Range { get { return xlRange; } }

        #endregion


        #region methods

        /// <summary>
        /// close excel file, save if needed
        /// </summary>
        /// <param name="save"></param>
        /// <returns>boolean indicating success status</returns>
        public bool CloseWorkbook(bool save = false)
        {
            bool result = false;
            try
            {
                try
                {
                    xlWorkbook.Close(save);
                    xlWorkbook = null;
                }
                catch(COMException)
                {
                    // ignore if already closed
                }
                workbookOpen = false;
                LastError = "";
                result = true;
            }
            catch(Exception ex)
            {
                LastError = $"Error closing excel file: {ex.Message}";
                result = false;
            }
            return result;
        }

        /// <summary>
        /// start ms excel and open supplied workbook name
        /// </summary>
        /// <param name="fileName">full path-name of file to open</param>
        /// <param name="sheet">name or number of worksheet to activate</param>
        /// <returns>boolean indicating success status</returns>
        public bool OpenExcel(string fileName, string sheet = "")
        {
            var xlFile = Path.GetFileName(fileName);
            bool result = false;
            if(File.Exists(fileName))
            {
                int sheetNo = 0;
                try
                {
                    xlType = Type.GetTypeFromProgID("Excel.Application");
                    xlApp = Activator.CreateInstance(xlType);
                    xlApp.Visible = Visible;
                    xlWorkbook = xlApp.Workbooks.Open(fileName);
                    if(sheet == "")
                        xlWorksheet = xlWorkbook.Sheets[1];
                    else if(int.TryParse(sheet, out sheetNo))
                        xlWorksheet = xlWorkbook.Sheets[sheetNo];
                    else
                    {
                        try
                        {
                            xlWorksheet = xlWorkbook.Sheets[sheet];
                        }
                        catch(Exception)
                        {
                            xlWorksheet = xlWorkbook.Sheets[1];
                        }
                    }
                    xlWorksheet.Activate();
                    workbookOpen = true;
                    LastError = "";
                    result = true;
                }
                catch(Exception ex)
                {
                    LastError = $"Error opening Excel file \"{xlFile}\": {ex.Message}";
                    result = false;
                }
            }
            else
            {
                LastError = $"Could not find Excel file \"{xlFile}\"";
                result = false;
            }
            return result;
        }

        public bool CloseExcel()
        {
            bool result = true;
            if(workbookOpen)
                result = CloseWorkbook();

            release_objects();

            return result;
        }

        public bool GetRange(string nameOrAddress, bool select = false)
        {
            bool result = false;

            if(xlApp != null || xlWorkbook != null)
            {
                try
                {
                    xlRange = xlApp.Range(nameOrAddress);
                    if(select)
                        xlRange.Activate();

                    result = true;
                    LastError = "";
                }
                catch(Exception ex)
                {
                    LastError = ex.Message;
                }
            }
            return result;
        }

        /// <summary>
        /// release com objects to fully kill excel process from running in the background
        /// </summary>
        private void release_objects()
        {
            //try
            //{
            //    Marshal.ReleaseComObject(xlCell);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            //try
            //{
            //    Marshal.ReleaseComObject(xlCol);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            //try
            //{
            //    Marshal.ReleaseComObject(xlRow);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            //try
            //{
            //    Marshal.ReleaseComObject(xlRange);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            //try
            //{
            //    Marshal.ReleaseComObject(xlWorksheet);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            //// quit and release
            //try
            //{
            //    xlApp.Quit();
            //    Marshal.ReleaseComObject(xlApp);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            try
            {
                xlRange = null;
                xlWorksheet = null;
                xlWorkbook = null;
                xlApp.quit();
                xlApp = null;

            }
            catch(Exception)
            {
            }
            finally
            {
            GC.Collect();
            }

        }

        #endregion

    }
}
