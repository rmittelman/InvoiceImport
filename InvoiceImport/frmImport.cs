using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using Aimm.Logging;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using QuickBooks;

namespace InvoiceImport
{
    public partial class frmImport : Form
    {
        // used to scroll text box if needed.
        //[DllImport("user32.dll", CharSet = CharSet.Auto)]
        //private static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, ref Point lParam);

        public frmImport()
        {
            InitializeComponent();
        }

        enum cols
        {
            vendor = 1,
            invNum = 2,
            invDate = 3,
            jobID = 4,
            woNum = 5,
            invAmt = 6,
            invDesc = 7,
            vendorID = 8,
            fullName = 9,
            billFrom1 = 10,
            billFrom2 = 11,
            billFrom3 = 12,
            billFrom4 = 13,
            billFrom5 = 14,
            terms = 15,
            expAcct = 16
        }

        // to scroll textbox
        //private int WM_VSCROLL= 277;
        //private System.IntPtr SB_BOTTOM = (IntPtr)7;
        //private Point pt = new Point();

        string connString = Properties.Settings.Default.POLSQL;
        private static string newLogLine = "";

        bool isValid = false;
        bool allValid = true;

        string sourcePath = Properties.Settings.Default.SourceFolder;
        string archivePath = Properties.Settings.Default.ArchiveFolder;
        string errorPath = Properties.Settings.Default.ErrorFolder;
        string logPath = Properties.Settings.Default.LogFolder;
        string pdfPath = Properties.Settings.Default.PdfFolder;
        bool showExcel = (bool)Properties.Settings.Default.ShowExcel;
        string xlPathName = "";
        string xlFile = "";
        string pdfPathName = "";
        string pdfFile = "";
        string destPath = "";
        string destFile = "";
        string destPathName = "";
        string logFile = "InvoiceImport.log";
        string logPathName = "";
        Excel.Application xlApp = null;
        Excel.Workbook xlWorkbook = null;
        Excel._Worksheet xlWorksheet = null;
        Excel.Range xlRange = null;
        Excel.Range xlCell = null;

        ToolTip toolTip1 = new ToolTip();


        private void frmImport_Load(object sender, EventArgs e)
        {

            // set tooltips
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            toolTip1.SetToolTip(this.btnFindExcel, "Find Excel Invoices File");
            toolTip1.SetToolTip(this.btnFindPDF, "Find PDF Invoices File");
            toolTip1.SetToolTip(this.btnImport, "Import Invoices from Excel File");
            toolTip1.SetToolTip(this.btnSplitPDF, "Split PDF File into Multiple Files");

            logPathName = Path.Combine(logPath, logFile);

            // to monitor log
            //txtLog.Lines = File.ReadAllLines(logPathName);
            //MonitorLog(logPathName);

            // event handlers
            btnFindExcel.Click += btnFindFile_Click;
            btnFindPDF.Click += btnFindFile_Click;

            // to monitor log
            //txtLog.TextChanged += txtLog_TextChanged;

            LogIt.LogMethod();

        }

        // to monitor log
        //private void txtLog_TextChanged(object sender, EventArgs e)
        //{

        //    SendMessage(txtLog.Handle, WM_VSCROLL, SB_BOTTOM, ref pt);
        //}

        private void btnFindFile_Click(object sender, EventArgs e)
        {
            var btn = sender as Button;
            if(btn != null)
            {
                var btnName = btn.Name;

                using(OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.InitialDirectory = sourcePath;
                    if(btn.Name == "btnFindExcel")
                        ofd.Filter = "Excel files (*.xlsx, *.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*";
                    else
                        ofd.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*";
                    ofd.FilterIndex = 1;
                    if(ofd.ShowDialog() == DialogResult.OK)
                    {
                        if(btn.Name == "btnFindExcel")
                            txtExcelFile.Text = ofd.FileName;
                        else
                            txtPDFFile.Text = ofd.FileName;
                    }
                }
            }
        }

        private void btnFindFolder_Click(object sender, EventArgs e)
        {

            using(FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Select the folder to save PDF files to.";
                fbd.SelectedPath = pdfPath;
                if(fbd.ShowDialog() == DialogResult.OK)
                {
                    txtPDFFolder.Text = fbd.SelectedPath;
                }
            }

        }


        private bool move_file(string sourcePath, string destPath)
        {
            try
            {
                File.Move(sourcePath, destPath);
                return true;
            }
            catch(Exception ex)
            {
                var msg = $"Error moving file \"{Path.GetFileName(sourcePath)}\" to \"{Path.GetDirectoryName(sourcePath)}\": {ex.Message}";
                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                LogIt.LogError(msg);
                return false;
            }
        }

        /// <summary>
        /// start ms excel and open supplied workbook name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>boolean indicating success status</returns>
        private bool open_excel(string fileName)
        {
            if(File.Exists(fileName))
            {
                try
                {
                    xlApp = new Excel.Application();
                    xlApp.Visible = showExcel;
                    xlWorkbook = xlApp.Workbooks.Open(xlPathName);
                    xlWorksheet = xlWorkbook.Sheets[1];
                    xlRange = xlApp.get_Range("invoice_info");
                    LogIt.LogInfo($"Opened Excel file \"{xlFile}\"");
                    return true;
                }
                catch(Exception ex)
                {
                    var msg = $"Error opening Excel file \"{xlFile}\": {ex.Message}";
                    MessageBox.Show(msg, "Error", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                    LogIt.LogError(msg);
                    return false;
                }
            }
            else
            {
                var msg = $"Could not find Excel file \"{xlFile}\"";
                LogIt.LogError(msg);
                return false;
            }
        }

        /// <summary>
        /// close excel file, save if needed, kill objects
        /// </summary>
        /// <param name="needToSave"></param>
        /// <returns>boolean indicating success status</returns>
        private bool close_excel(bool needToSave = false)
        {
            try
            {
                // close workbook, cleanup excel
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // close and release

                try
                {
                    xlWorkbook.Close(needToSave);

                }
                catch(COMException ex)
                {
                    // ignore if already closed
                }

                // release com objects to fully kill excel process from running in the background
                try
                {
                    Marshal.ReleaseComObject(xlCell);
                }
                catch(NullReferenceException ex)
                {
                    // ignore if not yet instantiated
                }
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

                // quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                LogIt.LogInfo($"Closed Excel file, save = {needToSave}");
                return true;
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error closing excel file: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// add an invoice to a job / work order
        /// </summary>
        /// <param name="jobID"></param>
        /// <param name="woNo"></param>
        /// <param name="vendorID"></param>
        /// <param name="invNo"></param>
        /// <param name="invAmt"></param>
        /// <param name="invDesc"></param>
        /// <param name="connectionString"></param>
        /// <returns>string job# + work order # (12345-W01)</returns>
        private static string add_invoice_to_job(int jobID, string woNo, int vendorID, string invNo, Single invAmt, string invDesc, string connectionString)
        {
            string projFinalID = jobID.ToString() + "-" + woNo;
            string sql_insert = "INSERT INTO POL.tblProjectFinalMatEquip ( "
                       + "ProjectFinalID,JobID,BuildingMaterialID,OtherMaterial,CostEach,"
                       + "Quantity,TotalCost,Notes,EnteredDate,Correction,JobErrorID ) "
                       + "VALUES ( '{0}','{1}',{2},'{3}',{4},1,{5},'{6}','{7}',0,NULL );";

            string sql = string.Format(sql_insert, projFinalID, jobID.ToString(), vendorID, invNo, invAmt, invAmt, invDesc, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    using(SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        return projFinalID;
                    }
                }

            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error adding invoice for job {projFinalID}: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// verify work order belongs to job
        /// </summary>
        /// <param name="jobID"></param>
        /// <param name="woNo"></param>
        /// <returns>job status or -1 if job not found</returns>
        private static bool valid_work_order(int jobID, string woNo, string connectionString)
        {
            bool isValid = false;
            string projFinalID = $"{jobID}-{woNo}";
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    LogIt.LogInfo($"Validating work order \"{woNo}\" for job {jobID}");
                    string cmdText = "SELECT COUNT(ProjectFinalID) FROM MLG.POL.tblProjectFinal where ProjectFinalID = @projFinalID";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.Parameters.AddWithValue("@projFinalID", projFinalID);
                        conn.Open();
                        int rows = (int)cmd.ExecuteScalar();
                        isValid = (rows > 0);
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error validating work order \"{woNo}\" for job {jobID}: {ex.Message}");
            }
            return isValid;
        }

        /// <summary>
        /// get job status for supplied job number
        /// </summary>
        /// <param name="jobID"></param>
        /// <returns>job status or -1 if job not found</returns>
        private static object get_job_status(int jobID, string connectionString)
        {
            var status = -1;
            SqlDataReader reader = null;
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {

                    LogIt.LogInfo($"Getting status for job {jobID}");
                    string cmdText = "SELECT JobStatusID FROM MLG.POL.tblJobs where JobID = @jobID";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.Parameters.AddWithValue("@jobID", jobID);
                        conn.Open();
                        reader = cmd.ExecuteReader();
                        if(reader.HasRows)
                        {
                            while(reader.Read())
                            {
                                var record = (IDataRecord)reader;
                                var jobStatus = (int)record[0];
                                if(jobStatus != 7 & jobStatus != 11)
                                {
                                    status = jobStatus;
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error getting status for job {jobID}: {ex.Message}");
            }
            finally
            {
                try
                {
                    reader.Close();
                }
                catch(Exception)
                {

                }
            }
            return status;
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            xlPathName = txtExcelFile.Text;
            if(open_excel(xlPathName))
            {
                var xlFile = Path.GetFileName(xlPathName);
                try
                {
                    allValid = true;

                    // loop thru each invoice row on worksheet
                    foreach(Excel.Range xlRow in xlRange.Rows)
                    {
                        isValid = true;

                        // get non-validated data for the invoice
                        xlCell = (Excel.Range)xlRow.Cells[cols.vendor];
                        string vendor = (xlCell.Value2 ?? "").ToString();

                        xlCell = (Excel.Range)xlRow.Cells[cols.invNum];
                        string invNo = (xlCell.Value2 ?? "").ToString();
                        if(invNo == "5299")
                        {

                        }
                        xlCell = (Excel.Range)xlRow.Cells[cols.invDesc];
                        string invDesc = (xlCell.Value2 ?? "").ToString();
                        if(invDesc == "0")
                            invDesc = "";

                        // if any blank items, we're done
                        if(vendor == "" || invNo == "")
                            break;

                        // get and validate remaining items
                        DateTime invDate;
                        int jobID = 0;
                        string woNo = "";
                        Single invAmt;
                        string jobWO;
                        int vendorID;

                        // validate invoice date is date
                        // try both date number and string date conversion just in case
                        xlCell = (Excel.Range)xlRow.Cells[cols.invDate];
                        if(xlCell.Value2 == null)
                            isValid = false;
                        else
                        {
                            try
                            {
                                invDate = DateTime.FromOADate(xlCell.Value2);
                                isValid = true;
                            }
                            catch(Exception)
                            {
                                isValid = DateTime.TryParse(xlCell.Value2.ToString(), out invDate);
                            }

                        }
                        if(isValid)
                        {
                            // validate jobID is int
                            xlCell = (Excel.Range)xlRow.Cells[cols.jobID];

                            if(xlCell.Value2 != null && int.TryParse(xlCell.Value2.ToString(), out jobID))
                            {
                                var jobStatus = (int)get_job_status(jobID, connString);
                                if(jobStatus != -1)
                                {
                                    // validate WO belongs to job
                                    xlCell = (Excel.Range)xlRow.Cells[cols.woNum];
                                    woNo = xlCell.Value2.ToString().ToUpper();
                                    isValid = valid_work_order(jobID, woNo, connString);
                                    if(isValid)
                                    {
                                        // validate invAmt is numeric
                                        xlCell = (Excel.Range)xlRow.Cells[cols.invAmt];
                                        if(Single.TryParse(xlCell.Value2.ToString(), out invAmt))
                                        {
                                            // if vendor missing, warn & continue
                                            xlCell = (Excel.Range)xlRow.Cells[cols.vendorID];
                                            if(int.TryParse(xlCell.Value2.ToString(), out vendorID))
                                            {
                                                // add the invoice, return formatted work order #
                                                if(vendorID == 0)
                                                {
                                                    jobWO = add_invoice_to_job(jobID, woNo, vendorID, null, invAmt, invDesc, connString);
                                                    ((Excel.Range)xlRow.Cells[cols.vendorID]).Interior.ColorIndex = 53;
                                                    LogIt.LogWarn($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has vendor missing from AIMM");
                                                    isValid = false;
                                                }
                                                else
                                                    jobWO = add_invoice_to_job(jobID, woNo, vendorID, invNo, invAmt, invDesc, connString);

                                                if(jobWO.Length != 0)
                                                {
                                                    LogIt.LogInfo($"Added invoice {invNo} for vendor {vendor} to job {jobWO}");
                                                }
                                                else
                                                {
                                                    LogIt.LogError($"Couldn't add invoice {invNo} for vendor {vendor} to job ID {jobWO}");
                                                    isValid = false;
                                                } // added invoice

                                            }
                                            else
                                            {
                                                isValid = false;
                                                ((Excel.Range)xlRow.Cells[cols.vendorID]).Interior.ColorIndex = 3;
                                                LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad vendor ID: {xlCell.Value2.ToString()}");
                                            } // vendor ID is numeric

                                        }

                                        else
                                        {
                                            isValid = false;
                                            ((Excel.Range)xlRow.Cells[cols.invAmt]).Interior.ColorIndex = 3;
                                            LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad invoice amount: {xlCell.Value2.ToString()}");
                                        } // inv amt is numeric

                                    }
                                    else
                                    {
                                        ((Excel.Range)xlRow.Cells[cols.woNum]).Interior.ColorIndex = 3;
                                        LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has invalid work order number: {woNo}");
                                    } // WO belongs to job

                                }
                                else
                                {
                                    isValid = false;
                                    ((Excel.Range)xlRow.Cells[cols.jobID]).Interior.ColorIndex = 3;
                                    LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\", job ID \"{jobID}\" is closed, cancelled or missing from database");
                                } // valid job status
                            }
                            else
                            {
                                isValid = false;
                                ((Excel.Range)xlRow.Cells[cols.jobID]).Interior.ColorIndex = 3;
                                ((Excel.Range)xlRow.Cells[cols.woNum]).Interior.ColorIndex = 3;
                                LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad JobID: {xlCell.Value2.ToString()}");
                            } // valid job id

                        }
                        else
                        {
                            isValid = false;
                            ((Excel.Range)xlRow.Cells[cols.invDate]).Interior.ColorIndex = 3;
                            LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlCell.Value2.ToString()}\" has bad date: {3}");
                        } // valid date

                        allValid = allValid && isValid;
                    }

                    var isOk = close_excel(!allValid);

                    // move workbook to archive/errors folder
                    destPath = allValid ? archivePath : errorPath;
                    destFile = string.Concat(
                        Path.GetFileNameWithoutExtension(xlFile),
                        DateTime.Now.ToString("_yyyy-MM-dd_HH-mm-ss"),
                        Path.GetExtension(xlFile));
                    destPathName = Path.Combine(destPath, destFile);
                    if(move_file(xlPathName, destPathName))
                    {
                        txtExcelFile.Text = destPathName;
                        if(allValid)
                            LogIt.LogInfo($"Moved \"{xlFile}\" to \"{destPathName}\"");
                        else
                            LogIt.LogWarn($"File \"{xlFile}\" had errors. Moved it to \"{destPathName}\"");
                    }

                }
                catch(Exception ex)
                {
                    var msg = $"Error processing Excel file \"{xlFile}\": {ex.Message}";
                    MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    LogIt.LogError(msg);
                }
            }
            else
            {
                var msg = $"Could not find Excel file \"{xlPathName}\"";
                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void btnSplitPDF_Click(object sender, EventArgs e)
        {
            pdfPathName = txtPDFFile.Text;
            if(File.Exists(pdfPathName))
            {
                pdfFile = Path.GetFileName(pdfPathName);

                // get list of invoice numbers from xl document.
                xlPathName = txtExcelFile.Text;
                var xlFile = Path.GetFileName(xlPathName);
                if(open_excel(xlPathName))
                {
                    var invoiceList = new List<string>();
                    try
                    {
                        // loop thru each invoice row on worksheet, collect invoice numbers
                        foreach(Excel.Range xlRow in xlRange.Rows)
                        {
                            xlCell = (Excel.Range)xlRow.Cells[cols.vendor];
                            string vendor = (xlCell.Value2 ?? "").ToString();
                            xlCell = (Excel.Range)xlRow.Cells[cols.invNum];
                            string invNo = (xlCell.Value2 ?? "").ToString();
                            if(vendor == "" || invNo == "")
                                break;
                            //xlCell = (Excel.Range)xlRow.Cells[cols.vendorID];
                            //string vendorId = (xlCell.Value2 ?? "").ToString();
                            invoiceList.Add($"{vendor}_Invoice_{invNo}");
                        }
                        var isOk = close_excel();
                    }
                    catch(Exception ex)
                    {
                        var msg = $"Error processing Excel file \"{xlFile}\": {ex.Message}";
                        MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }

                    try
                    {
                        // split the supplied PDF into separate page documents
                        using(PdfDocument combinedPdf = PdfReader.Open(pdfPathName, PdfDocumentOpenMode.Import))
                        {
                            for(int pg = 0; pg < combinedPdf.PageCount; pg++)
                            {
                                using(PdfDocument pageDoc = new PdfDocument())
                                {
                                    pageDoc.Version = combinedPdf.Version;
                                    pageDoc.Info.Title = invoiceList[pg];
                                    pageDoc.Info.Creator = combinedPdf.Info.Creator;
                                    pageDoc.AddPage(combinedPdf.Pages[pg]);
                                    destFile = $"{invoiceList[pg]}.pdf";
                                    destPathName = Path.Combine(pdfPath, destFile);
                                    pageDoc.Save(destPathName);
                                }
                            }
                            LogIt.LogInfo($"Split PDF document \"{pdfFile}\" into multiple documents");
                        }
                    }
                    catch(Exception ex)
                    {
                        var msg = $"Error processing PDF file \"{pdfFile}\": {ex.Message}";
                        MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        LogIt.LogError($"Error processing PDF file \"{pdfFile}\": {ex.Message}");
                    }

                    // move original PDF to archive folder
                    destPath = archivePath;
                    destFile = string.Concat(
                        Path.GetFileNameWithoutExtension(pdfFile),
                        DateTime.Now.ToString("_yyyy-MM-dd_HH-mm-ss"),
                        Path.GetExtension(pdfFile));
                    destPathName = Path.Combine(destPath, destFile);
                    if(move_file(pdfPathName, destPathName))
                    {
                        txtPDFFile.Text = destPathName;
                        LogIt.LogInfo($"Moved \"{pdfFile}\" to \"{destPathName}\"");
                    }

                }
                else
                {
                    var msg = $"Could not find or open Excel file \"{xlFile}\"";
                    MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

            }
            else
            {
                var msg = $"Could not find PDF file \"{pdfPathName}\"";
                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnQuickBooks_Click(object sender, EventArgs e)
        {
            xlPathName = txtExcelFile.Text;
            var xlFile = Path.GetFileName(xlPathName);
            if(open_excel(xlPathName))
            {
                var billList = new List<BillData>(); // define list here
                try
                {
                    allValid = true;

                    // loop thru each invoice row on worksheet
                    foreach(Excel.Range xlRow in xlRange.Rows)
                    {
                        isValid = true;

                        // get non-validated data for the invoice
                        xlCell = (Excel.Range)xlRow.Cells[cols.fullName];
                        string vendor = (xlCell.Value2 ?? "").ToString();

                        xlCell = (Excel.Range)xlRow.Cells[cols.invNum];
                        string invNo = (xlCell.Value2 ?? "").ToString();

                        // if blank items, we're done
                        if(vendor == "" || invNo == "")
                            break;

                        xlCell = (Excel.Range)xlRow.Cells[cols.fullName];
                        string fullName = (xlCell.Value2 ?? "").ToString();
                        if(fullName == "0")
                            fullName = "";

                        xlCell = (Excel.Range)xlRow.Cells[cols.billFrom1];
                        string billFrom1 = (xlCell.Value2 ?? "").ToString();
                        if(billFrom1 == "0")
                            billFrom1 = "";

                        xlCell = (Excel.Range)xlRow.Cells[cols.billFrom2];
                        string billFrom2 = (xlCell.Value2 ?? "").ToString();
                        if(billFrom2 == "0")
                            billFrom2 = "";

                        xlCell = (Excel.Range)xlRow.Cells[cols.billFrom3];
                        string billFrom3 = (xlCell.Value2 ?? "").ToString();
                        if(billFrom3 == "0")
                            billFrom3 = "";

                        xlCell = (Excel.Range)xlRow.Cells[cols.billFrom4];
                        string billFrom4 = (xlCell.Value2 ?? "").ToString();
                        if(billFrom4 == "0")
                            billFrom4 = "";

                        xlCell = (Excel.Range)xlRow.Cells[cols.billFrom5];
                        string billFrom5 = (xlCell.Value2 ?? "").ToString();
                        if(billFrom5 == "0")
                            billFrom5 = "";

                        // get and validate remaining items
                        DateTime invDate;
                        Single invAmt;
                        string expAcct;
                        int jobID;
                        string woNo;
                        int vendorID;

                        // validate invoice date is date
                        // try both date number and string date conversion just in case
                        xlCell = (Excel.Range)xlRow.Cells[cols.invDate];
                        if(xlCell.Value2 == null)
                            isValid = false;
                        else
                        {
                            try
                            {
                                invDate = DateTime.FromOADate(xlCell.Value2);
                                isValid = true;
                            }
                            catch(Exception)
                            {
                                isValid = DateTime.TryParse(xlCell.Value2.ToString(), out invDate);
                            }

                        }
                        if(isValid)
                        {
                            // validate invAmt is numeric
                            xlCell = (Excel.Range)xlRow.Cells[cols.invAmt];
                            if(Single.TryParse(xlCell.Value2.ToString(), out invAmt))
                            {

                                // get customer for job





                                //// validate exp acct is present
                                //xlCell = (Excel.Range)xlRow.Cells[cols.expAcct];
                                //expAcct = (xlCell.Value2 ?? "").ToString();
                                //if(expAcct == "0")
                                //    expAcct = "";

                                //if(expAcct == "")
                                //{
                                //    ((Excel.Range)xlRow.Cells[cols.invAmt]).Interior.ColorIndex = 53;
                                //    LogIt.LogWarn($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad invoice amount: {xlCell.Value2.ToString()}");
                                //}

                                //if(int.TryParse(xlCell.Value2.ToString(), out jobID))
                                //{
                                //    var jobStatus = (int)get_job_status(jobID, connString);
                                //    if(jobStatus != -1)
                                //    {
                                //        // validate WO belongs to job
                                //        xlCell = (Excel.Range)xlRow.Cells[cols.woNum];
                                //        woNo = xlCell.Value2.ToString().ToUpper();
                                //        isValid = valid_work_order(jobID, woNo, connString);
                                //        if(isValid)
                                //        {
                                //            // validate vendor ID is numeric
                                //            xlCell = (Excel.Range)xlRow.Cells[cols.vendorID];
                                //            if(Int32.TryParse(xlCell.Value2.ToString(), out vendorID))
                                //            {
                                //                // add the invoice, return formatted work order #
                                //                var jobWO = add_invoice_to_job(jobID, woNo, vendorID, invNo, invAmt, invDesc, connString);
                                //                if(jobWO.Length != 0)
                                //                {
                                //                    LogIt.LogInfo($"Added invoice {invNo} for vendor {vendor} to job {jobWO}");
                                //                    // add invoice to quickbooks???
                                //                }
                                //                else
                                //                {
                                //                    LogIt.LogError($"Couldn't add invoice {invNo} for vendor {vendor} to job ID {jobWO}");
                                //                    isValid = false;
                                //                } // added invoice

                                //            }
                                //            else
                                //            {
                                //                isValid = false;
                                //                ((Excel.Range)xlRow.Cells[cols.vendorID]).Interior.ColorIndex = 3;
                                //                LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad vendor ID: {xlCell.Value2.ToString()}");
                                //            } // vendor ID is numeric

                                //        }


                                //                }
                                //                else
                                //                {
                                //                    ((Excel.Range)xlRow.Cells[cols.woNum]).Interior.ColorIndex = 3;
                                //                    LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has invalid work order number: {woNo}");
                                //                } // WO belongs to job

                                //    }
                                //    else
                                //    {
                                //        isValid = false;
                                //        ((Excel.Range)xlRow.Cells[cols.jobID]).Interior.ColorIndex = 3;
                                //        LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\", job ID \"{jobID}\" is closed, cancelled or missing from database");
                                //    } // valid job status
                                //}
                                //else
                                //{
                                //    isValid = false;
                                //    ((Excel.Range)xlRow.Cells[cols.jobID]).Interior.ColorIndex = 3;
                                //    LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad JobID: {xlCell.Value2.ToString()}");
                                //} // valid job id

                            }

                            else
                            {
                                isValid = false;
                                ((Excel.Range)xlRow.Cells[cols.invAmt]).Interior.ColorIndex = 3;
                                LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad invoice amount: {xlCell.Value2.ToString()}");
                            } // inv amt is numeric


                        }
                        else
                        {
                            isValid = false;
                            ((Excel.Range)xlRow.Cells[cols.invDate]).Interior.ColorIndex = 3;
                            LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlCell.Value2.ToString()}\" has bad date: {3}");
                        } // valid date

                        allValid = allValid && isValid;
                    }

                    var isOk = close_excel(!allValid);

                    // move workbook to archive/errors folder
                    destPath = allValid ? archivePath : errorPath;
                    destFile = string.Concat(
                        Path.GetFileNameWithoutExtension(xlFile),
                        DateTime.Now.ToString("_yyyy-MM-dd_HH-mm-ss"),
                        Path.GetExtension(xlFile));
                    destPathName = Path.Combine(destPath, destFile);
                    if(move_file(xlPathName, destPathName))
                    {
                        txtExcelFile.Text = destPathName;
                        if(allValid)
                            LogIt.LogInfo($"Moved \"{xlFile}\" to \"{destPathName}\"");
                        else
                            LogIt.LogWarn($"File \"{xlFile}\" had errors. Moved it to \"{destPathName}\"");
                    }

                }
                catch(Exception ex)
                {
                    var msg = $"Error processing Excel file \"{xlFile}\": {ex.Message}";
                    MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    LogIt.LogError(msg);
                }
            }
            else
            {
                var msg = $"Could not find Excel file \"{xlPathName}\"";
                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        // the following is in case we want to show the log on the form.
        // for now, we're not doing that.

        ///// <summary>
        ///// watch log file for changes
        ///// </summary>
        ///// <param name="path"></param>
        //private static void MonitorLog(string path)
        //{
        //    FileSystemWatcher fileSystemWatcher = new FileSystemWatcher();
        //    fileSystemWatcher.Path = path;
        //    fileSystemWatcher.Created += FileSystemWatcher_Changed;
        //    fileSystemWatcher.Changed += FileSystemWatcher_Changed;
        //    fileSystemWatcher.EnableRaisingEvents = true;
        //}

        //private static void FileSystemWatcher_Changed(object sender, FileSystemEventArgs e)
        //{
        //    var data = File.ReadAllLines("C:\\test.log");
        //    string last = data[data.Length - 1];
        //    newLogLine = last;

        //}


    }

}
