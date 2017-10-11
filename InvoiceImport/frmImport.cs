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
using iTextSharp.text;
using iTextSharp.text.pdf;
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

        #region enums
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
            expAcct = 10,
            status = 11,
            message = 12
        }

        enum importTypes
        {
            standard,
            stock
        }

        #endregion

        #region objects

        clsExcel oXl = null;

        dynamic xlRange = null;
        dynamic xlCell = null;

        //Excel.Application xlApp = null;
        //Excel.Workbook xlWorkbook = null;
        //Excel._Worksheet xlWorksheet = null;
        //Excel.Range xlRange = null;
        //Excel.Range xlCell = null;
        ToolTip toolTip1 = new ToolTip();
        #endregion

        #region variables

        // to scroll textbox
        //private int WM_VSCROLL= 277;
        //private System.IntPtr SB_BOTTOM = (IntPtr)7;
        //private Point pt = new Point();

        string connString = Properties.Settings.Default.POLSQL;
        importTypes importType;

        // for displaying log
        // private static string newLogLine = "";

        bool isValid = false;
        bool qbValid = false;
        bool allValid = true;

        string sourcePath = Properties.Settings.Default.SourceFolder;
        string archivePath = Properties.Settings.Default.ArchiveFolder;
        string errorPath = Properties.Settings.Default.ErrorFolder;
        string logPath = Properties.Settings.Default.LogFolder;
        string pdfDestPath = Properties.Settings.Default.PdfFolder;
        bool showExcel = (bool)Properties.Settings.Default.ShowExcel;
        string apAcct = Properties.Settings.Default.APAcct;
        string billClass = Properties.Settings.Default.BillClass;
        string xlPathName = "";
        string xlFile = "";
        string pdfPathName = "";
        string pdfFile = "";
        string destPath = "";
        string destFile = "";
        string destPathName = "";
        string logFile = "InvoiceImport.log";
        string logPathName = "";

        #endregion

        #region properties

        public string Status { set { txtStatus.Text = value; } }

        #endregion

        #region events

        private void frmImport_Load(object sender, EventArgs e)
        {

            // set tooltips
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            toolTip1.SetToolTip(this.btnFindExcel, "Find Excel Invoices File");
            toolTip1.SetToolTip(this.btnFindPDF, "Find PDF Invoices File");
            toolTip1.SetToolTip(this.btnImportStandard, "Import Invoices from Excel File");
            toolTip1.SetToolTip(this.btnSplitPDF, "Split PDF File into Multiple Files");

            logPathName = Path.Combine(logPath, logFile);
            ckAddToAIMM.Checked = true;
            ckAddToQB.Checked = true;

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

        // this code is to monitor log, not currently being used
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
                fbd.SelectedPath = pdfDestPath;
                if(fbd.ShowDialog() == DialogResult.OK)
                {
                    txtPDFFolder.Text = fbd.SelectedPath;
                }
            }

        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if(((Button)sender).Name == "btnImportStandard")
                importType = importTypes.standard;
            else
                importType = importTypes.stock;

            var addToAimm = ckAddToAIMM.Checked;
            var addToQB = ckAddToQB.Checked;

            // continue if we can open excel file
            xlPathName = txtExcelFile.Text;
            if(open_excel(xlPathName))
            {

                // setup connection to quickbooks, continue if connected
                var qb = new QuickBooks.QuickBooks();
                qb.StatusChanged += qb_StatusChanged;
                if(qb.Connect())
                {

                    var billList = new List<BillData>();
                    var xlFile = Path.GetFileName(xlPathName);
                    try
                    {
                        allValid = true;

                        // loop thru each invoice row on worksheet
                        foreach(dynamic xlRow in oXl.Range.Rows)
                        {
                            isValid = true;
                            var billData = new BillData();
                            DateTime invDate = new DateTime();
                            string tempDate = "";
                            int jobID = 0;
                            string woNo = "";
                            Single invAmt = 0;
                            string jobWO = "";
                            int vendorID = 0;
                            billData.APAccount = apAcct;
                            billData.ClassRef = billClass;

                            // get first few items for the invoice
                            string vendor = (xlRow.Cells[cols.vendor].Value ?? "").ToString().Trim();

                            xlCell = xlRow.Cells[cols.invNum];
                            string invNo = (xlCell.Value ?? "").ToString().Trim();

                            // if blank items, we're done
                            if(vendor == "" & invNo == "")
                                break;

                            // validate invoice number has value
                            if(invNo != "")
                            {
                                billData.InvoiceNumber = invNo;
                                var msg = $"Got invoice {invNo} for vendor {vendor}";
                                Status = msg;
                                LogIt.LogInfo(msg);

                                string invDesc = (xlRow.Cells[cols.invDesc].Value ?? "").ToString().Trim();
                                if(invDesc == "0")
                                    invDesc = "";

                                // validate vendor full name has value
                                xlCell = xlRow.Cells[cols.fullName];
                                var vfn = (xlCell.Value ?? "").ToString().Trim();
                                if(vfn == "0")
                                    vfn = "";
                                if(vfn != "")
                                {
                                    billData.VendorFullName = vfn;


                                    // validate invoice date is date
                                    // try both date number and string date conversion just in case
                                    xlCell = xlRow.Cells[cols.invDate];
                                    tempDate = (xlCell.Value ?? "");
                                    if(tempDate == "")
                                        isValid = false;
                                    else
                                    {
                                        try
                                        {
                                            invDate = DateTime.FromOADate(xlCell.Value);
                                            isValid = true;
                                        }
                                        catch(Exception)
                                        {
                                            isValid = DateTime.TryParse(tempDate, out invDate);
                                        }

                                    }
                                    if(isValid)
                                    {
                                        billData.InvoiceDate = invDate;

                                        // validate jobID is int
                                        xlCell = xlRow.Cells[cols.jobID];
                                        var jid = (xlCell.Value ?? "").ToString().Trim();
                                        if(int.TryParse(jid, out jobID))
                                        {
                                            var jobStatus = (int)get_job_status(jobID, connString);
                                            if(jobStatus != -1)
                                            {
                                                // validate WO belongs to job
                                                xlCell = xlRow.Cells[cols.woNum];
                                                woNo = (xlCell.Value ?? "").ToString().Trim().ToUpper();
                                                isValid = valid_work_order(jobID, woNo, connString);
                                                if(isValid)
                                                {
                                                    // validate invAmt is numeric
                                                    xlCell = xlRow.Cells[cols.invAmt];
                                                    if(Single.TryParse((xlCell.Value ?? "").ToString().Trim(), out invAmt))
                                                    {
                                                        billData.InvoiceAmount = invAmt;

                                                        // validate vendor id
                                                        xlCell = xlRow.Cells[cols.vendorID];
                                                        string vid = (xlCell.Value ?? "").ToString().Trim();
                                                        if(vid != "0" && int.TryParse(vid, out vendorID))
                                                        {

                                                            // validate customer
                                                            string jobCust = (string)get_job_customer(jobID, connString);
                                                            if(jobCust != "")
                                                            {
                                                                billData.Customer = jobCust;

                                                                // validate expense acct is present
                                                                var expAcct = "";
                                                                xlCell = xlRow.Cells[cols.expAcct];
                                                                expAcct = (xlCell.Value ?? "").ToString().Trim();
                                                                if(expAcct == "0")
                                                                    expAcct = "";
                                                                if(expAcct != "")
                                                                {
                                                                    billData.ExpenseAcct = expAcct;

                                                                    // add invoice to aimm, get jobWO if successful
                                                                    if(addToAimm)
                                                                    {
                                                                        jobWO = add_invoice_to_job(jobID, woNo, vendorID, invNo, invAmt, invDesc, connString);
                                                                        isValid = (jobWO.Length != 0);
                                                                        if(isValid)
                                                                        {
                                                                            var msg = $"Added invoice {invNo} for vendor {vendor} to job {jobWO}";
                                                                            Status = msg;
                                                                            LogIt.LogInfo(msg);
                                                                        }
                                                                        else
                                                                        {
                                                                            (xlRow.Cells[cols.invNum]).Interior.ColorIndex = 3;
                                                                            LogIt.LogError($"Couldn't add invoice {invNo} for vendor {vendor} to job ID {jobWO}");
                                                                            set_excel_status(xlRow, "Error", "Couldn't add invoice to job");
                                                                        } // added aimm invoice

                                                                    }

                                                                    // add qb vendor bill if standard, else save for later
                                                                    if(addToQB)
                                                                    {
                                                                        if(importType == importTypes.standard)
                                                                        {
                                                                            qbValid = qb.AddStandardVendorBill(billData);
                                                                            if(qbValid)
                                                                            {
                                                                                var msg = $"Added invoice {invNo} for vendor {vendor} to QuickBooks";
                                                                                Status = msg;
                                                                                LogIt.LogInfo(msg);
                                                                            }
                                                                            else
                                                                            {
                                                                                xlRow.Cells[cols.invNum].Interior.ColorIndex = 3;
                                                                                var msg = $"Couldn't add invoice {invNo} for vendor {vendor} to QuickBooks: {billData.QBMessage}";
                                                                                Status = msg;
                                                                                LogIt.LogError(msg);
                                                                                set_excel_status(xlRow, billData.QBStatus, billData.QBMessage);
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            billList.Add(billData);
                                                                            qbValid = true;
                                                                        }
                                                                    }


                                                                }
                                                                else
                                                                {
                                                                    xlRow.Cells[cols.invNum].Interior.ColorIndex = 3;
                                                                    LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" is missing expense account");
                                                                    isValid = false;
                                                                } // valid expense account

                                                            }
                                                            else
                                                            {
                                                                (xlRow.Cells[cols.invNum]).Interior.ColorIndex = 3;
                                                                LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has missing customer");
                                                                isValid = false;
                                                            } // valid customer

                                                        }
                                                        else
                                                        {
                                                            isValid = false;
                                                            xlCell.Interior.ColorIndex = 3;
                                                            LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad vendor ID: {(xlCell.Value ?? "")}");
                                                            set_excel_status(xlRow, "Error", "Bad vendor ID");
                                                        } // vendor ID is numeric

                                                    }

                                                    else
                                                    {
                                                        isValid = false;
                                                        xlCell.Interior.ColorIndex = 3;
                                                        LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad invoice amount: {(xlCell.Value ?? "")}");
                                                        set_excel_status(xlRow, "Error", "Invalid invoice amount");
                                                    } // inv amt is numeric

                                                }
                                                else
                                                {
                                                    xlCell.Interior.ColorIndex = 3;
                                                    LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has invalid work order number: {woNo}");
                                                    set_excel_status(xlRow, "Error", "Invalid work order number");
                                                } // WO belongs to job

                                            }
                                            else
                                            {
                                                isValid = false;
                                                xlCell.Interior.ColorIndex = 3;
                                                LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\", job ID \"{jobID}\" is closed, cancelled or missing from database");
                                                set_excel_status(xlRow, "Error", "Job closed, cancelled or not found in database");
                                            } // valid job status
                                        }
                                        else
                                        {
                                            isValid = false;
                                            xlCell.Interior.ColorIndex = 3;
                                            LogIt.LogError($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad JobID: {(xlCell.Value ?? "").ToString()}");
                                            set_excel_status(xlRow, "Error", "Bad job ID");
                                        } // valid job id

                                    }
                                    else
                                    {
                                        isValid = false;
                                        var msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad date: {invDate}";
                                        xlCell.Interior.ColorIndex = 3;
                                        set_excel_status(xlRow, "Error", "Bad invoice date");
                                        LogIt.LogError(msg);
                                    } // valid date

                                }
                                else
                                {
                                    isValid = false;
                                    var msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" is missing full name";
                                    xlRow.Cells[cols.vendor].Interior.ColorIndex = 3;
                                    set_excel_status(xlRow, "Error", "Missing vendor full name");
                                    LogIt.LogError(msg);

                                } // has vendor full name
                            }
                            else
                            {
                                isValid = false;
                                var msg = $"Invoice number missing for vendor \"{vendor}\" in file \"{xlFile}\"";
                                xlCell.Interior.ColorIndex = 3;
                                set_excel_status(xlRow, "Error", "Missing invoice number");
                                LogIt.LogError(msg);
                            } // has invoice number

                            // keep track if all items were valid
                            allValid = allValid && isValid && qbValid;
                        }

                        // post all bills to quickbooks if stock invoices
                        if(addToQB && importType == importTypes.stock)
                            qbValid = qb.AddStockVendorBills(billList);

                        allValid = allValid && qbValid;

                        // save excel file if any invalid items.
                        var isOk = oXl.CloseWorkbook(!allValid);

                        // move workbook to archive/errors folder
                        destPath = allValid ? archivePath : errorPath;
                        destFile = string.Concat(
                            Path.GetFileNameWithoutExtension(xlFile),
                            DateTime.Now.ToString("_yyyy-MM-dd_HH-mm-ss"),
                            Path.GetExtension(xlFile));
                        destPathName = Path.Combine(destPath, destFile);
                        if(move_file(xlPathName, destPathName))
                        {
                            var msg = "";
                            txtExcelFile.Text = destPathName;
                            if(allValid)
                            {
                                msg = $"Moved \"{xlFile}\" to \"{destPathName}\"";
                                LogIt.LogInfo(msg);
                            }
                            else
                            {
                                msg = $"File \"{xlFile}\" had errors. Moved it to \"{destPathName}\"";
                                LogIt.LogWarn(msg);
                            }
                            Status = msg;
                        }
                    }
                    catch(Exception ex)
                    {
                        var msg = $"Error processing Excel file \"{xlFile}\": {ex.Message}";
                        LogIt.LogError(msg);
                        MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        Status = msg;
                    }

                }
                else
                {
                    var msg = "Could not connect to QuickBooks";
                    LogIt.LogError(msg);
                    MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    Status = msg;
                    qb = null;
                } // connected to quickbooks
                qb = null;

                oXl.CloseExcel();
                oXl = null;
            }
            else
            {
                var msg = $"Could not find Excel file \"{xlPathName}\"";
                LogIt.LogError(msg);
                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Status = msg;
            } // opened excel file

        }

        private void btnSplitPDF_Click(object sender, EventArgs e)
        {
            pdfPathName = txtPDFFile.Text;
            if(File.Exists(pdfPathName))
            {
                pdfFile = Path.GetFileName(pdfPathName);

                // make a sub-folder for today's date, use that for PDFs
                var subFolder = DateTime.Today.ToString("yyyy-MM-dd");
                pdfDestPath = Path.Combine(pdfDestPath, subFolder);

                // if dest folder exists delete contents, otherwise create
                if(Directory.Exists(pdfDestPath))
                {
                    var msg = $"Directory {pdfDestPath} already exists.\nOK to delete any PDF files?";
                    if(MessageBox.Show(msg, "Delete Files?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        return;

                    foreach(var file in Directory.GetFiles(pdfDestPath, "*.pdf", SearchOption.TopDirectoryOnly))
                    {
                        File.Delete(file);
                    }
                }
                else
                {
                    Directory.CreateDirectory(pdfDestPath);
                }

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
                            invoiceList.Add($"{vendor}_Invoice_{invNo}");
                        }
                        var isOk = close_excel();
                    }
                    catch(Exception ex)
                    {
                        var msg = $"Error processing Excel file \"{xlFile}\": {ex.Message}";
                        MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        LogIt.LogError(msg);
                        return;
                    }

                    var isOK = split_pdfs(pdfPathName, pdfDestPath, invoiceList);

                    // move original PDF to archive folder
                    if(isOK)
                    {
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
            Status = $"Opening Excel file \"{xlFile}\"";
            if(open_excel(xlPathName))
            {
                var billList = new List<BillData>();
                try
                {
                    allValid = true;

                    // loop thru each invoice row on worksheet
                    foreach(Excel.Range xlRow in xlRange.Rows)
                    {
                        isValid = true;
                        var billData = new BillData();
                        billData.APAccount = apAcct;
                        billData.ClassRef = billClass;

                        // get non-validated data for the invoice
                        xlCell = (Excel.Range)xlRow.Cells[cols.vendor];
                        string vendor = (xlCell.Value2 ?? "").ToString();

                        xlCell = (Excel.Range)xlRow.Cells[cols.invNum];
                        string invNo = (xlCell.Value2 ?? "").ToString();

                        // if blank items, we're done
                        if(vendor == "" || invNo == "")
                            break;

                        billData.InvoiceNumber = invNo;
                        Status = $"Processing invoice {invNo}";

                        xlCell = (Excel.Range)xlRow.Cells[cols.fullName];
                        billData.VendorFullName = (xlCell.Value2 ?? "").ToString();
                        if(billData.VendorFullName == "0")
                            billData.VendorFullName = "";

                        //xlCell = (Excel.Range)xlRow.Cells[cols.billFrom1];
                        //billData.BillFrom1 = (xlCell.Value2 ?? "").ToString();
                        //if(billData.BillFrom1 == "0")
                        //    billData.BillFrom1 = "";

                        //xlCell = (Excel.Range)xlRow.Cells[cols.billFrom2];
                        //billData.BillFrom2 = (xlCell.Value2 ?? "").ToString();
                        //if(billData.BillFrom2 == "0")
                        //    billData.BillFrom2 = "";

                        //xlCell = (Excel.Range)xlRow.Cells[cols.billFrom3];
                        //billData.BillFrom3 = (xlCell.Value2 ?? "").ToString();
                        //if(billData.BillFrom3 == "0")
                        //    billData.BillFrom3 = "";

                        //xlCell = (Excel.Range)xlRow.Cells[cols.billFrom4];
                        //billData.BillFrom4 = (xlCell.Value2 ?? "").ToString();
                        //if(billData.BillFrom4 == "0")
                        //    billData.BillFrom4 = "";

                        //xlCell = (Excel.Range)xlRow.Cells[cols.billFrom5];
                        //billData.BillFrom5 = (xlCell.Value2 ?? "").ToString();
                        //if(billData.BillFrom5 == "0")
                        //    billData.BillFrom5 = "";

                        // get and validate remaining items
                        DateTime invDate = new DateTime();
                        Single invAmt;
                        int jobID;
                        string expAcct;

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
                            billData.InvoiceDate = invDate;

                            // validate invAmt is numeric
                            xlCell = (Excel.Range)xlRow.Cells[cols.invAmt];
                            if(Single.TryParse(xlCell.Value2.ToString(), out invAmt))
                            {
                                billData.InvoiceAmount = invAmt;

                                // validate jobID is numeric
                                xlCell = (Excel.Range)xlRow.Cells[cols.jobID];
                                if(int.TryParse(xlCell.Value2.ToString(), out jobID))
                                {
                                    // get customer for job
                                    string jobCust = (string)get_job_customer(jobID, connString);
                                    if(jobCust == null)
                                    {
                                        xlCell = (Excel.Range)xlRow.Cells[cols.invNum];
                                        xlCell.Interior.ColorIndex = 44;
                                        LogIt.LogWarn($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has missing customer");
                                        billData.Customer = "";
                                        isValid = false;
                                    }
                                    else
                                    {
                                        billData.Customer = jobCust;
                                        Status = $"Got customer: {jobCust}";
                                    } // valid customer

                                    // validate expense acct is present
                                    xlCell = (Excel.Range)xlRow.Cells[cols.expAcct];
                                    expAcct = (xlCell.Value2 ?? "").ToString();
                                    if(expAcct == "0")
                                        expAcct = "";
                                    if(expAcct == "")
                                    {
                                        ((Excel.Range)xlRow.Cells[cols.invAmt]).Interior.ColorIndex = 44;
                                        LogIt.LogWarn($"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" is missing expense account");
                                        billData.ExpenseAcct = "";
                                        isValid = false;
                                    }
                                    else
                                    {
                                        billData.ExpenseAcct = expAcct;
                                    } // valid expense account

                                    //// get terms, calculate due date
                                    //xlCell = (Excel.Range)xlRow.Cells[cols.terms];
                                    //var terms = (xlCell.Value2 ?? "").ToString();
                                    //if(terms == "0")
                                    //    terms = "";
                                    //billData.Terms = terms;
                                    //billData.DueDate = get_due_date(vendor, invDate, terms);

                                    // add the QB invoice
                                    billList.Add(billData);

                                }
                                else
                                {
                                    isValid = false;
                                } // valid job id

                            }
                            else
                            {
                                var msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad invoice amount: {xlCell.Value2.ToString()}";
                                ((Excel.Range)xlRow.Cells[cols.invAmt]).Interior.ColorIndex = 3;
                                Status = msg;
                                LogIt.LogError(msg);
                                isValid = false;
                            } // inv amt is numeric


                        }
                        else
                        {
                            var msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad date: {xlCell.Value2.ToString()}";
                            ((Excel.Range)xlRow.Cells[cols.invDate]).Interior.ColorIndex = 3;
                            Status = msg;
                            LogIt.LogError(msg);
                            isValid = false;
                        } // valid date

                        allValid = allValid && isValid;
                        Status = $"Processed invoice # {invNo}";
                    }

                    // post the data to quickbooks
                    var qb = new QuickBooks.QuickBooks();
                    qb.StatusChanged += qb_StatusChanged;

                    var isOk = qb.AddStockVendorBills(billList);
                    if(isOk)
                    {
                        // update excel with results
                        bool anyErrors = update_excel(billList);
                        allValid = allValid && !anyErrors;
                    }
                    isOk = close_excel(!allValid);
                    qb = null;


                }
                catch(Exception ex)
                {
                    var msg = $"Error processing Excel file \"{xlFile}\": {ex.Message}";
                    MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    Status = msg;
                    LogIt.LogError(msg);
                }
            }
            else
            {
                var msg = $"Could not find Excel file \"{xlPathName}\"";
                Status = msg;
                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void qb_StatusChanged(object sender, StatusChangedEventArgs e)
        {
            Status = e.Status;
        }

        #endregion

        #region methods

        /// <summary>
        /// split large PDF file into individual page files
        /// </summary>
        /// <param name="bigPdfFile">full path-name of PDF file to split</param>
        /// <param name="destFolder">path to destination folder for split PDFs</param>
        /// <param name="invoiceList">list of file names to apply to split files</param>
        /// <returns>boolean indicating success</returns>
        private bool split_pdfs(string bigPdfFile, string destFolder, List<string> invoiceList)
        {
            var pdf = Path.GetFileName(bigPdfFile);
            var isOK = false;
            PdfCopy copy = null;
            try
            {
                using(PdfReader reader = new PdfReader(bigPdfFile))
                {
                    isOK = reader.NumberOfPages == invoiceList.Count;
                    if(isOK)
                    {
                        for(int pg = 0; pg < reader.NumberOfPages; pg++)
                        {
                            destFile = $"{invoiceList[pg]}.pdf";
                            destPathName = Path.Combine(destFolder, destFile);
                            using(Document document = new Document())
                            {
                                copy = new PdfCopy(document, new FileStream(destPathName, FileMode.Create));
                                document.Open();
                                copy.AddPage(copy.GetImportedPage(reader, pg + 1));
                                document.Close();
                            }
                        }
                    }
                    else
                    {
                        var msg = $"PDFs not split: Invoice list has {invoiceList.Count} pages, but PDF has {reader.NumberOfPages} pages.";
                        Status = msg;
                        LogIt.LogError(msg);
                    }
                }
            }
            catch(Exception ex)
            {
                isOK = false;
                var msg = $"Error processing PDF file \"{pdf}\": {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            finally
            {
                copy = null;
            }
            return isOK;
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
            bool result = false;
            if(File.Exists(fileName))
            {
                var xlFile = Path.GetFileName(fileName);
                try
                {
                    oXl = new clsExcel();
                    oXl.Visible = showExcel;
                    result = oXl.OpenExcel(xlPathName);//, "1");
                    if(result)
                        result = oXl.GetRange("invoices");
                    if(result)
                        LogIt.LogInfo($"Opened Excel file \"{xlFile}\"");
                    else
                        LogIt.LogError($"Could not open Excel file \"{xlFile}\"");
                }
                catch(Exception ex)
                {
                    var msg = $"Error opening Excel file \"{xlFile}\": {ex.Message}";
                    MessageBox.Show(msg, "Error", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                    LogIt.LogError(msg);
                    result = false;
                }
            }
            else
            {
                var msg = $"Could not find Excel file \"{xlFile}\"";
                LogIt.LogError(msg);
                result = false;
            }
            return result;
        }

        private void set_excel_status(dynamic row, string status, string message)
        {
            if(row != null)
            {
                row.Cells[cols.status].Value = status;
                string txt = (row.Cells[cols.message].Value ?? "").ToString().Trim();
                if(txt != "")
                {
                    row.Cells[cols.message] = $"{txt}\n{message}";
                    row.Cells[cols.message].Style.WrapText = true;
                    row.EntireRow.AutoFit();
                }
                else
                    row.Cells[cols.message].Value = message;
            }
        }

        /// <summary>
        /// update invoices in excel worksheet if errors
        /// </summary>
        /// <returns>boolean indicating whether changes were made to worksheet</returns>
        private bool update_excel(List<BillData> billList)
        {
            bool anyErrors = false;
            bool billError = false;
            try
            {
                foreach(BillData billData in billList)
                {
                    var row = billList.IndexOf(billData);
                    Excel.Range xlRow = xlRange.Rows[row + 1];
                    var status = billData.QBStatus;
                    var message = billData.QBMessage;
                    billError = false;
                    if(status == "Error")
                    {
                        xlCell = (Excel.Range)xlRow.Cells[cols.status];
                        xlCell.Value2 = status;
                        xlCell = (Excel.Range)xlRow.Cells[cols.message];
                        xlCell.Value2 = message;
                        ((Excel.Range)xlRow.Cells[cols.vendor]).Interior.ColorIndex = 3;
                        billError = true;
                    }
                    anyErrors = anyErrors || billError;
                }

                return anyErrors;
            }
            catch(Exception)
            {
                // don't want to save Excel if errors
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
                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                // close and release

                try
                {
                    oXl.CloseWorkbook(needToSave);
                    //xlWorkbook.Close(needToSave);

                }
                catch(COMException ex)
                {
                    // ignore if already closed
                }

                // release com objects to fully kill excel process from running in the background
                //try
                //{
                //    Marshal.ReleaseComObject(xlCell);
                //}
                //catch(NullReferenceException ex)
                //{
                //    // ignore if not yet instantiated
                //}
                //Marshal.ReleaseComObject(xlRange);
                //Marshal.ReleaseComObject(xlWorksheet);
                //Marshal.ReleaseComObject(xlWorkbook);

                //// quit and release
                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
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
            string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string sql = "INSERT INTO POL.tblProjectFinalMatEquip ( ProjectFinalID, JobID, BuildingMaterialID, "
                       + "OtherMaterial, CostEach, Quantity, TotalCost, Notes, EnteredDate, Correction, JobErrorID ) "
                       + $"VALUES ( '{projFinalID}','{jobID.ToString()}',{vendorID},'{invNo}',{invAmt},1,{invAmt},'{invDesc}','{now}',0,NULL );";

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
        /// get customer name for supplied job number
        /// </summary>
        /// <param name="jobID"></param>
        /// <param name="connectionString"></param>
        /// <returns>customer for job or null if not found</returns>
        private static string get_job_customer(int jobID, string connectionString)
        {
            string cust = "";
            SqlDataReader reader = null;
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    LogIt.LogInfo($"Getting customer for job {jobID}");
                    string cmdText = "SELECT TheCustomerSimple FROM MLG.dbo.vJobs where JobID = @jobID";
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
                                cust = (string)record[0];
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error getting customer for job {jobID}: {ex.Message}");
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
            return cust;
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

        /// <summary>
        /// calculate due date from supplied invoice date & terms
        /// </summary>
        /// <param name="vendor"></param>
        /// <param name="invoiceDate"></param>
        /// <param name="terms"></param>
        /// <returns></returns>
        private static DateTime get_due_date(string vendor, DateTime invoiceDate, string terms)
        {
            terms = terms.ToLower();
            DateTime dueDate = new DateTime();
            DateTime nextMo = invoiceDate.AddMonths(1);
            int days = 30;

            // no terms, use default 30 days
            if(terms == "")
            {
                dueDate = invoiceDate.AddDays(days);
                LogIt.LogInfo($"Vendor \"{vendor}\" has no terms supplied");
            }

            // get number of days from terms (ex: Net 30 days)
            else if(terms.EndsWith("days"))
            {
                terms = terms.Replace("net", "").Replace("days", "").Trim();
                try
                {
                    days = int.Parse(terms);
                }
                catch
                {
                    // do nothing, we already have a default days value
                }
                finally
                {
                    dueDate = invoiceDate.AddDays(days);
                }
            }

            // get particular day of month from terms (ex: Net 10th)
            else if(terms.EndsWith("th"))
            {
                terms = terms.Replace("net", "").Replace("th", "").Trim();
                try
                {
                    days = int.Parse(terms);
                }
                catch
                {
                    // do nothing, we already have a default days value
                }
                finally
                {
                    dueDate = new DateTime(nextMo.Year, nextMo.Month, days);
                }
            }

            // get particular day of month from terms (ex: Due 10th of mo.)
            else if(terms.StartsWith("due "))
            {
                terms = terms.Replace("due", "").Replace("th of mo", "").Replace(".", "").Trim();
                try
                {
                    days = int.Parse(terms);
                }
                catch
                {
                    // do nothing, we already have a default days value
                }
                finally
                {
                    dueDate = new DateTime(nextMo.Year, nextMo.Month, days);
                }
            }

            // get number of days from terms (ex: Net 30)
            else if(terms.StartsWith("net"))
            {
                terms = terms.Replace("net", "").Trim();
                try
                {
                    days = int.Parse(terms);
                }
                catch
                {
                    // do nothing, we already have a default days value
                }
                finally
                {
                    dueDate = invoiceDate.AddDays(days);
                }
            }

            // if none of above terms types, default to 30 days
            else
            {
                dueDate = invoiceDate.AddDays(days);
                LogIt.LogWarn($"Vendor \"{vendor}\" has unrecognized terms \"{terms}\"");
            }
            return dueDate;
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


        #endregion

    }

}
