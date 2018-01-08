using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using Aimm.Logging;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using QuickBooks;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Drawing;
using System.Xml;
using System.Diagnostics;

namespace InvoiceImport
{
    public partial class frmImport : Form
    {
        // used to scroll text box if needed.
        //[DllImport("user32.dll", CharSet = CharSet.Auto)]
        //private static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, ref Point lParam);

        String[] args = Environment.GetCommandLineArgs();

        public frmImport()
        {
            InitializeComponent();
        }

        #region enums

        enum cols
        {
            vendor = 1,
            invType = 2,
            invNum = 3,
            pages = 4,
            invDate = 5,
            jobID = 6,
            woNum = 7,
            invAmt = 8,
            invDesc = 9,
            vendorID = 10,
            fullName = 11,
            expAcct = 12,
            status = 13,
            message = 14
        }

        enum importTypes
        {
            standard,
            stock
        }

        #endregion

        #region objects

        clsExcel oXl = null;
        dynamic xlCell = null;
        ToolTip toolTip1 = new ToolTip();
        QuickBooks.QuickBooks qb = null;
        #endregion

        #region variables

        // to scroll textbox
        //private int WM_VSCROLL= 277;
        //private System.IntPtr SB_BOTTOM = (IntPtr)7;
        //private Point pt = new Point();

        private bool isIDE = (Debugger.IsAttached == true);
        private string settingsPath;
        private string settingsFile;

        string connString;

        // for displaying log
        // private static string newLogLine = "";

        bool isValid = false;
        bool qbValid = false;
        bool allValid = true;

        string sourcePath = "";
        string archivePath = "";
        string errorPath = "";
        string logPath = "";
        string pdfDestPath = "";
        bool showExcel = true;
        string apAcct = "";
        string billClass = "";
        string quickBooksFileName = "";
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

        private string _status;
        public string Status
        {
            set
            {
                _status = value;
                txtStatus.Text = value.Replace("\n", "\r\n");
            }
            get { return _status; }
        }

        #endregion

        #region events

        private void frmImport_Load(object sender, EventArgs e)
        {
            // get settings
            try
            {
                if(isIDE)
                    settingsPath = Path.GetDirectoryName(Application.ExecutablePath);
                else
                    settingsPath = Path.GetDirectoryName(Application.CommonAppDataPath);
                settingsFile = Path.Combine(settingsPath, "Settings.xml");
                XmlDocument doc = new XmlDocument();
                doc.Load(settingsFile);
                connString = GetSetting(doc, "POLSQL");
                sourcePath = GetSetting(doc, "SourceFolder");
                archivePath = GetSetting(doc, "ArchiveFolder");
                errorPath = GetSetting(doc, "ErrorFolder");
                logPath = GetSetting(doc, "LogFolder");
                pdfDestPath = GetSetting(doc, "PdfFolder");
                apAcct = GetSetting(doc, "APAcct");
                billClass = GetSetting(doc, "BillClass");
                quickBooksFileName = GetSetting(doc, "QuickBooksFile");
                bool isOk = bool.TryParse(GetSetting(doc, "ShowExcel"), out showExcel);
                doc = null;
            }
            catch(Exception ex)
            {
                string msg = $"Could not read settings from \"{settingsFile}\": {ex.Message}";
                LogIt.LogError(msg);
                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Application.Exit();
            }

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
            xlPathName = txtExcelFile.Text;
            xlFile = Path.GetFileName(xlPathName);
            string msg = "";
            bool needValidJobId = false;
            bool needValidWO = false;
            bool needJobIdOrCustInJobId = false;

            bool addToAimmWo = ckAddToAIMM.Checked;
            bool addToQB = ckAddToQB.Checked;
            bool addToAimmFees = false;
            bool deferQbEntry = false;


            // fix so submit button is disabled unless user picks another workbook
            btnImportStandard.Enabled = false;

            if(addToAimmWo | addToQB)
            {
                if(xlPathName.Trim() != "")
                {
                    // continue if we can open excel file
                    if(open_excel(xlPathName))
                    {
                        // get and resize range
                        isValid = oXl.GetRange("invoices");
                        if(isValid)
                        {
                            msg = "Identified active invoices range";
                            Status = msg;
                            LogIt.LogInfo(msg);

                            // setup connection to quickbooks if needed, continue if connected
                            bool qbConnectedOrNotNeeded = false;
                            if(!addToQB)
                                qbConnectedOrNotNeeded = true;
                            else
                            {
                                qb = new QuickBooks.QuickBooks();
                                qb.StatusChanged += qb_StatusChanged;
                                qbConnectedOrNotNeeded = qb.Connect(quickBooksFileName);
                            }
                            if(qbConnectedOrNotNeeded)
                            {

                                var billList = new List<BillData>();
                                try
                                {
                                    allValid = true;

                                    // loop thru each invoice row on worksheet
                                    foreach(dynamic xlRow in oXl.Range.Rows)
                                    {
                                        isValid = false;
                                        var billData = new BillData();
                                        DateTime invDate = new DateTime();
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
                                            msg = $"Got invoice {invNo} for vendor {vendor}";
                                            Status = msg;
                                            LogIt.LogInfo(msg);

                                            // get the invoice type (as of 10/24/17, Vendor, Stock Out, Internal, Stock In/Svcs)
                                            // Vendor:          Goes into AIMM and QuickBooks; Requires valid JobID, Work Order; PDFs are split
                                            // Stock Out:       Goes into AIMM and QuickBooks (defer QB to end); Requires valid JobID, Work Order; NO PDFs split
                                            // Internal:        Goes into AIMM (as fee) and QuickBooks; Requires valid JobID, no Work Order; PDFs are split
                                            // Stock In/Svcs:   Goes into QuickBooks; Requires JobID or valid QB customer in JobID column; PDFs are split
                                            string invType = (xlRow.Cells[cols.invType].Value ?? "").ToString().Trim();
                                            switch(invType)
                                            {
                                                case "Vendor":
                                                    addToAimmWo = true;
                                                    addToAimmFees = false;
                                                    needValidJobId = true;
                                                    needValidWO = true;
                                                    needJobIdOrCustInJobId = false;
                                                    deferQbEntry = false;
                                                    break;
                                                case "Stock Out":
                                                    addToAimmWo = true;
                                                    addToAimmFees = false;
                                                    needValidJobId = true;
                                                    needValidWO = true;
                                                    needJobIdOrCustInJobId = false;
                                                    deferQbEntry = true;
                                                    break;
                                                case "Internal":
                                                    addToAimmWo = false;
                                                    addToAimmFees = true;
                                                    needValidJobId = true;
                                                    needValidWO = false;
                                                    needJobIdOrCustInJobId = false;
                                                    deferQbEntry = false;
                                                    break;
                                                case "Stock In/Svcs":
                                                    addToAimmWo = false;
                                                    addToAimmFees = false;
                                                    needValidJobId = false;
                                                    needValidWO = false;
                                                    needJobIdOrCustInJobId = true;
                                                    deferQbEntry = false;
                                                    break;

                                                default:
                                                    break;
                                            }

                                            // override above choices if checkbox is uncheccked
                                            addToAimmWo = addToAimmWo && ckAddToAIMM.Checked;
                                            addToAimmFees = addToAimmFees && ckAddToAIMM.Checked;

                                            string invDesc = (xlRow.Cells[cols.invDesc].Value ?? "").ToString().Trim();
                                            if(invDesc == "0")
                                                invDesc = "";

                                            // validate vendor full name has value
                                            xlCell = xlRow.Cells[cols.fullName];
                                            bool vendorFullNameValidOrNotNeeded = false;
                                            string vfn = "";
                                            if(!addToQB)
                                                vendorFullNameValidOrNotNeeded = true;
                                            else
                                            {
                                                vfn = (xlCell.Value ?? "").ToString().Trim();
                                                if(vfn == "0")
                                                    vfn = "";
                                                vendorFullNameValidOrNotNeeded = vfn != "";
                                            }
                                            if(vendorFullNameValidOrNotNeeded)
                                            {
                                                billData.VendorFullName = vfn;

                                                // validate invoice date is date
                                                // try both date number and string date conversion just in case
                                                xlCell = xlRow.Cells[cols.invDate];
                                                bool isValidDate = false;
                                                try
                                                {
                                                    // using value2 for date because it returns excel OA date
                                                    invDate = DateTime.FromOADate(xlCell.Value2);
                                                    isValidDate = true;
                                                }
                                                catch(Exception)
                                                {
                                                    isValidDate = DateTime.TryParse(xlCell.Value, out invDate);
                                                }

                                                if(isValidDate)
                                                {
                                                    billData.InvoiceDate = invDate;

                                                    // validate jobID is int
                                                    xlCell = xlRow.Cells[cols.jobID];
                                                    string jid = (xlCell.Value ?? "").ToString().Trim();

                                                    if(!needValidJobId || int.TryParse(jid, out jobID))
                                                    {
                                                        int jobStatus = 0;
                                                        bool isValidJobID = false;
                                                        if(needValidJobId)
                                                            jobStatus = (int)get_job_status(jobID, connString);
                                                        else
                                                        {
                                                            if(needJobIdOrCustInJobId && jid == "")
                                                                jobStatus = -1;
                                                        }
                                                        isValidJobID = (jobStatus != -1);
                                                        if(isValidJobID)
                                                        {

                                                            // validate WO belongs to job
                                                            xlCell = xlRow.Cells[cols.woNum];
                                                            woNo = (xlCell.Value ?? "").ToString().Trim().ToUpper();
                                                            if(!needValidWO || valid_work_order(jobID, woNo, connString))
                                                            {
                                                                // validate invAmt is numeric
                                                                xlCell = xlRow.Cells[cols.invAmt];
                                                                if(Single.TryParse((xlCell.Value ?? "").ToString().Trim(), out invAmt))
                                                                {
                                                                    billData.InvoiceAmount = invAmt;

                                                                    // validate vendor id
                                                                    bool vendorIdValidOrNotNeeded = false;
                                                                    if(!addToAimmWo)
                                                                        vendorIdValidOrNotNeeded = true;
                                                                    else
                                                                    {
                                                                        xlCell = xlRow.Cells[cols.vendorID];
                                                                        string vid = (xlCell.Value ?? "").ToString().Trim();
                                                                        vendorIdValidOrNotNeeded = vid != "0" && int.TryParse(vid, out vendorID);
                                                                    }
                                                                    if(vendorIdValidOrNotNeeded)
                                                                    {

                                                                        // if needJobIdOrCustInJobId and job id not numeric,
                                                                        // use it as customer. otherwise get customer from AIMM
                                                                        string jobCust = "";
                                                                        if(needJobIdOrCustInJobId && !int.TryParse(jid, out jobID))
                                                                            jobCust = jid;
                                                                        else
                                                                            jobCust = get_job_customer(jobID, connString);

                                                                        if(jobCust != "")
                                                                        {
                                                                            billData.Customer = jobCust;

                                                                            // validate customer in QB if needed
                                                                            bool isQbCustOrNotAddingToQb = true;
                                                                            if(addToQB)
                                                                                isQbCustOrNotAddingToQb = qb.IsQbCustomer(billData);

                                                                            if(isQbCustOrNotAddingToQb)
                                                                            {
                                                                                // validate expense acct is present if needed
                                                                                bool expAcctValidOrNotNeeded = false;
                                                                                string expAcct = "";
                                                                                if(!addToQB)
                                                                                    expAcctValidOrNotNeeded = true;
                                                                                else
                                                                                {
                                                                                    xlCell = xlRow.Cells[cols.expAcct];
                                                                                    expAcct = (xlCell.Value ?? "").ToString().Trim();
                                                                                    if(expAcct == "0")
                                                                                        expAcct = "";
                                                                                    expAcctValidOrNotNeeded = expAcct != "";
                                                                                }
                                                                                if(expAcctValidOrNotNeeded)
                                                                                {
                                                                                    billData.ExpenseAcct = expAcct;

                                                                                    // add invoice to aimm, get jobWO if successful
                                                                                    if(addToAimmWo)
                                                                                    {
                                                                                        jobWO = add_invoice_to_job(xlRow, jobID, woNo, vendorID, invNo, invAmt, invDesc, connString);
                                                                                        isValid = jobWO != "";
                                                                                    }
                                                                                    else if(addToAimmFees)
                                                                                    {
                                                                                        float totalInternalFees = add_invoice_to_internal_fees(xlRow, jobID, vendor, invNo, invAmt, invDesc, connString);
                                                                                        isValid = totalInternalFees != 0;
                                                                                        if(isValid)
                                                                                            isValid = update_job_internal_fees(xlRow, jobID, totalInternalFees, connString);
                                                                                    }
                                                                                    else
                                                                                        isValid = true;

                                                                                    // add qb vendor bill if needed
                                                                                    if(addToQB)
                                                                                    {
                                                                                        // if deferring QB entry, save detail for later. otherwise enter now
                                                                                        if(deferQbEntry)
                                                                                        {
                                                                                            billList.Add(billData);
                                                                                            qbValid = true;
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            qbValid = add_quickbooks_bill(xlRow, billData);
                                                                                        }
                                                                                    }
                                                                                    else
                                                                                        qbValid = true;
                                                                                }
                                                                                else
                                                                                {
                                                                                    xlRow.Cells[cols.invNum].Interior.ColorIndex = 3;
                                                                                    msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" is missing expense account";
                                                                                    LogIt.LogError(msg);
                                                                                    isValid = false;
                                                                                    set_excel_status(xlRow, "Error", "Missing expense account");
                                                                                } // valid expense account

                                                                            }
                                                                            else
                                                                            {
                                                                                isValid = false;
                                                                                xlRow.Cells[cols.invNum].Interior.ColorIndex = 3;
                                                                                msg = $"Error getting QuickBooks customer, invoice: {invNo}, vendor: {vendor}: {billData.QBMessage}";
                                                                                Status = msg;
                                                                                LogIt.LogError(msg);
                                                                                set_excel_status(xlRow, billData.QBStatus, billData.QBMessage);
                                                                            }

                                                                        }
                                                                        else
                                                                        {
                                                                            (xlRow.Cells[cols.invNum]).Interior.ColorIndex = 3;
                                                                            msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has missing customer";
                                                                            LogIt.LogError(msg);
                                                                            isValid = false;
                                                                            set_excel_status(xlRow, "Error", "Missing customer");
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
                                                            if(needValidJobId)
                                                                msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\", job ID \"{jobID}\" is missing from database";
                                                                // use this instead if we are skipping closed/cancelled jobs: msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\", job ID \"{jobID}\" is closed, cancelled or missing from database";
                                                            else if(needJobIdOrCustInJobId)
                                                                msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" is missing customer name";

                                                            LogIt.LogError(msg);

                                                            if(needValidJobId)
                                                                msg = "Job closed, cancelled or not found in database";
                                                            else if(needJobIdOrCustInJobId)
                                                                msg = "Invoice missing customer name";

                                                            set_excel_status(xlRow, "Error", msg);
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
                                                    msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" has bad date: {invDate}";
                                                    xlCell.Interior.ColorIndex = 3;
                                                    set_excel_status(xlRow, "Error", "Bad invoice date");
                                                    LogIt.LogError(msg);
                                                } // valid date

                                            }
                                            else
                                            {
                                                isValid = false;
                                                msg = $"Invoice {invNo} for vendor \"{vendor}\" in file \"{xlFile}\" is missing full name";
                                                xlRow.Cells[cols.vendor].Interior.ColorIndex = 3;
                                                set_excel_status(xlRow, "Error", "Missing vendor full name");
                                                LogIt.LogError(msg);

                                            } // has vendor full name
                                        }
                                        else
                                        {
                                            isValid = false;
                                            msg = $"Invoice number missing for vendor \"{vendor}\" in file \"{xlFile}\"";
                                            xlCell.Interior.ColorIndex = 3;
                                            set_excel_status(xlRow, "Error", "Missing invoice number");
                                            LogIt.LogError(msg);
                                        } // has invoice number

                                        // keep track if all items were valid
                                        allValid = allValid && isValid && qbValid;
                                    }

                                    // post any Stock Out invoices to quickbooks
                                    if(addToQB && billList.Count > 0)
                                    {
                                        Status = "Now importing all stock out vendor bills to QuickBooks. Please wait...";
                                        Cursor = Cursors.WaitCursor;
                                        qbValid = qb.AddStockVendorBills(billList);
                                        Cursor = Cursors.Default;
                                    }

                                    allValid = allValid && qbValid;

                                }
                                catch(Exception ex)
                                {
                                    msg = $"Error processing Excel file \"{xlFile}\", some invoices not imported: {ex.Message}";
                                    LogIt.LogError(msg);
                                    MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    Status = msg;
                                }

                            }
                            else
                            {
                                msg = "Could not connect to QuickBooks";
                                LogIt.LogError(msg);
                                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                Status = msg;
                                qb = null;
                            } // connected to quickbooks
                            qb = null;

                            // close excel file, save if any invalid items.
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
                                txtExcelFile.Text = destPathName;
                                if(allValid)
                                {
                                    msg = $"Import completed without errors. Moved \"{xlFile}\" to \"{destPathName}\"";
                                    LogIt.LogInfo(msg);
                                }
                                else
                                {
                                    msg = $"Import completed with errors. Moved \"{xlFile}\" to \"{destPathName}\"";
                                    LogIt.LogWarn(msg);
                                }
                                Status = msg;
                            }

                        }
                        else
                        {
                            msg = "Could not identify active invoices range, invoices not imported.";
                            Status = msg;
                            LogIt.LogError(msg);
                        } // got active invoices

                        oXl.CloseExcel();
                        oXl = null;
                    }
                    else
                    {
                        msg = $"Could not open Excel file \"{xlPathName}\", invoices not imported";
                        LogIt.LogError(msg);
                        Status = msg;
                    } // opened excel file

                }
                else
                {
                    Status = "You must enter an Excel file name to process.";
                }
            }
            else
            {
                Status = "You must choose \"Add Invoices to AIMM\" or \"Add Invoices to QuickBooks\".\nFile NOT processed.";
            }
        }

        private void btnSplitPDF_Click(object sender, EventArgs e)
        {
            pdfPathName = txtPDFFile.Text;
            string msg = "";
            if(File.Exists(pdfPathName))
            {
                pdfFile = Path.GetFileName(pdfPathName);

                // make a sub-folder for today's date, use that for PDFs
                var subFolder = DateTime.Today.ToString("yyyy-MM-dd");
                pdfDestPath = Path.Combine(txtPDFFolder.Text, subFolder);

                // if dest folder exists delete contents, otherwise create
                if(Directory.Exists(pdfDestPath))
                {
                    msg = $"Directory {pdfDestPath} already exists.\nOK to delete any PDF files?";
                    var result = MessageBox.Show(msg, "Delete Files?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3);
                    if(result == DialogResult.Cancel)
                        return;
                    else if(result == DialogResult.Yes)
                    {
                        foreach(var file in Directory.GetFiles(pdfDestPath, "*.pdf", SearchOption.TopDirectoryOnly))
                        {
                            File.Delete(file);
                        }
                    }
                }
                else
                {
                    Directory.CreateDirectory(pdfDestPath);
                }

                // get list of invoice numbers from xl document.
                xlPathName = txtExcelFile.Text;
                if(xlPathName.Trim() != "")
                {
                    var xlFile = Path.GetFileName(xlPathName);
                    if(open_excel(xlPathName))
                    {
                        // get and resize range
                        isValid = oXl.GetRange("invoices");
                        if(isValid)
                        {
                            msg = "Identified active invoices range";
                            Status = msg;
                            LogIt.LogInfo(msg);

                            var invoiceList = new List<KeyValuePair<string, int>>();
                            try
                            {
                                // loop thru each invoice row on worksheet, collect invoice numbers
                                foreach(dynamic xlRow in oXl.Range.Rows)
                                {
                                    string vendor = (xlRow.Cells[cols.vendor].Value ?? "").ToString().Trim();
                                    string invNo = (xlRow.Cells[cols.invNum].Value ?? "").ToString().Trim();
                                    if(vendor == "" || invNo == "")
                                        break;

                                    // get the invoice type (as of 10/24/17, Vendor, Stock Out, Internal, Stock In/Svcs)
                                    // Vendor:          Goes into AIMM and QuickBooks; Requires valid JobID, Work Order; PDFs are split
                                    // Stock Out:       Goes into AIMM and QuickBooks; Requires valid JobID, Work Order; NO PDFs split
                                    // Internal:        Goes into AIMM (as fee) and QuickBooks; Requires valid JobID, no Work Order; PDFs are split
                                    // Stock In/Svcs:   Goes into QuickBooks; Requires valid QB customer in JobID column; PDFs are split
                                    string invType = (xlRow.Cells[cols.invType].Value ?? "").ToString().Trim();
                                    if(invType != "Stock Out")
                                    {
                                        int pages = 1;
                                        var pgs = (xlRow.cells[cols.pages].Value ?? "1").ToString().Trim();
                                        int.TryParse(pgs, out pages);
                                        invoiceList.Add(new KeyValuePair<string, int>($"{vendor}_Invoice_{invNo}", pages));
                                    }
                                }
                                var isOk = oXl.CloseWorkbook();
                            }
                            catch(Exception ex)
                            {
                                msg = $"Error processing Excel file \"{xlFile}\": {ex.Message}";
                                Status = msg;
                                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                LogIt.LogError(msg);
                                return;
                            }

                            var isOK = split_pdfs(pdfPathName, pdfDestPath, invoiceList);

                            // move original PDF to archive folder
                            if(isOK)
                            {
                                msg = "Split PDFs";
                                Status = msg;
                                LogIt.LogInfo(msg);
                                destPath = archivePath;
                                destFile = string.Concat(
                                    Path.GetFileNameWithoutExtension(pdfFile),
                                    DateTime.Now.ToString("_yyyy-MM-dd_HH-mm-ss"),
                                    Path.GetExtension(pdfFile));
                                destPathName = Path.Combine(destPath, destFile);
                                if(move_file(pdfPathName, destPathName))
                                {
                                    txtPDFFile.Text = destPathName;
                                    msg = $"Moved \"{pdfFile}\" to \"{destPathName}\"";
                                    Status = msg;
                                    LogIt.LogInfo(msg);
                                }
                            }
                        }
                        else
                        {
                            msg = "Could not identify active invoices range, invoices not imported.";
                            Status = msg;
                            LogIt.LogError(msg);
                        } // got active invoices

                        oXl.CloseExcel();
                        oXl = null;
                    }
                    else
                    {
                        msg = $"Could not find or open Excel file \"{xlFile}\"";
                        Status = msg;
                        LogIt.LogWarn(msg);
                        MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }

                }
                else
                {
                    msg = $"Could not find PDF file \"{pdfPathName}\"";
                    Status = msg;
                    LogIt.LogWarn(msg);
                    MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                Status = "You must enter an Excel file name to process.";
            }

        }

        private void qb_StatusChanged(object sender, StatusChangedEventArgs e)
        {
            Status = e.Status;
        }

        #endregion

        #region methods

        /// <summary>
        /// return requested setting from xml settings file
        /// </summary>
        /// <param name="doc"><see cref="XmlDocument"/> containing settings</param>
        /// <param name="settingName">Name of setting to retrieve</param>
        /// <returns></returns>
        private string GetSetting(XmlDocument doc, string settingName)
        {
            string response = "";
            try
            {
                response = ((XmlElement)doc.SelectSingleNode($"/Settings/setting[@name='{settingName}']")).GetAttribute("value");
            }
            catch(Exception)
            {
            }
            return response;
        }

        /// <summary>
        /// split large PDF file into individual page files
        /// </summary>
        /// <param name="bigPdfFile">full path-name of PDF file to split</param>
        /// <param name="destFolder">path to destination folder for split PDFs</param>
        /// <param name="invoiceList">list of file names to apply to split files</param>
        /// <returns>boolean indicating success</returns>
        private bool split_pdfs(string bigPdfFile, string destFolder, List<KeyValuePair<string, int>> invoiceList)
        {
            var pdf = Path.GetFileName(bigPdfFile);
            var isOK = false;
            PdfCopy copy = null;
            try
            {
                using(PdfReader reader = new PdfReader(bigPdfFile))
                {
                    // make sure pdf has right number of pages
                    int xlPages = invoiceList.Sum(inv => inv.Value);
                    int pdfPages = reader.NumberOfPages;
                    isOK = xlPages == pdfPages;
                    if(isOK)
                    {
                        int pdfPage = 0;
                        for(int inv = 0; inv < invoiceList.Count(); inv++)
                        {
                            destFile = $"{invoiceList[inv].Key}.pdf";
                            int pages = invoiceList[inv].Value;
                            destPathName = Path.Combine(destFolder, destFile);
                            using(Document document = new Document())
                            {
                                copy = new PdfCopy(document, new FileStream(destPathName, FileMode.Create));
                                document.Open();

                                for(int pg = 0; pg < pages; pg++)
                                {
                                    pdfPage++;
                                    copy.AddPage(copy.GetImportedPage(reader, pdfPage));
                                }
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
            string msg = "";

            if(File.Exists(fileName))
            {
                var xlFile = Path.GetFileName(fileName);
                try
                {
                    oXl = new clsExcel();
                    oXl.Visible = showExcel;
                    result = oXl.OpenExcel(xlPathName);
                }
                catch(Exception ex)
                {
                    msg = $"Error opening Excel file \"{xlFile}\": {ex.Message}";
                    LogIt.LogError(msg);
                }
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
                row.Cells[cols.message].Style.WrapText = true;
                row.EntireRow.AutoFit();
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
                }
                catch(COMException)
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
        /// <param name="xlRow">Active Excel row being processed</param>
        /// <param name="jobID">Numeric Job ID</param>
        /// <param name="woNo">Work Order number</param>
        /// <param name="vendorID"></param>
        /// <param name="invNo"></param>
        /// <param name="invAmt"></param>
        /// <param name="invDesc"></param>
        /// <param name="connectionString"></param>
        /// <returns>string job# + work order # (12345-W01)</returns>
        private string add_invoice_to_job(dynamic xlRow, int jobID, string woNo, int vendorID, string invNo, Single invAmt, string invDesc, string connectionString)
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
                    string cmdText = $"INSERT INTO MLG.POL.tblInternalJobFees (JobID, FeeName, FeeAmount, Notes, CreateDate)";

                    using(SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        var msg = $"Added invoice {invNo} for vendor {vendorID} to AIMM work order \"{projFinalID}\"";
                        Status = msg;
                        LogIt.LogInfo(msg);
                        return projFinalID;
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error adding invoice for job {projFinalID}: {ex.Message}");
                (xlRow.Cells[cols.invNum]).Interior.ColorIndex = 3;
                LogIt.LogError($"Couldn't add invoice {invNo} for vendor {vendorID} to job ID {projFinalID}");
                set_excel_status(xlRow, "Error", "Couldn't add invoice to job");
                return string.Empty;
            }
        }

        /// <summary>
        /// add internal fee invoice to a job
        /// </summary>
        /// <param name="xlRow">Active Excel row being processed</param>
        /// <param name="jobID"></param>
        /// <param name="vendorName"></param>
        /// <param name="invNo"></param>
        /// <param name="invAmt"></param>
        /// <param name="invDesc"></param>
        /// <param name="connectionString"></param>
        /// <returns>Total amount of all internal fees for the job</returns>
        private float add_invoice_to_internal_fees(dynamic xlRow, int jobID, string vendorName, string invNo, Single invAmt, string invDesc, string connectionString)
        {
            float result = 0;
            string msg = "";
            string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string notes = $"Inv. {invNo}, {invDesc}";
            string cmdText = "";
            bool isOK = false;

            // first add internal fee record
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    cmdText = $"INSERT INTO MLG.POL.tblInternalJobFees (JobID, FeeName, FeeAmount, Notes, CreateDate) "
                            + $"VALUES ({jobID}, '{vendorName}', {invAmt}, '{notes}', '{now}')";

                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        conn.Open();
                        int rows = cmd.ExecuteNonQuery();
                        if(rows == 1)
                        {
                            isOK = true;
                            msg = $"Added internal fee to job ID {jobID}";
                            Status = msg;
                            LogIt.LogInfo(msg);
                        }
                        else
                        {
                            msg = $"Could not add internal fee for invoice {invNo} to job {jobID}";
                            LogIt.LogError(msg);
                            Status = msg;
                            set_excel_status(xlRow, "Error", "Couldn't add invoice to job");
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                result = 0;
                isOK = false;
                msg = $"Error adding invoice {invNo} to job {jobID}: {ex.Message}";
                LogIt.LogError(msg);
                Status = msg;
                set_excel_status(xlRow, "Error", "Couldn't add invoice to job");
            }


            // next get total of all internal fees for job
            if(isOK)
            {
                try
                {
                    using(SqlConnection conn = new SqlConnection(connectionString))
                    {
                        cmdText = $"SELECT CAST(SUM(FeeAmount) AS NUMERIC(9,2)) FROM MLG.POL.tblInternalJobFees WHERE JobID = {jobID}";
                        using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                        {
                            conn.Open();
                            isOK = float.TryParse(cmd.ExecuteScalar().ToString(), out result);
                            isOK = result != 0;
                        }
                    }
                    if(isOK)
                    {
                        msg = $"Got internal fees total for job ID {jobID}";
                        Status = msg;
                        LogIt.LogInfo(msg);
                    }
                }
                catch(Exception ex)
                {
                    result = 0;
                    isOK = false;
                    msg = $"Error getting internal fees total for job {jobID}: {ex.Message}";
                    LogIt.LogError(msg);
                    Status = msg;
                    set_excel_status(xlRow, "Error", "Couldn't add invoice to job");
                }
            }
            return result;
        }


        private bool update_job_internal_fees(dynamic xlRow, int jobID, Single jobInternalFeesTotal, string connectionString)
        {
            bool result = false;
            string msg = "";
            string cmdText = "";
            float jobTtlMinusLossAndNegMods = 0;

            // 1: get total price minus loss and neg mods from tbljobs
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    cmdText = $"SELECT TotalPriceMinusLossAndNegMods FROM MLG.POL.tblJobs WHERE JobID = {jobID}";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        conn.Open();
                        result = float.TryParse(cmd.ExecuteScalar().ToString(), out jobTtlMinusLossAndNegMods);
                    }
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting TotalPriceMinusLossAndNegMods for job {jobID}: {ex.Message}";
                LogIt.LogError(msg);
                Status = msg;
                set_excel_status(xlRow, "Error", "Couldn't add invoice to job");
                result = false;
            }

            // 2: update all mods for job if mod type not 3,4
            if(result)
            {
                string now = DateTime.Now.ToShortDateString();
                try
                {
                    using(SqlConnection conn = new SqlConnection(connectionString))
                    {
                        cmdText = $"SELECT * FROM MLG.POL.tblProjectEstimateCommon WHERE JobID = {jobID} AND ModuleTypeID NOT IN (3,4)";
                        using(SqlDataAdapter sda = new SqlDataAdapter(cmdText, conn))
                        {
                            using(DataTable dt = new DataTable())
                            {
                                sda.Fill(dt);
                                foreach(DataRow row in dt.Rows)
                                {
                                    // needed for calculations
                                    float estPrice = Convert.ToSingle(row["EstimatePrice"]);
                                    float ttlEquipMatCost = Convert.ToSingle(row["TotaEquipMaterialsCost"]);
                                    float ttlLaborCost = Convert.ToSingle(row["TotalLaborCost"]);
                                    float projEstBurden = Convert.ToSingle(row["ProjEstBurden"]);
                                    float modCommRate = Convert.ToSingle(row["ModCommRate"]);
                                    float manDays = Convert.ToSingle(row["ManDays"]);

                                    // fields to update
                                    float intFeesTtl;
                                    float gp;
                                    float salesCommis;
                                    float gpmd;


                                    intFeesTtl = estPrice / jobTtlMinusLossAndNegMods * jobInternalFeesTotal;
                                    gp = (estPrice - ttlEquipMatCost - ttlLaborCost - projEstBurden - intFeesTtl) / (1 + modCommRate);
                                    salesCommis = gp * modCommRate;
                                    gpmd = gp / manDays;

                                    row.BeginEdit();
                                    row["InternalFeestotal"] = intFeesTtl;
                                    row["InternalFeesNote"] = $"*** AUTO NOTE: INTERNAL FEES {intFeesTtl.ToString("$#,0.00")} APPLIED {now} ***";
                                    row["GP"] = gp;
                                    row["SalesCommission"] = salesCommis;
                                    row["GPMD"] = gp / manDays;
                                    row.EndEdit();
                                }
                                using(SqlCommandBuilder builder = new SqlCommandBuilder(sda))
                                {
                                    sda.UpdateCommand = builder.GetUpdateCommand();
                                    sda.Update(dt);
                                    dt.AcceptChanges();
                                }
                                msg = $"Updated {dt.Rows.Count} module(s) for job {jobID}";
                                Status = msg;
                                LogIt.LogInfo(msg);
                                result = true;
                            }
                        }
                    }
                }
                catch(Exception ex)
                {
                    msg = $"Error updating modules for job {jobID}: {ex.Message}";
                    LogIt.LogError(msg);
                    Status = msg;
                    set_excel_status(xlRow, "Error", "Couldn't add invoice to job");
                    result = false;
                }

                //try
                //{
                //    using(SqlConnection conn = new SqlConnection(connectionString))
                //    {
                //        //strSql = "SELECT * FROM tblProjectEstimateCommon WHERE JobID = '" & Me.txtJobID & "' AND ModuleTypeID not in (3,4)"
                //        //rs("InternalFeesTotal") = rs("EstimatePrice") / txtTotalPriceMinusLossAndNegMods * txtInternalFeesTotal
                //        //rs("InternalFeesNote") = "*** AUTO NOTE: INTERNAL FEES " & Format(rs("InternalFeesTotal"), "currency") & " APPLIED " & Format(Now, "m/d/yyyy") & " ***"
                //        //'*********** Re-Calc the GP, Com and GPMD!!!! ***********
                //        //rs("GP") = (rs("EstimatePrice") - rs("TotaEquipMaterialsCost") - rs("TotalLaborCost") - rs("ProjEstBurden") - rs("InternalFeesTotal")) / (1 + rs("ModCommRate"))
                //        //rs("SalesCommission") = rs("GP") * rs("ModCommRate")
                //        //rs("GPMD") = rs("GP") / rs("ManDays")

                //        cmdText = "UPDATE MLG.POL.tblProjectEstimateCommon SET "
                //        + $"InternalFeesTotal = EstimatePrice / {ttlMinusLossAndNegMods} * {internalFeesTotal}, "
                //        + $"InternalFeesNote = '*** AUTO NOTE: INTERNAL FEES {internalFeesTotal.ToString("$#,0.00")} APPLIED {DateTime.Now.ToShortDateString()} ***', "
                //        + $"GP = (EstimatePrice - TotaEquipMaterialsCost - TotalLaborCost - ProjEstBurden - (EstimatePrice / {ttlMinusLossAndNegMods} * {internalFeesTotal})) / (1 + ModCommRate), "
                //        + $"SalesCommission = ((EstimatePrice - TotaEquipMaterialsCost - TotalLaborCost - ProjEstBurden - (EstimatePrice / {ttlMinusLossAndNegMods} * {internalFeesTotal}))) * ModCommRate, "
                //        + $"GPMD = (EstimatePrice - TotaEquipMaterialsCost - TotalLaborCost - ProjEstBurden - (EstimatePrice / {ttlMinusLossAndNegMods} * {internalFeesTotal})) / (1 + ModCommRate) / ManDays "
                //        + $"WHERE JobID = {jobID} AND ModuleTypeID NOT IN (3,4)";

                //        using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                //        {
                //            conn.Open();
                //            int rows = cmd.ExecuteNonQuery();
                //            msg = $"Updated {rows} module(s) for job {jobID}";
                //            Status = msg;
                //            LogIt.LogInfo(msg);
                //            result = true;
                //        }
                //    }
                //}
                //catch(Exception ex)
                //{
                //    msg = $"Error updating modules for job {jobID}: {ex.Message}";
                //    LogIt.LogError(msg);
                //    Status = msg;
                //    set_excel_status(xlRow, "Error", "Couldn't add invoice to job");
                //    result = false;
                //}
            }

            // 3: update all subcontractor mods for job if mod type not 3,4
            if(result)
            {
                string now = DateTime.Now.ToShortDateString();
                try
                {
                    using(SqlConnection conn = new SqlConnection(connectionString))
                    {
                        cmdText = $"SELECT * FROM MLG.POL.tblSubModuleWOSubModule WHERE JobID = {jobID} AND ModuleTypeID NOT IN (3,4)";
                        using(SqlDataAdapter sda = new SqlDataAdapter(cmdText, conn))
                        {
                            using(DataTable dt = new DataTable())
                            {
                                sda.Fill(dt);

                                foreach(DataRow row in dt.Rows)
                                {
                                    // needed for calculations
                                    float estPrice = Convert.ToSingle(row["EstimatePrice"]);
                                    float subCost = Convert.ToSingle(row["SubCost"]);
                                    float matCost = Convert.ToSingle(row["MATCost"]);
                                    float subModCommRate = Convert.ToSingle(row["SubModCommRate"]);

                                    // fields to update
                                    float intFeesTtl;
                                    float gp;
                                    float salesCommis;

                                    intFeesTtl = estPrice / jobTtlMinusLossAndNegMods * jobInternalFeesTotal;
                                    float estGpPlusComm = estPrice - subCost - matCost - intFeesTtl;
                                    gp = estGpPlusComm / (1 + subModCommRate);
                                    salesCommis = estGpPlusComm - gp;

                                    row.BeginEdit();
                                    row["InternalFeestotal"] = intFeesTtl;
                                    row["InternalFeesNote"] = $"*** AUTO NOTE: INTERNAL FEES {intFeesTtl.ToString("$#,0.00")} APPLIED {now} ***";
                                    row["GP"] = gp;
                                    row["SalesCommission"] = salesCommis;
                                    row.EndEdit();
                                }
                                using(SqlCommandBuilder builder = new SqlCommandBuilder(sda))
                                {
                                    sda.UpdateCommand = builder.GetUpdateCommand();
                                    sda.Update(dt);
                                    dt.AcceptChanges();
                                }
                                msg = $"Updated {dt.Rows.Count} subcontractor module(s) for job {jobID}";
                                Status = msg;
                                LogIt.LogInfo(msg);
                                result = true;
                            }
                        }
                    }
                }
                catch(Exception ex)
                {
                    msg = $"Error updating subcontractor modules for job {jobID}: {ex.Message}";
                    LogIt.LogError(msg);
                    Status = msg;
                    set_excel_status(xlRow, "Error", "Couldn't add invoice to job");
                    result = false;
                }

                    //using(SqlConnection conn = new SqlConnection(connectionString))
                    //{
                    //    //strSql = "SELECT * FROM tblSubModuleWOSubModule WHERE JobID = '" & Me.txtJobID & "' AND ModuleTypeID not in (3,4)"  '<---- Not to NEg or Loss Mods!!!
                    //    //rs("InternalFeesTotal") = rs("EstimatePrice") / txtTotalPriceMinusLossAndNegMods * txtInternalFeesTotal
                    //    //rs("InternalFeesNote") = "*** AUTO NOTE: INTERNAL FEES " & Format(rs("InternalFeesTotal"), "currency") & " APPLIED " & Format(Now, "m/d/yyyy") & " ***"
                    //    //Est_GP_Plus_Comiss = rs("EstimatePrice") - rs("SubCost") - rs("MATCost") - rs("InternalFeesTotal")
                    //    //rs("GP") = Est_GP_Plus_Comiss / (1 + rs("SubModCommRate"))
                    //    //rs("SalesCommission") = Est_GP_Plus_Comiss - rs("GP")

                    //    cmdText = "UPDATE MLG.POL.tblSubModuleWOSubModule SET "
                    //    + $"InternalFeesTotal = EstimatePrice / {jobTtlMinusLossAndNegMods} * {internalFeesTotal}, "
                    //    + $"InternalFeesNote = '*** AUTO NOTE: INTERNAL FEES {internalFeesTotal.ToString("$#,0.00")} APPLIED {DateTime.Now.ToShortDateString()} ***', "
                    //    + $"GP = (EstimatePrice - SubCost - MATCost - (EstimatePrice / {jobTtlMinusLossAndNegMods} * {internalFeesTotal})) / (1 + SubModCommRate), "
                    //    + $"SalesCommission = (EstimatePrice - SubCost - MATCost - (EstimatePrice / {jobTtlMinusLossAndNegMods} * {internalFeesTotal})) "
                    //    + $"- ((EstimatePrice - SubCost - MATCost - (EstimatePrice / {jobTtlMinusLossAndNegMods} * {internalFeesTotal})) / (1 + SubModCommRate)) "
                    //    + $"WHERE JobID = {jobID} AND ModuleTypeID NOT IN (3,4)";

                    //    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    //    {
                    //        conn.Open();
                    //        int rows = cmd.ExecuteNonQuery();
                    //        msg = $"Updated {rows} subcontractor module(s) for job {jobID}";
                    //        Status = msg;
                    //        LogIt.LogInfo(msg);
                    //        result = true;
                    //    }
                    //}
            }

            // 4: update job totals and applied flag
            if(result)
            {
                try
                {
                    using(SqlConnection conn = new SqlConnection(connectionString))
                    {
                        // InternalFeesTotal
                        // InternalFeesApplied to true (-1?)
                        // SumInternalFeesTotalMods??
                        // SumInternalFeesTotalSubMods??
                        // SumInternalFeesTotalWOs ??
                        // SumInternalFeesTotalSubWO ??
                        cmdText = "UPDATE MLG.POL.tblJobs SET "
                        + $"InternalFeesTotal = {jobInternalFeesTotal}, SumInternalFeesTotalMods = {jobInternalFeesTotal}, SumInternalFeesTotalSubMods = {jobInternalFeesTotal}, InternalFeesApplied = -1 "
                        + $"WHERE JobID = {jobID}";

                        using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                        {
                            conn.Open();
                            int rows = cmd.ExecuteNonQuery();
                            if(rows == 1)
                            {
                                msg = $"Updated internal fees totals for job {jobID}";
                                LogIt.LogInfo(msg);
                                result = true;
                            }
                            else
                            {
                                msg = $"Could not update internal fees totals for job {jobID}";
                                LogIt.LogWarn(msg);
                                result = false;
                            }
                            Status = msg;
                        }
                    }
                }
                catch(Exception ex)
                {
                    msg = $"Error updating modules for job {jobID}: {ex.Message}";
                    LogIt.LogError(msg);
                    Status = msg;
                    set_excel_status(xlRow, "Error", "Couldn't add invoice to job");
                    result = false;
                }
            }
            return result;
        }

        /// <summary>
        /// Add a vendor bill to quickbooks
        /// </summary>
        /// <param name="xlRow">Active Excel row being processed</param>
        /// <param name="billData">A <see cref="BillData"/> object containing invoice information</param>
        /// <returns></returns>
        private bool add_quickbooks_bill(dynamic xlRow, BillData billData)
        {
            bool response = false;
            string msg = "";
            response = qb.AddStandardVendorBill(billData);
            if(response)
            {
                msg = $"Added invoice {billData.InvoiceNumber} for vendor {billData.VendorFullName} to QuickBooks";
                Status = msg;
                LogIt.LogInfo(msg);
            }
            else
            {
                xlRow.Cells[cols.invNum].Interior.ColorIndex = 3;
                msg = $"Couldn't add invoice {billData.InvoiceNumber} for vendor {billData.VendorFullName} to QuickBooks: {billData.QBMessage}";
                Status = msg;
                LogIt.LogError(msg);
                set_excel_status(xlRow, billData.QBStatus, billData.QBMessage);
            }
            return response;
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

                                // use this if any job status is ok to process
                                status = jobStatus;

                                // use this if we are considering closed or cancelled as invalid job statuses
                                //if(jobStatus != 7 & jobStatus != 11)
                                //{
                                //    status = jobStatus;
                                //}
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
