using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Interop.QBXMLRP2;
using Aimm.Logging;
using System.Runtime.InteropServices;

namespace QuickBooks
{
    /// <summary>
    /// for reporting status back to caller
    /// </summary>
    public class StatusChangedEventArgs : EventArgs
    {
        private string v;
        public StatusChangedEventArgs(string v)
        {
            this.Status = v;
        }
        public string Status { get; set; }
    }

    /// <summary>
    /// for transporting bill info and results to/from caller
    /// </summary>
    public class BillData
    {
        public string VendorFullName { get; set; }
        public string InvoiceNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public string Terms { get; set; }
        public DateTime DueDate { get; set; }
        public Single InvoiceAmount { get; set; }
        public string Customer { get; set; }
        public string BillFrom1 { get; set; }
        public string BillFrom2 { get; set; }
        public string BillFrom3 { get; set; }
        public string BillFrom4 { get; set; }
        public string BillFrom5 { get; set; }
        public string ExpenseAcct { get; set; }
        public string APAccount { get; set; }
        public string ClassRef { get; set; }
        public string QBStatus { get; set; }
        public string QBMessage { get; set; }
    }

    /// <summary>
    /// encapsulates QuickBooks integration functionality
    /// </summary>
    public class QuickBooks
    {

        bool sessionBegun = false;
        bool connectionOpen = false;
        RequestProcessor2 rp = null;
        XmlDocument reqDoc = null;
        XmlElement outer = null;
        XmlElement inner = null;
        string ticket = null;

        ~QuickBooks()
        {
            Disconnect();
            inner = null;
            outer = null;
            reqDoc = null;
            rp = null;
        }

        /// <summary>
        /// Opens QuickBooks connection and starts session
        /// </summary>
        /// <param name="qbFileName">Full path name of QuickBooks file, or "" to use currently open file</param>
        /// <returns>boolean indicating success</returns>
        public bool Connect(string qbFileName="")
        {
            try
            {
                // appID not required, send ""
                string appID = "";
                string appName = "AIMM";
                QBXMLRPConnectionType connType = QBXMLRPConnectionType.localQBD;
                QBFileMode fileMode = QBFileMode.qbFileOpenDoNotCare;

                OnStatusChanged(new StatusChangedEventArgs("Opening connection to QuickBooks"));
                rp = new RequestProcessor2();
                rp.OpenConnection2(appID,appName, connType);
                connectionOpen = true;
                ticket = rp.BeginSession(qbFileName, fileMode);
                sessionBegun = true;
                OnStatusChanged(new StatusChangedEventArgs("Connected to QuickBooks"));
                return true;
            }
            catch(Exception ex)
            {
                OnStatusChanged(new StatusChangedEventArgs($"Could not connect to QuickBooks: {ex.Message}"));
                return false;
            }
        }

        private void Disconnect()
        {
            if(sessionBegun)
            {
                try
                {
                    rp.EndSession(ticket);
                }
                catch(Exception)
                {
                }
                finally
                {
                    sessionBegun = false;
                }

            }

            if(connectionOpen)
            {
                try
                {
                    rp.CloseConnection();
                }
                catch(Exception)
                {
                }
                finally
                {
                    connectionOpen = false;
                }
            }
            rp = null;
        }

        /// <summary>
        /// for reporting status back to caller
        /// </summary>
        public event EventHandler<StatusChangedEventArgs> StatusChanged;
        protected virtual void OnStatusChanged(StatusChangedEventArgs e)
        {
            StatusChanged?.Invoke(this, e);
        }

        /// <summary>
        /// Validate customer is in QuickBooks
        /// </summary>
        /// <param name="billData"><see cref="BillData"/> object containing customer name</param>
        /// <returns>boolean indicating whether customer is in quickbooks</returns>
        public bool IsQbCustomer(BillData billData)
        {
            return valid_qb_customer(rp, ticket, billData);
        }

        /// <summary>
        /// setup new request document
        /// </summary>
        private bool BuildXmlDoc()
        {
            bool result = false;
            try
            {
                reqDoc = new XmlDocument();
                reqDoc.AppendChild(reqDoc.CreateXmlDeclaration("1.0", null, null));
                reqDoc.AppendChild(reqDoc.CreateProcessingInstruction("qbxml", "version=\"13.0\""));

                //Create the outer request envelope tag
                outer = reqDoc.CreateElement("QBXML");
                reqDoc.AppendChild(outer);

                //Create the inner request envelope & any needed attributes
                inner = reqDoc.CreateElement("QBXMLMsgsRq");
                outer.AppendChild(inner);
                inner.SetAttribute("onError", "continueOnError");

                result = true;
            }
            catch(Exception ex)
            {
                var msg = $"Error occurred in \"BuildXmlDoc\" adding vendor bill: {ex.Message}";
                OnStatusChanged(new StatusChangedEventArgs(msg));
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// setup envelope for request data
        /// </summary>
        /// <param name="billData">A <see cref="BillData"/> object containing info for the vendor bill</param>
        private bool PrepareXmlRequest(BillData billData)
        {
            bool result = false;

            try
            {
                // get terms and due date
                GetQbVendorInfo(rp, ticket, billData);
                GetQbVendorDueDate(rp, ticket, billData);

                // create BillAddRq aggregate
                XmlElement billAddRq = reqDoc.CreateElement("BillAddRq");
                inner.AppendChild(billAddRq);

                // create BillAdd aggregate and fill in field values for it
                XmlElement billAdd = reqDoc.CreateElement("BillAdd");
                billAddRq.AppendChild(billAdd);

                // add bill header info
                result = AddBillHeader(billData);

            }
            catch(Exception ex)
            {
                var msg = $"Error occurred in \"PrepareXmlRequest\" adding vendor bill: {ex.Message}";
                OnStatusChanged(new StatusChangedEventArgs(msg));
                LogIt.LogError(msg);
                result = false;
            }
            return result;
        }

        /// <summary>
        /// add bill header information to xml request
        /// </summary>
        /// <param name="billData">A <see cref="BillData"/> object containing info for the vendor bill</param>
        private bool AddBillHeader(BillData billData)
        {
            bool result = false;
            try
            {
                var billAdd = reqDoc.SelectSingleNode("//BillAdd") as XmlElement;
                if(billAdd != null)
                {
                    // create VendorRef aggregate and fill in field values for it
                    XmlElement VendorRef = reqDoc.CreateElement("VendorRef");
                    billAdd.AppendChild(VendorRef);
                    VendorRef.AppendChild(MakeSimpleElem(reqDoc, "FullName", billData.VendorFullName));

                    // create VendorAddress aggregate and fill in field values for it
                    XmlElement VendorAddress = reqDoc.CreateElement("VendorAddress");
                    billAdd.AppendChild(VendorAddress);
                    VendorAddress.AppendChild(MakeSimpleElem(reqDoc, "Addr1", billData.BillFrom1));
                    VendorAddress.AppendChild(MakeSimpleElem(reqDoc, "Addr2", billData.BillFrom2));
                    VendorAddress.AppendChild(MakeSimpleElem(reqDoc, "Addr3", billData.BillFrom3));
                    VendorAddress.AppendChild(MakeSimpleElem(reqDoc, "Addr4", billData.BillFrom4));
                    VendorAddress.AppendChild(MakeSimpleElem(reqDoc, "Addr5", billData.BillFrom5));

                    // create APAccountRef aggregate and fill in field values for it
                    XmlElement APAccountRef = reqDoc.CreateElement("APAccountRef");
                    billAdd.AppendChild(APAccountRef);
                    APAccountRef.AppendChild(MakeSimpleElem(reqDoc, "FullName", billData.APAccount));

                    // set field value for TxnDate
                    billAdd.AppendChild(MakeSimpleElem(reqDoc, "TxnDate", billData.InvoiceDate.ToString("yyyy-MM-dd")));

                    // set field value for DueDate
                    billAdd.AppendChild(MakeSimpleElem(reqDoc, "DueDate", billData.DueDate.ToString("yyyy-MM-dd")));

                    // set field value for RefNumber
                    billAdd.AppendChild(MakeSimpleElem(reqDoc, "RefNumber", billData.InvoiceNumber));

                    //Create TermsRef aggregate and fill in field values for it
                    XmlElement TermsRef = reqDoc.CreateElement("TermsRef");
                    billAdd.AppendChild(TermsRef);
                    TermsRef.AppendChild(MakeSimpleElem(reqDoc, "FullName", billData.Terms));

                    result = true;
                }
                else
                {
                    var msg = $"Could not add bill header to xml request";
                    OnStatusChanged(new StatusChangedEventArgs(msg));
                    LogIt.LogError(msg);
                }

            }
            catch(Exception ex)
            {
                var msg = $"Error occurred in \"AddBillHeader\" adding vendor bill: {ex.Message}";
                OnStatusChanged(new StatusChangedEventArgs(msg));
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// add a bill expense line to xml request
        /// </summary>
        /// <param name="billData">A <see cref="BillData"/> object containing info for the vendor bill</param>
        private bool AddBillExpenseLine(BillData billData)
        {
            bool result = false;
            try
            {
                var billAdd = reqDoc.SelectSingleNode("//BillAdd") as XmlElement;
                if(billAdd != null)
                {
                    // create ExpenseLineAdd aggregate and fill in field values for it
                    XmlElement expenseLineAdd = reqDoc.CreateElement("ExpenseLineAdd");
                    billAdd.AppendChild(expenseLineAdd);

                    XmlElement AccountRef = reqDoc.CreateElement("AccountRef");
                    expenseLineAdd.AppendChild(AccountRef);
                    AccountRef.AppendChild(MakeSimpleElem(reqDoc, "FullName", billData.ExpenseAcct));

                    expenseLineAdd.AppendChild(MakeSimpleElem(reqDoc, "Amount", billData.InvoiceAmount.ToString("#.00")));
                    expenseLineAdd.AppendChild(MakeSimpleElem(reqDoc, "Memo", billData.InvoiceNumber));

                    XmlElement CustomerRef = reqDoc.CreateElement("CustomerRef");
                    expenseLineAdd.AppendChild(CustomerRef);
                    CustomerRef.AppendChild(MakeSimpleElem(reqDoc, "FullName", billData.Customer));

                    XmlElement ClassRef = reqDoc.CreateElement("ClassRef");
                    expenseLineAdd.AppendChild(ClassRef);
                    ClassRef.AppendChild(MakeSimpleElem(reqDoc, "FullName", billData.ClassRef));

                    result = true;
                }
                else
                {
                    var msg = $"Could not add bill expense line to xml request";
                    OnStatusChanged(new StatusChangedEventArgs(msg));
                    LogIt.LogError(msg);
                }
            }
            catch(Exception ex)
            {
                var msg = $"Error occurred in \"AddBillExpenseLine\" adding vendor bill: {ex.Message}";
                OnStatusChanged(new StatusChangedEventArgs(msg));
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// add a single STANDARD vendor bill to quickbooks
        /// </summary>
        /// <param name="billData">A <see cref="BillData"/> object containing info for the vendor bill</param>
        /// <returns></returns>
        public bool AddStandardVendorBill(BillData billData)
        {
            bool result = false;

            // build or clear xml doc
            if(reqDoc == null)
                result = BuildXmlDoc();
            else
            {
                inner.IsEmpty = true;
                result = true;
            }

            if(result)
                result = PrepareXmlRequest(billData);
            if(result)
                result = AddBillExpenseLine(billData);

            if(result)
                result = AddVendorBill(billData);

            return result;
        }

        /// <summary>
        /// add multiple STOCK vendor bills to a single quickbooks vendor bill
        /// </summary>
        /// <param name="billList">A list of <see cref="BillData"/> objects containing info for the vendor bills</param>
        /// <returns></returns>
        /// <remarks>assumes all bills from same vendor</remarks>
        public bool AddStockVendorBills(List<BillData> billList)
        {
            bool result = false;

            if(billList.Count > 0)
            {
                // build or clear xml doc
                if(reqDoc == null)
                    result = BuildXmlDoc();
                else
                {
                    inner.IsEmpty = true;
                    result = true;
                }

                if(result)
                    result = PrepareXmlRequest(billList[0]);
                if(result)
                {
                    foreach(var billData in billList)
                    {
                        result = AddBillExpenseLine(billData);
                        if(!result)
                        {
                            var msg = $"Could not add bill expense line to xml request";
                            OnStatusChanged(new StatusChangedEventArgs(msg));
                            LogIt.LogError(msg);
                            break;
                        }
                    }
                }
                if(result)
                    result = AddVendorBill(billList[0]);
                if(!result)
                {
                    //xlRow.Cells[cols.invNum].Interior.ColorIndex = 3;
                    //msg = $"Couldn't add invoice {invNo} for vendor {vendor} to QuickBooks: {billData.QBMessage}";
                    //Status = msg;
                    //LogIt.LogError(msg);
                    //set_excel_status(xlRow, billData.QBStatus, billData.QBMessage);


                }
            }
            return result;
        }

        /// <summary>
        /// add a vendor bill to QuickBooks
        /// </summary>
        /// <param name="billData">A <see cref="BillData"/> object containing info for the vendor bill</param>
        /// <returns></returns>
        public bool AddVendorBill(BillData billData)
        {
            bool result = false;

            try
            {
                // build, submit bill, get response
                //OnStatusChanged(new StatusChangedEventArgs("Building bill request"));
                //BuildBillAddRq(reqDoc, inner, billData);
                OnStatusChanged(new StatusChangedEventArgs($"Submitting invoice {billData.InvoiceNumber} to QuickBooks"));
                string responseStr = rp.ProcessRequest(ticket, reqDoc.OuterXml);
                WalkBillAddRs(responseStr, billData);
                if(billData.QBStatus == "Error")
                {
                    var msg = $"Error submitting invoice {billData.InvoiceNumber} to QuickBooks: {billData.QBMessage}";
                    OnStatusChanged(new StatusChangedEventArgs(msg));
                    LogIt.LogError(msg);
                    result = false;

                }
                else
                {
                    var msg = $"Submitted invoice {billData.InvoiceNumber} to QuickBooks";
                    OnStatusChanged(new StatusChangedEventArgs(msg));
                    LogIt.LogInfo(msg);
                    result = true;
                }
            }
            catch(Exception ex)
            {
                var msg = $"Error occurred adding vendor bill: {ex.Message}";
                OnStatusChanged(new StatusChangedEventArgs(msg));
                LogIt.LogError(msg);
                result = false;
            }
            return result;
        }

        /// <summary>
        /// creates an XmlElement to add to document
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="tagName"></param>
        /// <param name="tagVal"></param>
        /// <returns></returns>
        private XmlElement MakeSimpleElem(XmlDocument doc, string tagName, string tagVal)
        {
            XmlElement elem = doc.CreateElement(tagName);
            elem.InnerText = tagVal ?? "";
            return elem;
        }

        /// <summary>
        /// evaluates response from quickbooks and returns status and error message
        /// </summary>
        /// <param name="response"></param>
        /// <param name="billData"></param>
        void WalkBillAddRs(string response, BillData billData)
        {
            //Parse the response XML string into an XmlDocument
            XmlDocument responseXmlDoc = new XmlDocument();
            responseXmlDoc.LoadXml(response);

            //Get the response for our request
            XmlNodeList BillAddRsList = responseXmlDoc.GetElementsByTagName("BillAddRs");
            if(BillAddRsList.Count == 1) //Should always be true since we only did one request in this sample
            {
                XmlNode responseNode = BillAddRsList.Item(0);

                // get the status code, info, and severity
                XmlAttributeCollection rsAttributes = responseNode.Attributes;
                string statusCode = rsAttributes.GetNamedItem("statusCode").Value;
                string statusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
                string statusMessage = rsAttributes.GetNamedItem("statusMessage").Value;

                billData.QBStatus = statusSeverity;
                billData.QBMessage = statusMessage;

                // this code block would be used if we needed to iterate results and get field data
                ////status code = 0 all OK, > 0 is error
                //if(Convert.ToInt32(statusCode) >= 0)
                //{
                //    XmlNodeList BillRetList = responseNode.SelectNodes("//BillRet");//XPath Query
                //    for(int i = 0; i < BillRetList.Count; i++)
                //    {
                //        XmlNode BillRet = BillRetList.Item(i);
                //        WalkBillRet(BillRet);
                //    }
                //}
                rsAttributes = null;
            }
            BillAddRsList = null;
            responseXmlDoc = null;
        }

        /// <summary>
        /// This is for pulling out individual fields from response if needed.
        /// We are not currently using this.
        /// </summary>
        /// <param name="BillRet"></param>
        void WalkBillRet(XmlNode BillRet)
        {
            if(BillRet == null)
                return;

            //Go through all the elements of BillRet
            //Get value of TxnID
            string TxnID = BillRet.SelectSingleNode("./TxnID").InnerText;
            //Get value of TimeCreated
            string TimeCreated = BillRet.SelectSingleNode("./TimeCreated").InnerText;
            //Get value of TimeModified
            string TimeModified = BillRet.SelectSingleNode("./TimeModified").InnerText;
            //Get value of EditSequence
            string EditSequence = BillRet.SelectSingleNode("./EditSequence").InnerText;
            //Get value of TxnNumber
            if(BillRet.SelectSingleNode("./TxnNumber") != null)
            {
                string TxnNumber = BillRet.SelectSingleNode("./TxnNumber").InnerText;
            }
            //Get all field values for VendorRef aggregate
            //Get value of ListID
            if(BillRet.SelectSingleNode("./VendorRef/ListID") != null)
            {
                string ListID = BillRet.SelectSingleNode("./VendorRef/ListID").InnerText;
            }
            //Get value of FullName
            if(BillRet.SelectSingleNode("./VendorRef/FullName") != null)
            {
                string FullName = BillRet.SelectSingleNode("./VendorRef/FullName").InnerText;
            }
            //Done with field values for VendorRef aggregate

            //Get all field values for VendorAddress aggregate
            XmlNode VendorAddress = BillRet.SelectSingleNode("./VendorAddress");
            if(VendorAddress != null)
            {
                //Get value of Addr1
                if(BillRet.SelectSingleNode("./VendorAddress/Addr1") != null)
                {
                    string Addr1 = BillRet.SelectSingleNode("./VendorAddress/Addr1").InnerText;
                }
                //Get value of Addr2
                if(BillRet.SelectSingleNode("./VendorAddress/Addr2") != null)
                {
                    string Addr2 = BillRet.SelectSingleNode("./VendorAddress/Addr2").InnerText;
                }
                //Get value of Addr3
                if(BillRet.SelectSingleNode("./VendorAddress/Addr3") != null)
                {
                    string Addr3 = BillRet.SelectSingleNode("./VendorAddress/Addr3").InnerText;
                }
                //Get value of Addr4
                if(BillRet.SelectSingleNode("./VendorAddress/Addr4") != null)
                {
                    string Addr4 = BillRet.SelectSingleNode("./VendorAddress/Addr4").InnerText;
                }
                //Get value of Addr5
                if(BillRet.SelectSingleNode("./VendorAddress/Addr5") != null)
                {
                    string Addr5 = BillRet.SelectSingleNode("./VendorAddress/Addr5").InnerText;
                }
                //Get value of City
                if(BillRet.SelectSingleNode("./VendorAddress/City") != null)
                {
                    string City = BillRet.SelectSingleNode("./VendorAddress/City").InnerText;
                }
                //Get value of State
                if(BillRet.SelectSingleNode("./VendorAddress/State") != null)
                {
                    string State = BillRet.SelectSingleNode("./VendorAddress/State").InnerText;
                }
                //Get value of PostalCode
                if(BillRet.SelectSingleNode("./VendorAddress/PostalCode") != null)
                {
                    string PostalCode = BillRet.SelectSingleNode("./VendorAddress/PostalCode").InnerText;
                }
                //Get value of Country
                if(BillRet.SelectSingleNode("./VendorAddress/Country") != null)
                {
                    string Country = BillRet.SelectSingleNode("./VendorAddress/Country").InnerText;
                }
                //Get value of Note
                if(BillRet.SelectSingleNode("./VendorAddress/Note") != null)
                {
                    string Note = BillRet.SelectSingleNode("./VendorAddress/Note").InnerText;
                }
            }
            //Done with field values for VendorAddress aggregate

            //Get all field values for APAccountRef aggregate
            XmlNode APAccountRef = BillRet.SelectSingleNode("./APAccountRef");
            if(APAccountRef != null)
            {
                //Get value of ListID
                if(BillRet.SelectSingleNode("./APAccountRef/ListID") != null)
                {
                    string ListID = BillRet.SelectSingleNode("./APAccountRef/ListID").InnerText;
                }
                //Get value of FullName
                if(BillRet.SelectSingleNode("./APAccountRef/FullName") != null)
                {
                    string FullName = BillRet.SelectSingleNode("./APAccountRef/FullName").InnerText;
                }
            }
            //Done with field values for APAccountRef aggregate

            //Get value of TxnDate
            string TxnDate = BillRet.SelectSingleNode("./TxnDate").InnerText;
            //Get value of DueDate
            if(BillRet.SelectSingleNode("./DueDate") != null)
            {
                string DueDate = BillRet.SelectSingleNode("./DueDate").InnerText;
            }
            //Get value of AmountDue
            string AmountDue = BillRet.SelectSingleNode("./AmountDue").InnerText;
            //Get all field values for CurrencyRef aggregate
            XmlNode CurrencyRef = BillRet.SelectSingleNode("./CurrencyRef");
            if(CurrencyRef != null)
            {
                //Get value of ListID
                if(BillRet.SelectSingleNode("./CurrencyRef/ListID") != null)
                {
                    string ListID = BillRet.SelectSingleNode("./CurrencyRef/ListID").InnerText;
                }
                //Get value of FullName
                if(BillRet.SelectSingleNode("./CurrencyRef/FullName") != null)
                {
                    string FullName = BillRet.SelectSingleNode("./CurrencyRef/FullName").InnerText;
                }
            }
            //Done with field values for CurrencyRef aggregate

            //Get value of ExchangeRate
            if(BillRet.SelectSingleNode("./ExchangeRate") != null)
            {
                string ExchangeRate = BillRet.SelectSingleNode("./ExchangeRate").InnerText;
            }
            //Get value of AmountDueInHomeCurrency
            if(BillRet.SelectSingleNode("./AmountDueInHomeCurrency") != null)
            {
                string AmountDueInHomeCurrency = BillRet.SelectSingleNode("./AmountDueInHomeCurrency").InnerText;
            }
            //Get value of RefNumber
            if(BillRet.SelectSingleNode("./RefNumber") != null)
            {
                string RefNumber = BillRet.SelectSingleNode("./RefNumber").InnerText;
            }
            //Get all field values for TermsRef aggregate
            XmlNode TermsRef = BillRet.SelectSingleNode("./TermsRef");
            if(TermsRef != null)
            {
                //Get value of ListID
                if(BillRet.SelectSingleNode("./TermsRef/ListID") != null)
                {
                    string ListID = BillRet.SelectSingleNode("./TermsRef/ListID").InnerText;
                }
                //Get value of FullName
                if(BillRet.SelectSingleNode("./TermsRef/FullName") != null)
                {
                    string FullName = BillRet.SelectSingleNode("./TermsRef/FullName").InnerText;
                }
            }
            //Done with field values for TermsRef aggregate

            //Get value of Memo
            if(BillRet.SelectSingleNode("./Memo") != null)
            {
                string Memo = BillRet.SelectSingleNode("./Memo").InnerText;
            }
            //Get value of IsPaid
            if(BillRet.SelectSingleNode("./IsPaid") != null)
            {
                string IsPaid = BillRet.SelectSingleNode("./IsPaid").InnerText;
            }
            //Get value of ExternalGUID
            if(BillRet.SelectSingleNode("./ExternalGUID") != null)
            {
                string ExternalGUID = BillRet.SelectSingleNode("./ExternalGUID").InnerText;
            }
            //Walk list of LinkedTxn aggregates
            XmlNodeList LinkedTxnList = BillRet.SelectNodes("./LinkedTxn");
            if(LinkedTxnList != null)
            {
                for(int i = 0; i < LinkedTxnList.Count; i++)
                {
                    XmlNode LinkedTxn = LinkedTxnList.Item(i);
                    //Get value of TxnID
                    string TxnID2 = LinkedTxn.SelectSingleNode("./TxnID").InnerText;
                    //Get value of TxnType
                    string TxnType = LinkedTxn.SelectSingleNode("./TxnType").InnerText;
                    //Get value of TxnDate
                    string TxnDate2 = LinkedTxn.SelectSingleNode("./TxnDate").InnerText;
                    //Get value of RefNumber
                    if(LinkedTxn.SelectSingleNode("./RefNumber") != null)
                    {
                        string RefNumber = LinkedTxn.SelectSingleNode("./RefNumber").InnerText;
                    }
                    //Get value of LinkType
                    if(LinkedTxn.SelectSingleNode("./LinkType") != null)
                    {
                        string LinkType = LinkedTxn.SelectSingleNode("./LinkType").InnerText;
                    }
                    //Get value of Amount
                    string Amount = LinkedTxn.SelectSingleNode("./Amount").InnerText;
                }
            }

            //Walk list of ExpenseLineRet aggregates
            XmlNodeList ExpenseLineRetList = BillRet.SelectNodes("./ExpenseLineRet");
            if(ExpenseLineRetList != null)
            {
                for(int i = 0; i < ExpenseLineRetList.Count; i++)
                {
                    XmlNode ExpenseLineRet = ExpenseLineRetList.Item(i);
                    //Get value of TxnLineID
                    string TxnLineID = ExpenseLineRet.SelectSingleNode("./TxnLineID").InnerText;
                    //Get all field values for AccountRef aggregate
                    XmlNode AccountRef = ExpenseLineRet.SelectSingleNode("./AccountRef");
                    if(AccountRef != null)
                    {
                        //Get value of ListID
                        if(ExpenseLineRet.SelectSingleNode("./AccountRef/ListID") != null)
                        {
                            string ListID = ExpenseLineRet.SelectSingleNode("./AccountRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(ExpenseLineRet.SelectSingleNode("./AccountRef/FullName") != null)
                        {
                            string FullName = ExpenseLineRet.SelectSingleNode("./AccountRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for AccountRef aggregate

                    //Get value of Amount
                    if(ExpenseLineRet.SelectSingleNode("./Amount") != null)
                    {
                        string Amount = ExpenseLineRet.SelectSingleNode("./Amount").InnerText;
                    }
                    //Get value of Memo
                    if(ExpenseLineRet.SelectSingleNode("./Memo") != null)
                    {
                        string Memo = ExpenseLineRet.SelectSingleNode("./Memo").InnerText;
                    }
                    //Get all field values for CustomerRef aggregate
                    XmlNode CustomerRef = ExpenseLineRet.SelectSingleNode("./CustomerRef");
                    if(CustomerRef != null)
                    {
                        //Get value of ListID
                        if(ExpenseLineRet.SelectSingleNode("./CustomerRef/ListID") != null)
                        {
                            string ListID = ExpenseLineRet.SelectSingleNode("./CustomerRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(ExpenseLineRet.SelectSingleNode("./CustomerRef/FullName") != null)
                        {
                            string FullName = ExpenseLineRet.SelectSingleNode("./CustomerRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for CustomerRef aggregate

                    //Get all field values for ClassRef aggregate
                    XmlNode ClassRef = ExpenseLineRet.SelectSingleNode("./ClassRef");
                    if(ClassRef != null)
                    {
                        //Get value of ListID
                        if(ExpenseLineRet.SelectSingleNode("./ClassRef/ListID") != null)
                        {
                            string ListID = ExpenseLineRet.SelectSingleNode("./ClassRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(ExpenseLineRet.SelectSingleNode("./ClassRef/FullName") != null)
                        {
                            string FullName = ExpenseLineRet.SelectSingleNode("./ClassRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for ClassRef aggregate

                    //Get value of BillableStatus
                    if(ExpenseLineRet.SelectSingleNode("./BillableStatus") != null)
                    {
                        string BillableStatus = ExpenseLineRet.SelectSingleNode("./BillableStatus").InnerText;
                    }
                    //Get all field values for SalesRepRef aggregate
                    XmlNode SalesRepRef = ExpenseLineRet.SelectSingleNode("./SalesRepRef");
                    if(SalesRepRef != null)
                    {
                        //Get value of ListID
                        if(ExpenseLineRet.SelectSingleNode("./SalesRepRef/ListID") != null)
                        {
                            string ListID = ExpenseLineRet.SelectSingleNode("./SalesRepRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(ExpenseLineRet.SelectSingleNode("./SalesRepRef/FullName") != null)
                        {
                            string FullName = ExpenseLineRet.SelectSingleNode("./SalesRepRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for SalesRepRef aggregate

                    //Walk list of DataExtRet aggregates
                    XmlNodeList DataExtRetList = ExpenseLineRet.SelectNodes("./DataExtRet");
                    if(DataExtRetList != null)
                    {
                        for(int i1 = 0; i1 < DataExtRetList.Count; i1++)
                        {
                            XmlNode DataExtRet = DataExtRetList.Item(i1);
                            //Get value of OwnerID
                            if(DataExtRet.SelectSingleNode("./OwnerID") != null)
                            {
                                string OwnerID = DataExtRet.SelectSingleNode("./OwnerID").InnerText;
                            }
                            //Get value of DataExtName
                            string DataExtName = DataExtRet.SelectSingleNode("./DataExtName").InnerText;
                            //Get value of DataExtType
                            string DataExtType = DataExtRet.SelectSingleNode("./DataExtType").InnerText;
                            //Get value of DataExtValue
                            string DataExtValue = DataExtRet.SelectSingleNode("./DataExtValue").InnerText;
                        }
                    }

                }
            }

            XmlNodeList ORItemLineRetListChildren = BillRet.SelectNodes("./*");
            for(int i = 0; i < ORItemLineRetListChildren.Count; i++)
            {
                XmlNode Child = ORItemLineRetListChildren.Item(i);
                if(Child.Name == "ItemLineRet")
                {
                    //Get value of TxnLineID
                    string TxnLineID = Child.SelectSingleNode("./TxnLineID").InnerText;
                    //Get all field values for ItemRef aggregate
                    XmlNode ItemRef = Child.SelectSingleNode("./ItemRef");
                    if(ItemRef != null)
                    {
                        //Get value of ListID
                        if(Child.SelectSingleNode("./ItemRef/ListID") != null)
                        {
                            string ListID = Child.SelectSingleNode("./ItemRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(Child.SelectSingleNode("./ItemRef/FullName") != null)
                        {
                            string FullName = Child.SelectSingleNode("./ItemRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for ItemRef aggregate

                    //Get all field values for InventorySiteRef aggregate
                    XmlNode InventorySiteRef = Child.SelectSingleNode("./InventorySiteRef");
                    if(InventorySiteRef != null)
                    {
                        //Get value of ListID
                        if(Child.SelectSingleNode("./InventorySiteRef/ListID") != null)
                        {
                            string ListID = Child.SelectSingleNode("./InventorySiteRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(Child.SelectSingleNode("./InventorySiteRef/FullName") != null)
                        {
                            string FullName = Child.SelectSingleNode("./InventorySiteRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for InventorySiteRef aggregate

                    //Get all field values for InventorySiteLocationRef aggregate
                    XmlNode InventorySiteLocationRef = Child.SelectSingleNode("./InventorySiteLocationRef");
                    if(InventorySiteLocationRef != null)
                    {
                        //Get value of ListID
                        if(Child.SelectSingleNode("./InventorySiteLocationRef/ListID") != null)
                        {
                            string ListID = Child.SelectSingleNode("./InventorySiteLocationRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(Child.SelectSingleNode("./InventorySiteLocationRef/FullName") != null)
                        {
                            string FullName = Child.SelectSingleNode("./InventorySiteLocationRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for InventorySiteLocationRef aggregate

                    XmlNodeList ORSerialLotNumberChildren = Child.SelectNodes("./*");
                    for(int i1 = 0; i1 < ORSerialLotNumberChildren.Count; i1++)
                    {
                        XmlNode Child1 = ORSerialLotNumberChildren.Item(i1);
                        if(Child1.Name == "SerialNumber")
                        {
                        }

                        if(Child1.Name == "LotNumber")
                        {
                        }

                    }

                    //Get value of Desc
                    if(Child.SelectSingleNode("./Desc") != null)
                    {
                        string Desc = Child.SelectSingleNode("./Desc").InnerText;
                    }
                    //Get value of Quantity
                    if(Child.SelectSingleNode("./Quantity") != null)
                    {
                        string Quantity = Child.SelectSingleNode("./Quantity").InnerText;
                    }
                    //Get value of UnitOfMeasure
                    if(Child.SelectSingleNode("./UnitOfMeasure") != null)
                    {
                        string UnitOfMeasure = Child.SelectSingleNode("./UnitOfMeasure").InnerText;
                    }
                    //Get all field values for OverrideUOMSetRef aggregate
                    XmlNode OverrideUOMSetRef = Child.SelectSingleNode("./OverrideUOMSetRef");
                    if(OverrideUOMSetRef != null)
                    {
                        //Get value of ListID
                        if(Child.SelectSingleNode("./OverrideUOMSetRef/ListID") != null)
                        {
                            string ListID = Child.SelectSingleNode("./OverrideUOMSetRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(Child.SelectSingleNode("./OverrideUOMSetRef/FullName") != null)
                        {
                            string FullName = Child.SelectSingleNode("./OverrideUOMSetRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for OverrideUOMSetRef aggregate

                    //Get value of Cost
                    if(Child.SelectSingleNode("./Cost") != null)
                    {
                        string Cost = Child.SelectSingleNode("./Cost").InnerText;
                    }
                    //Get value of Amount
                    if(Child.SelectSingleNode("./Amount") != null)
                    {
                        string Amount = Child.SelectSingleNode("./Amount").InnerText;
                    }
                    //Get all field values for CustomerRef aggregate
                    XmlNode CustomerRef = Child.SelectSingleNode("./CustomerRef");
                    if(CustomerRef != null)
                    {
                        //Get value of ListID
                        if(Child.SelectSingleNode("./CustomerRef/ListID") != null)
                        {
                            string ListID = Child.SelectSingleNode("./CustomerRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(Child.SelectSingleNode("./CustomerRef/FullName") != null)
                        {
                            string FullName = Child.SelectSingleNode("./CustomerRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for CustomerRef aggregate

                    //Get all field values for ClassRef aggregate
                    XmlNode ClassRef = Child.SelectSingleNode("./ClassRef");
                    if(ClassRef != null)
                    {
                        //Get value of ListID
                        if(Child.SelectSingleNode("./ClassRef/ListID") != null)
                        {
                            string ListID = Child.SelectSingleNode("./ClassRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(Child.SelectSingleNode("./ClassRef/FullName") != null)
                        {
                            string FullName = Child.SelectSingleNode("./ClassRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for ClassRef aggregate

                    //Get value of BillableStatus
                    if(Child.SelectSingleNode("./BillableStatus") != null)
                    {
                        string BillableStatus = Child.SelectSingleNode("./BillableStatus").InnerText;
                    }
                    //Get all field values for SalesRepRef aggregate
                    XmlNode SalesRepRef = Child.SelectSingleNode("./SalesRepRef");
                    if(SalesRepRef != null)
                    {
                        //Get value of ListID
                        if(Child.SelectSingleNode("./SalesRepRef/ListID") != null)
                        {
                            string ListID = Child.SelectSingleNode("./SalesRepRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(Child.SelectSingleNode("./SalesRepRef/FullName") != null)
                        {
                            string FullName = Child.SelectSingleNode("./SalesRepRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for SalesRepRef aggregate

                    //Walk list of DataExtRet aggregates
                    XmlNodeList DataExtRetList2 = Child.SelectNodes("./DataExtRet");
                    if(DataExtRetList2 != null)
                    {
                        for(int i2 = 0; i2 < DataExtRetList2.Count; i2++)
                        {
                            XmlNode DataExtRet = DataExtRetList2.Item(i2);
                            //Get value of OwnerID
                            if(DataExtRet.SelectSingleNode("./OwnerID") != null)
                            {
                                string OwnerID = DataExtRet.SelectSingleNode("./OwnerID").InnerText;
                            }
                            //Get value of DataExtName
                            string DataExtName = DataExtRet.SelectSingleNode("./DataExtName").InnerText;
                            //Get value of DataExtType
                            string DataExtType = DataExtRet.SelectSingleNode("./DataExtType").InnerText;
                            //Get value of DataExtValue
                            string DataExtValue = DataExtRet.SelectSingleNode("./DataExtValue").InnerText;
                        }
                    }

                }

                if(Child.Name == "ItemGroupLineRet")
                {
                    //Get value of TxnLineID
                    string TxnLineID1 = Child.SelectSingleNode("./TxnLineID").InnerText;
                    //Get all field values for ItemGroupRef aggregate
                    //Get value of ListID
                    if(Child.SelectSingleNode("./ItemGroupRef/ListID") != null)
                    {
                        string ListID = Child.SelectSingleNode("./ItemGroupRef/ListID").InnerText;
                    }
                    //Get value of FullName
                    if(Child.SelectSingleNode("./ItemGroupRef/FullName") != null)
                    {
                        string FullName = Child.SelectSingleNode("./ItemGroupRef/FullName").InnerText;
                    }
                    //Done with field values for ItemGroupRef aggregate

                    //Get value of Desc
                    if(Child.SelectSingleNode("./Desc") != null)
                    {
                        string Desc = Child.SelectSingleNode("./Desc").InnerText;
                    }
                    //Get value of Quantity
                    if(Child.SelectSingleNode("./Quantity") != null)
                    {
                        string Quantity = Child.SelectSingleNode("./Quantity").InnerText;
                    }
                    //Get value of UnitOfMeasure
                    if(Child.SelectSingleNode("./UnitOfMeasure") != null)
                    {
                        string UnitOfMeasure = Child.SelectSingleNode("./UnitOfMeasure").InnerText;
                    }
                    //Get all field values for OverrideUOMSetRef aggregate
                    XmlNode OverrideUOMSetRef1 = Child.SelectSingleNode("./OverrideUOMSetRef");
                    if(OverrideUOMSetRef1 != null)
                    {
                        //Get value of ListID
                        if(Child.SelectSingleNode("./OverrideUOMSetRef/ListID") != null)
                        {
                            string ListID = Child.SelectSingleNode("./OverrideUOMSetRef/ListID").InnerText;
                        }
                        //Get value of FullName
                        if(Child.SelectSingleNode("./OverrideUOMSetRef/FullName") != null)
                        {
                            string FullName = Child.SelectSingleNode("./OverrideUOMSetRef/FullName").InnerText;
                        }
                    }
                    //Done with field values for OverrideUOMSetRef aggregate

                    //Get value of TotalAmount
                    string TotalAmount = Child.SelectSingleNode("./TotalAmount").InnerText;
                    //Walk list of ItemLineRet aggregates
                    XmlNodeList ItemLineRetList = Child.SelectNodes("./ItemLineRet");
                    if(ItemLineRetList != null)
                    {
                        for(int i3 = 0; i3 < ItemLineRetList.Count; i3++)
                        {
                            XmlNode ItemLineRet = ItemLineRetList.Item(i3);
                            //Get value of TxnLineID
                            string TxnLineID2 = ItemLineRet.SelectSingleNode("./TxnLineID").InnerText;
                            //Get all field values for ItemRef aggregate
                            XmlNode ItemRef = ItemLineRet.SelectSingleNode("./ItemRef");
                            if(ItemRef != null)
                            {
                                //Get value of ListID
                                if(ItemLineRet.SelectSingleNode("./ItemRef/ListID") != null)
                                {
                                    string ListID = ItemLineRet.SelectSingleNode("./ItemRef/ListID").InnerText;
                                }
                                //Get value of FullName
                                if(ItemLineRet.SelectSingleNode("./ItemRef/FullName") != null)
                                {
                                    string FullName = ItemLineRet.SelectSingleNode("./ItemRef/FullName").InnerText;
                                }
                            }
                            //Done with field values for ItemRef aggregate

                            //Get all field values for InventorySiteRef aggregate
                            XmlNode InventorySiteRef = ItemLineRet.SelectSingleNode("./InventorySiteRef");
                            if(InventorySiteRef != null)
                            {
                                //Get value of ListID
                                if(ItemLineRet.SelectSingleNode("./InventorySiteRef/ListID") != null)
                                {
                                    string ListID = ItemLineRet.SelectSingleNode("./InventorySiteRef/ListID").InnerText;
                                }
                                //Get value of FullName
                                if(ItemLineRet.SelectSingleNode("./InventorySiteRef/FullName") != null)
                                {
                                    string FullName = ItemLineRet.SelectSingleNode("./InventorySiteRef/FullName").InnerText;
                                }
                            }
                            //Done with field values for InventorySiteRef aggregate

                            //Get all field values for InventorySiteLocationRef aggregate
                            XmlNode InventorySiteLocationRef = ItemLineRet.SelectSingleNode("./InventorySiteLocationRef");
                            if(InventorySiteLocationRef != null)
                            {
                                //Get value of ListID
                                if(ItemLineRet.SelectSingleNode("./InventorySiteLocationRef/ListID") != null)
                                {
                                    string ListID = ItemLineRet.SelectSingleNode("./InventorySiteLocationRef/ListID").InnerText;
                                }
                                //Get value of FullName
                                if(ItemLineRet.SelectSingleNode("./InventorySiteLocationRef/FullName") != null)
                                {
                                    string FullName = ItemLineRet.SelectSingleNode("./InventorySiteLocationRef/FullName").InnerText;
                                }
                            }
                            //Done with field values for InventorySiteLocationRef aggregate

                            XmlNodeList ORSerialLotNumberChildren = ItemLineRet.SelectNodes("./*");
                            for(int i4 = 0; i4 < ORSerialLotNumberChildren.Count; i4++)
                            {
                                XmlNode Child2 = ORSerialLotNumberChildren.Item(i4);
                                if(Child2.Name == "SerialNumber")
                                {
                                }

                                if(Child2.Name == "LotNumber")
                                {
                                }

                            }

                            //Get value of Desc
                            if(ItemLineRet.SelectSingleNode("./Desc") != null)
                            {
                                string Desc = ItemLineRet.SelectSingleNode("./Desc").InnerText;
                            }
                            //Get value of Quantity
                            if(ItemLineRet.SelectSingleNode("./Quantity") != null)
                            {
                                string Quantity = ItemLineRet.SelectSingleNode("./Quantity").InnerText;
                            }
                            //Get value of UnitOfMeasure
                            if(ItemLineRet.SelectSingleNode("./UnitOfMeasure") != null)
                            {
                                string UnitOfMeasure = ItemLineRet.SelectSingleNode("./UnitOfMeasure").InnerText;
                            }
                            //Get all field values for OverrideUOMSetRef aggregate
                            XmlNode OverrideUOMSetRef = ItemLineRet.SelectSingleNode("./OverrideUOMSetRef");
                            if(OverrideUOMSetRef != null)
                            {
                                //Get value of ListID
                                if(ItemLineRet.SelectSingleNode("./OverrideUOMSetRef/ListID") != null)
                                {
                                    string ListID = ItemLineRet.SelectSingleNode("./OverrideUOMSetRef/ListID").InnerText;
                                }
                                //Get value of FullName
                                if(ItemLineRet.SelectSingleNode("./OverrideUOMSetRef/FullName") != null)
                                {
                                    string FullName = ItemLineRet.SelectSingleNode("./OverrideUOMSetRef/FullName").InnerText;
                                }
                            }
                            //Done with field values for OverrideUOMSetRef aggregate

                            //Get value of Cost
                            if(ItemLineRet.SelectSingleNode("./Cost") != null)
                            {
                                string Cost = ItemLineRet.SelectSingleNode("./Cost").InnerText;
                            }
                            //Get value of Amount
                            if(ItemLineRet.SelectSingleNode("./Amount") != null)
                            {
                                string Amount = ItemLineRet.SelectSingleNode("./Amount").InnerText;
                            }
                            //Get all field values for CustomerRef aggregate
                            XmlNode CustomerRef = ItemLineRet.SelectSingleNode("./CustomerRef");
                            if(CustomerRef != null)
                            {
                                //Get value of ListID
                                if(ItemLineRet.SelectSingleNode("./CustomerRef/ListID") != null)
                                {
                                    string ListID = ItemLineRet.SelectSingleNode("./CustomerRef/ListID").InnerText;
                                }
                                //Get value of FullName
                                if(ItemLineRet.SelectSingleNode("./CustomerRef/FullName") != null)
                                {
                                    string FullName = ItemLineRet.SelectSingleNode("./CustomerRef/FullName").InnerText;
                                }
                            }
                            //Done with field values for CustomerRef aggregate

                            //Get all field values for ClassRef aggregate
                            XmlNode ClassRef = ItemLineRet.SelectSingleNode("./ClassRef");
                            if(ClassRef != null)
                            {
                                //Get value of ListID
                                if(ItemLineRet.SelectSingleNode("./ClassRef/ListID") != null)
                                {
                                    string ListID = ItemLineRet.SelectSingleNode("./ClassRef/ListID").InnerText;
                                }
                                //Get value of FullName
                                if(ItemLineRet.SelectSingleNode("./ClassRef/FullName") != null)
                                {
                                    string FullName = ItemLineRet.SelectSingleNode("./ClassRef/FullName").InnerText;
                                }
                            }
                            //Done with field values for ClassRef aggregate

                            //Get value of BillableStatus
                            if(ItemLineRet.SelectSingleNode("./BillableStatus") != null)
                            {
                                string BillableStatus = ItemLineRet.SelectSingleNode("./BillableStatus").InnerText;
                            }
                            //Get all field values for SalesRepRef aggregate
                            XmlNode SalesRepRef = ItemLineRet.SelectSingleNode("./SalesRepRef");
                            if(SalesRepRef != null)
                            {
                                //Get value of ListID
                                if(ItemLineRet.SelectSingleNode("./SalesRepRef/ListID") != null)
                                {
                                    string ListID = ItemLineRet.SelectSingleNode("./SalesRepRef/ListID").InnerText;
                                }
                                //Get value of FullName
                                if(ItemLineRet.SelectSingleNode("./SalesRepRef/FullName") != null)
                                {
                                    string FullName = ItemLineRet.SelectSingleNode("./SalesRepRef/FullName").InnerText;
                                }
                            }
                            //Done with field values for SalesRepRef aggregate

                            //Walk list of DataExtRet aggregates
                            XmlNodeList DataExtRetList3 = ItemLineRet.SelectNodes("./DataExtRet");
                            if(DataExtRetList3 != null)
                            {
                                for(int i4 = 0; i4 < DataExtRetList3.Count; i4++)
                                {
                                    XmlNode DataExtRet = DataExtRetList3.Item(i4);
                                    //Get value of OwnerID
                                    if(DataExtRet.SelectSingleNode("./OwnerID") != null)
                                    {
                                        string OwnerID = DataExtRet.SelectSingleNode("./OwnerID").InnerText;
                                    }
                                    //Get value of DataExtName
                                    string DataExtName = DataExtRet.SelectSingleNode("./DataExtName").InnerText;
                                    //Get value of DataExtType
                                    string DataExtType = DataExtRet.SelectSingleNode("./DataExtType").InnerText;
                                    //Get value of DataExtValue
                                    string DataExtValue = DataExtRet.SelectSingleNode("./DataExtValue").InnerText;
                                }
                            }

                        }
                    }

                    //Walk list of DataExt aggregates
                    XmlNodeList DataExtList = Child.SelectNodes("./DataExt");
                    if(DataExtList != null)
                    {
                        for(int i5 = 0; i5 < DataExtList.Count; i5++)
                        {
                            XmlNode DataExt = DataExtList.Item(i5);
                            //Get value of OwnerID
                            string OwnerID = DataExt.SelectSingleNode("./OwnerID").InnerText;
                            //Get value of DataExtName
                            string DataExtName = DataExt.SelectSingleNode("./DataExtName").InnerText;
                            //Get value of DataExtValue
                            string DataExtValue = DataExt.SelectSingleNode("./DataExtValue").InnerText;
                        }
                    }

                }

            }

            //Get value of OpenAmount
            if(BillRet.SelectSingleNode("./OpenAmount") != null)
            {
                string OpenAmount = BillRet.SelectSingleNode("./OpenAmount").InnerText;
            }
            //Walk list of DataExtRet aggregates
            XmlNodeList DataExtRetList4 = BillRet.SelectNodes("./DataExtRet");
            if(DataExtRetList4 != null)
            {
                for(int i = 0; i < DataExtRetList4.Count; i++)
                {
                    XmlNode DataExtRet = DataExtRetList4.Item(i);
                    //Get value of OwnerID
                    if(DataExtRet.SelectSingleNode("./OwnerID") != null)
                    {
                        string OwnerID = DataExtRet.SelectSingleNode("./OwnerID").InnerText;
                    }
                    //Get value of DataExtName
                    string DataExtName = DataExtRet.SelectSingleNode("./DataExtName").InnerText;
                    //Get value of DataExtType
                    string DataExtType = DataExtRet.SelectSingleNode("./DataExtType").InnerText;
                    //Get value of DataExtValue
                    string DataExtValue = DataExtRet.SelectSingleNode("./DataExtValue").InnerText;
                }
            }

        }

        /// <summary>
        /// Lookup customer in QuickBooks
        /// </summary>
        /// <param name="rp">instantiated request processor with open session</param>
        /// <param name="ticket">existing session ticket</param>
        /// <param name="billData"><see cref="BillData"/> object containing customer name</param>
        /// <returns>boolean indicating whether customer is in quickbooks</returns>
        bool valid_qb_customer(RequestProcessor2 rp, string ticket, BillData billData)
        {
            XmlDocument doc = null;
            XmlElement docOuter = null;
            XmlElement docInner = null;
            XmlElement CustomerQueryRq = null;
            XmlDocument responseXmlDoc = null;
            XmlNodeList CustomerQueryRsList = null;
            XmlNode responseNode = null;
            XmlAttributeCollection rsAttributes = null;
            XmlNodeList CustomerRetList = null;
            XmlNode CustomerRet = null;
            bool isValid = false;

            try
            {
                // create doc and request envelope tags
                doc = new XmlDocument();
                doc.AppendChild(doc.CreateXmlDeclaration("1.0", null, null));
                doc.AppendChild(doc.CreateProcessingInstruction("qbxml", "version=\"13.0\""));

                docOuter = doc.CreateElement("QBXML");
                doc.AppendChild(docOuter);

                docInner = doc.CreateElement("QBXMLMsgsRq");
                docOuter.AppendChild(docInner);
                docInner.SetAttribute("onError", "continueOnError");

                CustomerQueryRq = doc.CreateElement("CustomerQueryRq");
                docInner.AppendChild(CustomerQueryRq);

                //Set field value for FullName
                CustomerQueryRq.AppendChild(MakeSimpleElem(doc, "FullName", billData.Customer));

                //Send the request and get the response from QuickBooks
                string responseStr = rp.ProcessRequest(ticket, doc.OuterXml);

                //Parse the response XML string into an XmlDocument
                responseXmlDoc = new XmlDocument();
                responseXmlDoc.LoadXml(responseStr);

                //Get the response for our request
                CustomerQueryRsList = responseXmlDoc.GetElementsByTagName("CustomerQueryRs");
                responseNode = CustomerQueryRsList.Item(0);

                //Check the status code, info, and severity
                rsAttributes = responseNode.Attributes;
                string statusCode = rsAttributes.GetNamedItem("statusCode").Value;
                string statusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
                string statusMessage = rsAttributes.GetNamedItem("statusMessage").Value;

                //status code = 0 all OK, > 0 is error
                if(Convert.ToInt32(statusCode) == 0)
                {
                    CustomerRetList = responseNode.SelectNodes("//CustomerRet");
                    if(CustomerRetList.Count > 0 && CustomerRetList.Item(0) != null)
                    {
                        // if we're here, the customer is valid.
                        // leave this in in case we later want to return customer data instead of bool.
                        CustomerRet = CustomerRetList.Item(0);
                        isValid = true;
                    }
                    else
                    {
                        var msg = $"Could not find customer \"{billData.Customer}\" in QuickBooks";
                        LogIt.LogError(msg);
                        billData.QBStatus = "Error";
                        billData.QBMessage = msg;
                    } // returned at least 1 valid customer
                }
                else
                {
                    LogIt.LogError($"Could not do customer lookup for \"{billData.Customer}\" in QuickBooks");
                    billData.QBStatus = statusSeverity;
                    billData.QBMessage = statusMessage;
                } // valid response status code
            }
            catch(Exception ex)
            {
                var msg = $"Error looking up customer \"{billData.Customer}\" in QuickBooks: {ex.Message}";
                LogIt.LogError(msg);
                billData.QBStatus = "Error";
                billData.QBMessage = msg;
            }
            finally
            {
                CustomerRet = null;
                CustomerRetList = null;
                rsAttributes = null;
                responseNode = null;
                CustomerQueryRsList = null;
                responseXmlDoc = null;
                CustomerQueryRq = null;
                docInner = null;
                docOuter = null;
                doc = null;
            }
            return isValid;
        }




        /// <summary>
        /// Lookup vendor in Quickbooks, get terms and address info
        /// </summary>
        /// <param name="rp">instantiated request processor with open session</param>
        /// <param name="ticket">existing session ticket</param>
        /// <param name="billData"><see cref="BillData"/> object containing vendor info</param>
        /// <returns></returns>
        void GetQbVendorInfo(RequestProcessor2 rp, string ticket, BillData billData)
        {
            XmlDocument doc = null;
            XmlElement docOuter = null;
            XmlElement docInner = null;
            XmlElement VendorQueryRq = null;
            XmlDocument responseXmlDoc = null;
            XmlNodeList VendorQueryRsList = null;
            XmlNode responseNode = null;
            XmlAttributeCollection rsAttributes = null;
            XmlNodeList VendorRetList = null;
            XmlNode VendorRet = null;
            XmlNode VendorAddressBlock = null;
            XmlNode TermsRef = null;

            try
            {
                // create doc and request envelope tags
                doc = new XmlDocument();
                doc.AppendChild(doc.CreateXmlDeclaration("1.0", null, null));
                doc.AppendChild(doc.CreateProcessingInstruction("qbxml", "version=\"13.0\""));

                docOuter = doc.CreateElement("QBXML");
                doc.AppendChild(docOuter);

                docInner = doc.CreateElement("QBXMLMsgsRq");
                docOuter.AppendChild(docInner);
                docInner.SetAttribute("onError", "continueOnError");

                VendorQueryRq = doc.CreateElement("VendorQueryRq");
                docInner.AppendChild(VendorQueryRq);

                //Set field value for FullName
                VendorQueryRq.AppendChild(MakeSimpleElem(doc, "FullName", billData.VendorFullName));

                //Send the request and get the response from QuickBooks
                string responseStr = rp.ProcessRequest(ticket, doc.OuterXml);

                //Parse the response XML string into an XmlDocument
                responseXmlDoc = new XmlDocument();
                responseXmlDoc.LoadXml(responseStr);

                //Get the response for our request
                VendorQueryRsList = responseXmlDoc.GetElementsByTagName("VendorQueryRs");
                responseNode = VendorQueryRsList.Item(0);

                //Check the status code, info, and severity
                rsAttributes = responseNode.Attributes;
                string statusCode = rsAttributes.GetNamedItem("statusCode").Value;
                string statusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
                string statusMessage = rsAttributes.GetNamedItem("statusMessage").Value;

                //status code = 0 all OK, > 0 is error
                if(Convert.ToInt32(statusCode) == 0)
                {
                    VendorRetList = responseNode.SelectNodes("//VendorRet");
                    if(VendorRetList.Count > 0 && VendorRetList.Item(0) != null)
                    {
                        VendorRet = VendorRetList.Item(0);
                        //Get all field values for VendorAddressBlock aggregate
                        VendorAddressBlock = VendorRet.SelectSingleNode("./VendorAddressBlock");
                        if(VendorAddressBlock != null)
                        {
                            //Get value of Addr1
                            if(VendorRet.SelectSingleNode("./VendorAddressBlock/Addr1") != null)
                            {
                                billData.BillFrom1 = VendorRet.SelectSingleNode("./VendorAddressBlock/Addr1").InnerText;
                            }
                            //Get value of Addr2
                            if(VendorRet.SelectSingleNode("./VendorAddressBlock/Addr2") != null)
                            {
                                billData.BillFrom2 = VendorRet.SelectSingleNode("./VendorAddressBlock/Addr2").InnerText;
                            }
                            //Get value of Addr3
                            if(VendorRet.SelectSingleNode("./VendorAddressBlock/Addr3") != null)
                            {
                                billData.BillFrom3 = VendorRet.SelectSingleNode("./VendorAddressBlock/Addr3").InnerText;
                            }
                            //Get value of Addr4
                            if(VendorRet.SelectSingleNode("./VendorAddressBlock/Addr4") != null)
                            {
                                billData.BillFrom4 = VendorRet.SelectSingleNode("./VendorAddressBlock/Addr4").InnerText;
                            }
                            //Get value of Addr5
                            if(VendorRet.SelectSingleNode("./VendorAddressBlock/Addr5") != null)
                            {
                                billData.BillFrom5 = VendorRet.SelectSingleNode("./VendorAddressBlock/Addr5").InnerText;
                            }
                        }

                        //Get all field values for TermsRef aggregate
                        TermsRef = VendorRet.SelectSingleNode("./TermsRef");
                        if(TermsRef != null)
                        {
                            //Get value of FullName
                            if(VendorRet.SelectSingleNode("./TermsRef/FullName") != null)
                            {
                                billData.Terms = VendorRet.SelectSingleNode("./TermsRef/FullName").InnerText;
                            }
                        }

                    }
                    else
                    {
                        var msg = $"Could not find vendor \"{billData.VendorFullName}\" in QuickBooks";
                        LogIt.LogError(msg);
                        billData.QBStatus = "Error";
                        billData.QBMessage = msg;
                    } // returned at least 1 valid vendor
                }
                else
                {
                    LogIt.LogError($"Could not do vendor lookup for \"{billData.VendorFullName}\" in QuickBooks");
                    billData.QBStatus = statusSeverity;
                    billData.QBMessage = statusMessage;
                } // valid response status code
            }
            catch(Exception ex)
            {
                var msg = $"Error looking up vendor \"{billData.VendorFullName}\" in QuickBooks: {ex.Message}";
                LogIt.LogError(msg);
                billData.QBStatus = "Error";
                billData.QBMessage = msg;
            }
            finally
            {
                TermsRef = null;
                VendorAddressBlock = null;
                VendorRet = null;
                VendorRetList = null;
                rsAttributes = null;
                responseNode = null;
                VendorQueryRsList = null;
                responseXmlDoc = null;
                VendorQueryRq = null;
                docInner = null;
                docOuter = null;
                doc = null;
            }
        }

        /// <summary>
        /// Lookup terms in Quickbooks, get due date
        /// </summary>
        /// <param name="rp">instantiated request processor with open session</param>
        /// <param name="ticket">existing session ticket</param>
        /// <param name="billData"><see cref="BillData"/> object containing vendor info</param>
        void GetQbVendorDueDate(RequestProcessor2 rp, string ticket, BillData billData)
        {
            XmlDocument doc = null;
            XmlElement docOuter = null;
            XmlElement docInner = null;
            XmlElement TermsQueryRq = null;
            XmlDocument responseXmlDoc = null;
            XmlNodeList TermsQueryRsList = null;
            XmlNode responseNode = null;
            XmlAttributeCollection rsAttributes = null;
            XmlNodeList ORList = null;
            XmlNode OR = null;
            XmlNode StandardTermsRet = null;
            XmlNode DateDrivenTermsRet = null;
            int days = 30;
            try
            {
                // create doc and request envelope tags
                doc = new XmlDocument();
                doc.AppendChild(doc.CreateXmlDeclaration("1.0", null, null));
                doc.AppendChild(doc.CreateProcessingInstruction("qbxml", "version=\"13.0\""));

                docOuter = doc.CreateElement("QBXML");
                doc.AppendChild(docOuter);

                docInner = doc.CreateElement("QBXMLMsgsRq");
                docOuter.AppendChild(docInner);
                docInner.SetAttribute("onError", "continueOnError");

                TermsQueryRq = doc.CreateElement("TermsQueryRq");
                docInner.AppendChild(TermsQueryRq);

                //Set field value for FullName
                TermsQueryRq.AppendChild(MakeSimpleElem(doc, "FullName", billData.Terms));

                //Send the request and get the response from QuickBooks
                string responseStr = rp.ProcessRequest(ticket, doc.OuterXml);

                //Parse the response XML string into an XmlDocument
                responseXmlDoc = new XmlDocument();
                responseXmlDoc.LoadXml(responseStr);

                //Get the response for our request
                TermsQueryRsList = responseXmlDoc.GetElementsByTagName("TermsQueryRs");
                responseNode = TermsQueryRsList.Item(0);

                //Check the status code, info, and severity
                rsAttributes = responseNode.Attributes;
                string statusCode = rsAttributes.GetNamedItem("statusCode").Value;
                string statusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
                string statusMessage = rsAttributes.GetNamedItem("statusMessage").Value;


                //status code = 0 all OK, > 0 is error
                if(Convert.ToInt32(statusCode) == 0)
                {

                    StandardTermsRet = responseNode.SelectSingleNode("./StandardTermsRet");
                    DateDrivenTermsRet = responseNode.SelectSingleNode("./DateDrivenTermsRet");

                    // standard
                    if(StandardTermsRet != null &&
                           StandardTermsRet.SelectSingleNode("./IsActive") != null &&
                           StandardTermsRet.SelectSingleNode("./IsActive").InnerText.ToLower() == "true")
                    {
                        // get value of StdDueDays, if invalid use default due days
                        string StdDueDays = "";
                        try
                        {
                            StdDueDays = StandardTermsRet.SelectSingleNode("./StdDueDays").InnerText;
                            days = int.Parse(StdDueDays);
                        }
                        catch
                        {
                            // do nothing, we already have a default days value
                            LogIt.LogWarn($"Couldn't parse StdDueDays \"{StdDueDays}\", using default due date");
                        }
                        finally
                        {
                            billData.DueDate = billData.InvoiceDate.AddDays(days);
                        }

                    }

                    // date-driven
                    else if(DateDrivenTermsRet != null &&
                            DateDrivenTermsRet.SelectSingleNode("./IsActive") != null &&
                            DateDrivenTermsRet.SelectSingleNode("./IsActive").InnerText.ToLower() == "true")
                    {
                        // get day of month due, if invalid use default due days
                        string DayOfMonthDue = "";
                        try
                        {
                            DayOfMonthDue = DateDrivenTermsRet.SelectSingleNode("./DayOfMonthDue").InnerText;
                            int dueDay = int.Parse(DayOfMonthDue);
                            DateTime nextMo = billData.InvoiceDate.AddMonths(1);
                            billData.DueDate = new DateTime(nextMo.Year, nextMo.Month, dueDay);

                        }
                        catch
                        {
                            // invalid day of month, so use standard net-30 terms
                            billData.DueDate = billData.InvoiceDate.AddDays(days);
                            LogIt.LogWarn($"Couldn't parse DayOfMonthDue \"{DayOfMonthDue}\", using default due date");
                        }

                    }

                    // if all else fails, use standard net-30 terms
                    else
                    {
                        billData.DueDate = billData.InvoiceDate.AddDays(days);
                    } // valid standard or date-driven terms

                    //}
                    //else
                    //{
                    //    var msg = $"Could not find terms \"{billData.Terms}\" in QuickBooks";
                    //    LogIt.LogError(msg);
                    //    billData.QBStatus = "Error";
                    //    billData.QBMessage = msg;
                    //} // returned at least 1 valid terms

                }
                else
                {
                    LogIt.LogError($"Could not do terms lookup for \"{billData.Terms}\" in QuickBooks");
                    billData.QBStatus = statusSeverity;
                    billData.QBMessage = statusMessage;
                } // valid response status code
            }

            catch(Exception ex)
            {
                var msg = $"Error looking up terms \"{billData.Terms}\" in QuickBooks: {ex.Message}";
                LogIt.LogError(msg);
                billData.QBStatus = "Error";
                billData.QBMessage = msg;
            }
            finally
            {
                DateDrivenTermsRet = null;
                StandardTermsRet = null;
                OR = null;
                ORList = null;
                rsAttributes = null;
                responseNode = null;
                TermsQueryRsList = null;
                responseXmlDoc = null;
                TermsQueryRq = null;
                docInner = null;
                docOuter = null;
                doc = null;
            }
        }
    }
}
