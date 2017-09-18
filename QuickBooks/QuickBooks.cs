using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Interop.QBFC13;

namespace QuickBooks
{
    public class QuickBooks
    {


        void BuildBillAddRq(IMsgSetRequest requestMsgSet)
        {
            
            /*
             * select vendor
             * enter date
             * due date should be pre-populated from vendor
             * enter ref # (invoice #)
             * enter amount (in sdk, enter this in details)
             * enter class (use PPI)
             * account should be pre-populated from vendor
             * choose customer job from drop-down
             * attach file
             * save
             */

            
            
            // following are all from sample code.
            // need to remove unnecessary ones.

            IBillAdd BillAddRq = requestMsgSet.AppendBillAddRq();
            
            //Set attributes
            
            //todo: get vendor list id, see if that will enter all fields?

            //Set field value for ListID
            BillAddRq.VendorRef.ListID.SetValue("200000-1011023419");
            
            //Set field value for FullName
            BillAddRq.VendorRef.FullName.SetValue("ab");
            //Set field value for Addr1
            BillAddRq.VendorAddress.Addr1.SetValue("ab");
            //Set field value for Addr2
            BillAddRq.VendorAddress.Addr2.SetValue("ab");
            //Set field value for Addr3
            BillAddRq.VendorAddress.Addr3.SetValue("ab");
            //Set field value for Addr4
            BillAddRq.VendorAddress.Addr4.SetValue("ab");
            //Set field value for Addr5
            BillAddRq.VendorAddress.Addr5.SetValue("ab");
            //Set field value for City
            BillAddRq.VendorAddress.City.SetValue("ab");
            //Set field value for State
            BillAddRq.VendorAddress.State.SetValue("ab");
            //Set field value for PostalCode
            BillAddRq.VendorAddress.PostalCode.SetValue("ab");
            //Set field value for Country
            BillAddRq.VendorAddress.Country.SetValue("ab");
            //Set field value for Note
            BillAddRq.VendorAddress.Note.SetValue("ab");
            
            
            //Set field value for ListID
            BillAddRq.APAccountRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            BillAddRq.APAccountRef.FullName.SetValue("ab");
            
            //Set field value for TxnDate
            BillAddRq.TxnDate.SetValue(DateTime.Parse("12/15/2007"));
            //Set field value for DueDate
            BillAddRq.DueDate.SetValue(DateTime.Parse("12/15/2007"));
            //Set field value for RefNumber
            BillAddRq.RefNumber.SetValue("ab");
            //Set field value for ListID
            BillAddRq.TermsRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            BillAddRq.TermsRef.FullName.SetValue("ab");
            //Set field value for Memo
            BillAddRq.Memo.SetValue("ab");
            //todo: fix or remove this item
            ////Set field value for ExchangeRate
            //BillAddRq.ExchangeRate.SetValue("IQBFloatType");
            //Set field value for ExternalGUID
            BillAddRq.ExternalGUID.SetValue(Guid.NewGuid().ToString());
            


            //Set field value for LinkToTxnIDList
            //May create more than one of these if needed
            BillAddRq.LinkToTxnIDList.Add("200000-1011023419");
            IExpenseLineAdd ExpenseLineAdd1 = BillAddRq.ExpenseLineAddList.Append();
            //Set attributes
            //Set field value for defMacro
            ExpenseLineAdd1.defMacro.SetValue("IQBStringType");
            //Set field value for ListID
            ExpenseLineAdd1.AccountRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ExpenseLineAdd1.AccountRef.FullName.SetValue("ab");
            
            //Set field value for Amount
            ExpenseLineAdd1.Amount.SetValue(10.01);
            //Set field value for Memo
            ExpenseLineAdd1.Memo.SetValue("ab");
            
            //Set field value for ListID
            ExpenseLineAdd1.CustomerRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ExpenseLineAdd1.CustomerRef.FullName.SetValue("ab");
            //Set field value for ListID
            ExpenseLineAdd1.ClassRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ExpenseLineAdd1.ClassRef.FullName.SetValue("ab");
            //Set field value for BillableStatus
            ExpenseLineAdd1.BillableStatus.SetValue(ENBillableStatus.bsBillable);
            //Set field value for ListID
            ExpenseLineAdd1.SalesRepRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ExpenseLineAdd1.SalesRepRef.FullName.SetValue("ab");
            IDataExt DataExt2 = ExpenseLineAdd1.DataExtList.Append();
            //Set field value for OwnerID
            DataExt2.OwnerID.SetValue(Guid.NewGuid().ToString());
            //Set field value for DataExtName
            DataExt2.DataExtName.SetValue("ab");
            //Set field value for DataExtValue
            DataExt2.DataExtValue.SetValue("ab");
            IORItemLineAdd ORItemLineAddListElement3 = BillAddRq.ORItemLineAddList.Append();
            string ORItemLineAddListElementType4 = "ItemLineAdd";
            if(ORItemLineAddListElementType4 == "ItemLineAdd")
            {
                //Set field value for ListID
                ORItemLineAddListElement3.ItemLineAdd.ItemRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORItemLineAddListElement3.ItemLineAdd.ItemRef.FullName.SetValue("ab");
                //Set field value for ListID
                ORItemLineAddListElement3.ItemLineAdd.InventorySiteRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORItemLineAddListElement3.ItemLineAdd.InventorySiteRef.FullName.SetValue("ab");
                //Set field value for ListID
                ORItemLineAddListElement3.ItemLineAdd.InventorySiteLocationRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORItemLineAddListElement3.ItemLineAdd.InventorySiteLocationRef.FullName.SetValue("ab");
                string ORSerialLotNumberElementType5 = "SerialNumber";
                if(ORSerialLotNumberElementType5 == "SerialNumber")
                {
                    //Set field value for SerialNumber
                    ORItemLineAddListElement3.ItemLineAdd.ORSerialLotNumber.SerialNumber.SetValue("ab");
                }
                if(ORSerialLotNumberElementType5 == "LotNumber")
                {
                    //Set field value for LotNumber
                    ORItemLineAddListElement3.ItemLineAdd.ORSerialLotNumber.LotNumber.SetValue("ab");
                }
                //Set field value for Desc
                ORItemLineAddListElement3.ItemLineAdd.Desc.SetValue("ab");
                //Set field value for Quantity
                ORItemLineAddListElement3.ItemLineAdd.Quantity.SetValue(2);
                //Set field value for UnitOfMeasure
                ORItemLineAddListElement3.ItemLineAdd.UnitOfMeasure.SetValue("ab");
                //Set field value for Cost
                ORItemLineAddListElement3.ItemLineAdd.Cost.SetValue(15.65);
                //Set field value for Amount
                ORItemLineAddListElement3.ItemLineAdd.Amount.SetValue(10.01);
                //Set field value for ListID
                ORItemLineAddListElement3.ItemLineAdd.CustomerRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORItemLineAddListElement3.ItemLineAdd.CustomerRef.FullName.SetValue("ab");
                //Set field value for ListID
                ORItemLineAddListElement3.ItemLineAdd.ClassRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORItemLineAddListElement3.ItemLineAdd.ClassRef.FullName.SetValue("ab");
                //Set field value for BillableStatus
                ORItemLineAddListElement3.ItemLineAdd.BillableStatus.SetValue(ENBillableStatus.bsBillable);
                //Set field value for ListID
                ORItemLineAddListElement3.ItemLineAdd.OverrideItemAccountRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORItemLineAddListElement3.ItemLineAdd.OverrideItemAccountRef.FullName.SetValue("ab");
                //Set field value for TxnID
                ORItemLineAddListElement3.ItemLineAdd.LinkToTxn.TxnID.SetValue("200000-1011023419");
                //Set field value for TxnLineID
                ORItemLineAddListElement3.ItemLineAdd.LinkToTxn.TxnLineID.SetValue("200000-1011023419");
                //Set field value for ListID
                ORItemLineAddListElement3.ItemLineAdd.SalesRepRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORItemLineAddListElement3.ItemLineAdd.SalesRepRef.FullName.SetValue("ab");
                IDataExt DataExt6 = ORItemLineAddListElement3.ItemLineAdd.DataExtList.Append();
                //Set field value for OwnerID
                DataExt6.OwnerID.SetValue(Guid.NewGuid().ToString());
                //Set field value for DataExtName
                DataExt6.DataExtName.SetValue("ab");
                //Set field value for DataExtValue
                DataExt6.DataExtValue.SetValue("ab");
            }
            if(ORItemLineAddListElementType4 == "ItemGroupLineAdd")
            {
                //Set field value for ListID
                ORItemLineAddListElement3.ItemGroupLineAdd.ItemGroupRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORItemLineAddListElement3.ItemGroupLineAdd.ItemGroupRef.FullName.SetValue("ab");
                //Set field value for Quantity
                ORItemLineAddListElement3.ItemGroupLineAdd.Quantity.SetValue(2);
                //Set field value for UnitOfMeasure
                ORItemLineAddListElement3.ItemGroupLineAdd.UnitOfMeasure.SetValue("ab");
                //Set field value for ListID
                ORItemLineAddListElement3.ItemGroupLineAdd.InventorySiteRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORItemLineAddListElement3.ItemGroupLineAdd.InventorySiteRef.FullName.SetValue("ab");
                //Set field value for ListID
                ORItemLineAddListElement3.ItemGroupLineAdd.InventorySiteLocationRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORItemLineAddListElement3.ItemGroupLineAdd.InventorySiteLocationRef.FullName.SetValue("ab");
                IDataExt DataExt7 = ORItemLineAddListElement3.ItemGroupLineAdd.DataExtList.Append();
                //Set field value for OwnerID
                DataExt7.OwnerID.SetValue(Guid.NewGuid().ToString());
                //Set field value for DataExtName
                DataExt7.DataExtName.SetValue("ab");
                //Set field value for DataExtValue
                DataExt7.DataExtValue.SetValue("ab");
            }
            
            //Set field value for IncludeRetElementList
            //May create more than one of these if needed
            BillAddRq.IncludeRetElementList.Add("ab");

        }

        void WalkBillAddRs(IMsgSetResponse responseMsgSet)
        {
            if(responseMsgSet == null)
                return;

            IResponseList responseList = responseMsgSet.ResponseList;
            if(responseList == null)
                return;

            //if we sent only one request, there is only one response, we'll walk the list for this sample
            for(int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                //check the status code of the response, 0=ok, >0 is warning
                if(response.StatusCode >= 0)
                {
                    //the request-specific response is in the details, make sure we have some
                    if(response.Detail != null)
                    {
                        //make sure the response is the type we're expecting
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if(responseType == ENResponseType.rtBillAddRs)
                        {
                            //upcast to more specific type here, this is safe because we checked with response.Type check above
                            IBillRet BillRet = (IBillRet)response.Detail;
                            WalkBillRet(BillRet);
                        }
                    }
                }
            }
        }


        void WalkBillRet(IBillRet BillRet)
        {
            if(BillRet == null)
                return;

            //Go through all the elements of IBillRet
            //Get value of TxnID
            string TxnID8 = (string)BillRet.TxnID.GetValue();
            //Get value of TimeCreated
            DateTime TimeCreated9 = (DateTime)BillRet.TimeCreated.GetValue();
            //Get value of TimeModified
            DateTime TimeModified10 = (DateTime)BillRet.TimeModified.GetValue();
            //Get value of EditSequence
            string EditSequence11 = (string)BillRet.EditSequence.GetValue();
            //Get value of TxnNumber
            if(BillRet.TxnNumber != null)
            {
                int TxnNumber12 = (int)BillRet.TxnNumber.GetValue();
            }
            //Get value of ListID
            if(BillRet.VendorRef.ListID != null)
            {
                string ListID13 = (string)BillRet.VendorRef.ListID.GetValue();
            }
            //Get value of FullName
            if(BillRet.VendorRef.FullName != null)
            {
                string FullName14 = (string)BillRet.VendorRef.FullName.GetValue();
            }
            if(BillRet.VendorAddress != null)
            {
                //Get value of Addr1
                if(BillRet.VendorAddress.Addr1 != null)
                {
                    string Addr115 = (string)BillRet.VendorAddress.Addr1.GetValue();
                }
                //Get value of Addr2
                if(BillRet.VendorAddress.Addr2 != null)
                {
                    string Addr216 = (string)BillRet.VendorAddress.Addr2.GetValue();
                }
                //Get value of Addr3
                if(BillRet.VendorAddress.Addr3 != null)
                {
                    string Addr317 = (string)BillRet.VendorAddress.Addr3.GetValue();
                }
                //Get value of Addr4
                if(BillRet.VendorAddress.Addr4 != null)
                {
                    string Addr418 = (string)BillRet.VendorAddress.Addr4.GetValue();
                }
                //Get value of Addr5
                if(BillRet.VendorAddress.Addr5 != null)
                {
                    string Addr519 = (string)BillRet.VendorAddress.Addr5.GetValue();
                }
                //Get value of City
                if(BillRet.VendorAddress.City != null)
                {
                    string City20 = (string)BillRet.VendorAddress.City.GetValue();
                }
                //Get value of State
                if(BillRet.VendorAddress.State != null)
                {
                    string State21 = (string)BillRet.VendorAddress.State.GetValue();
                }
                //Get value of PostalCode
                if(BillRet.VendorAddress.PostalCode != null)
                {
                    string PostalCode22 = (string)BillRet.VendorAddress.PostalCode.GetValue();
                }
                //Get value of Country
                if(BillRet.VendorAddress.Country != null)
                {
                    string Country23 = (string)BillRet.VendorAddress.Country.GetValue();
                }
                //Get value of Note
                if(BillRet.VendorAddress.Note != null)
                {
                    string Note24 = (string)BillRet.VendorAddress.Note.GetValue();
                }
            }
            if(BillRet.APAccountRef != null)
            {
                //Get value of ListID
                if(BillRet.APAccountRef.ListID != null)
                {
                    string ListID25 = (string)BillRet.APAccountRef.ListID.GetValue();
                }
                //Get value of FullName
                if(BillRet.APAccountRef.FullName != null)
                {
                    string FullName26 = (string)BillRet.APAccountRef.FullName.GetValue();
                }
            }
            //Get value of TxnDate
            DateTime TxnDate27 = (DateTime)BillRet.TxnDate.GetValue();
            //Get value of DueDate
            if(BillRet.DueDate != null)
            {
                DateTime DueDate28 = (DateTime)BillRet.DueDate.GetValue();
            }
            //Get value of AmountDue
            double AmountDue29 = (double)BillRet.AmountDue.GetValue();
            if(BillRet.CurrencyRef != null)
            {
                //Get value of ListID
                if(BillRet.CurrencyRef.ListID != null)
                {
                    string ListID30 = (string)BillRet.CurrencyRef.ListID.GetValue();
                }
                //Get value of FullName
                if(BillRet.CurrencyRef.FullName != null)
                {
                    string FullName31 = (string)BillRet.CurrencyRef.FullName.GetValue();
                }
            }
            
            //todo: fix or remove this item
            ////Get value of ExchangeRate
            //if(BillRet.ExchangeRate != null)
            //{
            //    IQBFloatType ExchangeRate32 = (IQBFloatType)BillRet.ExchangeRate.GetValue();
            //}
            //Get value of AmountDueInHomeCurrency
            if(BillRet.AmountDueInHomeCurrency != null)
            {
                double AmountDueInHomeCurrency33 = (double)BillRet.AmountDueInHomeCurrency.GetValue();
            }
            //Get value of RefNumber
            if(BillRet.RefNumber != null)
            {
                string RefNumber34 = (string)BillRet.RefNumber.GetValue();
            }
            if(BillRet.TermsRef != null)
            {
                //Get value of ListID
                if(BillRet.TermsRef.ListID != null)
                {
                    string ListID35 = (string)BillRet.TermsRef.ListID.GetValue();
                }
                //Get value of FullName
                if(BillRet.TermsRef.FullName != null)
                {
                    string FullName36 = (string)BillRet.TermsRef.FullName.GetValue();
                }
            }
            //Get value of Memo
            if(BillRet.Memo != null)
            {
                string Memo37 = (string)BillRet.Memo.GetValue();
            }
            //Get value of IsPaid
            if(BillRet.IsPaid != null)
            {
                bool IsPaid38 = (bool)BillRet.IsPaid.GetValue();
            }
            //Get value of ExternalGUID
            if(BillRet.ExternalGUID != null)
            {
                string ExternalGUID39 = (string)BillRet.ExternalGUID.GetValue();
            }
            if(BillRet.LinkedTxnList != null)
            {
                for(int i40 = 0; i40 < BillRet.LinkedTxnList.Count; i40++)
                {
                    ILinkedTxn LinkedTxn = BillRet.LinkedTxnList.GetAt(i40);
                    //Get value of TxnID
                    string TxnID41 = (string)LinkedTxn.TxnID.GetValue();
                    //Get value of TxnType
                    ENTxnType TxnType42 = (ENTxnType)LinkedTxn.TxnType.GetValue();
                    //Get value of TxnDate
                    DateTime TxnDate43 = (DateTime)LinkedTxn.TxnDate.GetValue();
                    //Get value of RefNumber
                    if(LinkedTxn.RefNumber != null)
                    {
                        string RefNumber44 = (string)LinkedTxn.RefNumber.GetValue();
                    }
                    //Get value of LinkType
                    if(LinkedTxn.LinkType != null)
                    {
                        ENLinkType LinkType45 = (ENLinkType)LinkedTxn.LinkType.GetValue();
                    }
                    //Get value of Amount
                    double Amount46 = (double)LinkedTxn.Amount.GetValue();
                }
            }
            if(BillRet.ExpenseLineRetList != null)
            {
                for(int i47 = 0; i47 < BillRet.ExpenseLineRetList.Count; i47++)
                {
                    IExpenseLineRet ExpenseLineRet = BillRet.ExpenseLineRetList.GetAt(i47);
                    //Get value of TxnLineID
                    string TxnLineID48 = (string)ExpenseLineRet.TxnLineID.GetValue();
                    if(ExpenseLineRet.AccountRef != null)
                    {
                        //Get value of ListID
                        if(ExpenseLineRet.AccountRef.ListID != null)
                        {
                            string ListID49 = (string)ExpenseLineRet.AccountRef.ListID.GetValue();
                        }
                        //Get value of FullName
                        if(ExpenseLineRet.AccountRef.FullName != null)
                        {
                            string FullName50 = (string)ExpenseLineRet.AccountRef.FullName.GetValue();
                        }
                    }
                    //Get value of Amount
                    if(ExpenseLineRet.Amount != null)
                    {
                        double Amount51 = (double)ExpenseLineRet.Amount.GetValue();
                    }
                    //Get value of Memo
                    if(ExpenseLineRet.Memo != null)
                    {
                        string Memo52 = (string)ExpenseLineRet.Memo.GetValue();
                    }
                    if(ExpenseLineRet.CustomerRef != null)
                    {
                        //Get value of ListID
                        if(ExpenseLineRet.CustomerRef.ListID != null)
                        {
                            string ListID53 = (string)ExpenseLineRet.CustomerRef.ListID.GetValue();
                        }
                        //Get value of FullName
                        if(ExpenseLineRet.CustomerRef.FullName != null)
                        {
                            string FullName54 = (string)ExpenseLineRet.CustomerRef.FullName.GetValue();
                        }
                    }
                    if(ExpenseLineRet.ClassRef != null)
                    {
                        //Get value of ListID
                        if(ExpenseLineRet.ClassRef.ListID != null)
                        {
                            string ListID55 = (string)ExpenseLineRet.ClassRef.ListID.GetValue();
                        }
                        //Get value of FullName
                        if(ExpenseLineRet.ClassRef.FullName != null)
                        {
                            string FullName56 = (string)ExpenseLineRet.ClassRef.FullName.GetValue();
                        }
                    }
                    //Get value of BillableStatus
                    if(ExpenseLineRet.BillableStatus != null)
                    {
                        ENBillableStatus BillableStatus57 = (ENBillableStatus)ExpenseLineRet.BillableStatus.GetValue();
                    }
                    if(ExpenseLineRet.SalesRepRef != null)
                    {
                        //Get value of ListID
                        if(ExpenseLineRet.SalesRepRef.ListID != null)
                        {
                            string ListID58 = (string)ExpenseLineRet.SalesRepRef.ListID.GetValue();
                        }
                        //Get value of FullName
                        if(ExpenseLineRet.SalesRepRef.FullName != null)
                        {
                            string FullName59 = (string)ExpenseLineRet.SalesRepRef.FullName.GetValue();
                        }
                    }
                    if(ExpenseLineRet.DataExtRetList != null)
                    {
                        for(int i60 = 0; i60 < ExpenseLineRet.DataExtRetList.Count; i60++)
                        {
                            IDataExtRet DataExtRet = ExpenseLineRet.DataExtRetList.GetAt(i60);
                            //Get value of OwnerID
                            if(DataExtRet.OwnerID != null)
                            {
                                string OwnerID61 = (string)DataExtRet.OwnerID.GetValue();
                            }
                            //Get value of DataExtName
                            string DataExtName62 = (string)DataExtRet.DataExtName.GetValue();
                            //Get value of DataExtType
                            ENDataExtType DataExtType63 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                            //Get value of DataExtValue
                            string DataExtValue64 = (string)DataExtRet.DataExtValue.GetValue();
                        }
                    }
                }
            }
            if(BillRet.ORItemLineRetList != null)
            {
                for(int i65 = 0; i65 < BillRet.ORItemLineRetList.Count; i65++)
                {
                    IORItemLineRet ORItemLineRet = BillRet.ORItemLineRetList.GetAt(i65);
                    if(ORItemLineRet.ItemLineRet != null)
                    {
                        if(ORItemLineRet.ItemLineRet != null)
                        {
                            //Get value of TxnLineID
                            string TxnLineID66 = (string)ORItemLineRet.ItemLineRet.TxnLineID.GetValue();
                            if(ORItemLineRet.ItemLineRet.ItemRef != null)
                            {
                                //Get value of ListID
                                if(ORItemLineRet.ItemLineRet.ItemRef.ListID != null)
                                {
                                    string ListID67 = (string)ORItemLineRet.ItemLineRet.ItemRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if(ORItemLineRet.ItemLineRet.ItemRef.FullName != null)
                                {
                                    string FullName68 = (string)ORItemLineRet.ItemLineRet.ItemRef.FullName.GetValue();
                                }
                            }
                            if(ORItemLineRet.ItemLineRet.InventorySiteRef != null)
                            {
                                //Get value of ListID
                                if(ORItemLineRet.ItemLineRet.InventorySiteRef.ListID != null)
                                {
                                    string ListID69 = (string)ORItemLineRet.ItemLineRet.InventorySiteRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if(ORItemLineRet.ItemLineRet.InventorySiteRef.FullName != null)
                                {
                                    string FullName70 = (string)ORItemLineRet.ItemLineRet.InventorySiteRef.FullName.GetValue();
                                }
                            }
                            if(ORItemLineRet.ItemLineRet.InventorySiteLocationRef != null)
                            {
                                //Get value of ListID
                                if(ORItemLineRet.ItemLineRet.InventorySiteLocationRef.ListID != null)
                                {
                                    string ListID71 = (string)ORItemLineRet.ItemLineRet.InventorySiteLocationRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if(ORItemLineRet.ItemLineRet.InventorySiteLocationRef.FullName != null)
                                {
                                    string FullName72 = (string)ORItemLineRet.ItemLineRet.InventorySiteLocationRef.FullName.GetValue();
                                }
                            }
                            if(ORItemLineRet.ItemLineRet.ORSerialLotNumber != null)
                            {
                                if(ORItemLineRet.ItemLineRet.ORSerialLotNumber.SerialNumber != null)
                                {
                                    //Get value of SerialNumber
                                    if(ORItemLineRet.ItemLineRet.ORSerialLotNumber.SerialNumber != null)
                                    {
                                        string SerialNumber73 = (string)ORItemLineRet.ItemLineRet.ORSerialLotNumber.SerialNumber.GetValue();
                                    }
                                }
                                if(ORItemLineRet.ItemLineRet.ORSerialLotNumber.LotNumber != null)
                                {
                                    //Get value of LotNumber
                                    if(ORItemLineRet.ItemLineRet.ORSerialLotNumber.LotNumber != null)
                                    {
                                        string LotNumber74 = (string)ORItemLineRet.ItemLineRet.ORSerialLotNumber.LotNumber.GetValue();
                                    }
                                }
                            }
                            //Get value of Desc
                            if(ORItemLineRet.ItemLineRet.Desc != null)
                            {
                                string Desc75 = (string)ORItemLineRet.ItemLineRet.Desc.GetValue();
                            }
                            //Get value of Quantity
                            if(ORItemLineRet.ItemLineRet.Quantity != null)
                            {
                                int Quantity76 = (int)ORItemLineRet.ItemLineRet.Quantity.GetValue();
                            }
                            //Get value of UnitOfMeasure
                            if(ORItemLineRet.ItemLineRet.UnitOfMeasure != null)
                            {
                                string UnitOfMeasure77 = (string)ORItemLineRet.ItemLineRet.UnitOfMeasure.GetValue();
                            }
                            if(ORItemLineRet.ItemLineRet.OverrideUOMSetRef != null)
                            {
                                //Get value of ListID
                                if(ORItemLineRet.ItemLineRet.OverrideUOMSetRef.ListID != null)
                                {
                                    string ListID78 = (string)ORItemLineRet.ItemLineRet.OverrideUOMSetRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if(ORItemLineRet.ItemLineRet.OverrideUOMSetRef.FullName != null)
                                {
                                    string FullName79 = (string)ORItemLineRet.ItemLineRet.OverrideUOMSetRef.FullName.GetValue();
                                }
                            }
                            //Get value of Cost
                            if(ORItemLineRet.ItemLineRet.Cost != null)
                            {
                                double Cost80 = (double)ORItemLineRet.ItemLineRet.Cost.GetValue();
                            }
                            //Get value of Amount
                            if(ORItemLineRet.ItemLineRet.Amount != null)
                            {
                                double Amount81 = (double)ORItemLineRet.ItemLineRet.Amount.GetValue();
                            }
                            if(ORItemLineRet.ItemLineRet.CustomerRef != null)
                            {
                                //Get value of ListID
                                if(ORItemLineRet.ItemLineRet.CustomerRef.ListID != null)
                                {
                                    string ListID82 = (string)ORItemLineRet.ItemLineRet.CustomerRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if(ORItemLineRet.ItemLineRet.CustomerRef.FullName != null)
                                {
                                    string FullName83 = (string)ORItemLineRet.ItemLineRet.CustomerRef.FullName.GetValue();
                                }
                            }
                            if(ORItemLineRet.ItemLineRet.ClassRef != null)
                            {
                                //Get value of ListID
                                if(ORItemLineRet.ItemLineRet.ClassRef.ListID != null)
                                {
                                    string ListID84 = (string)ORItemLineRet.ItemLineRet.ClassRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if(ORItemLineRet.ItemLineRet.ClassRef.FullName != null)
                                {
                                    string FullName85 = (string)ORItemLineRet.ItemLineRet.ClassRef.FullName.GetValue();
                                }
                            }
                            //Get value of BillableStatus
                            if(ORItemLineRet.ItemLineRet.BillableStatus != null)
                            {
                                ENBillableStatus BillableStatus86 = (ENBillableStatus)ORItemLineRet.ItemLineRet.BillableStatus.GetValue();
                            }
                            if(ORItemLineRet.ItemLineRet.SalesRepRef != null)
                            {
                                //Get value of ListID
                                if(ORItemLineRet.ItemLineRet.SalesRepRef.ListID != null)
                                {
                                    string ListID87 = (string)ORItemLineRet.ItemLineRet.SalesRepRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if(ORItemLineRet.ItemLineRet.SalesRepRef.FullName != null)
                                {
                                    string FullName88 = (string)ORItemLineRet.ItemLineRet.SalesRepRef.FullName.GetValue();
                                }
                            }
                            if(ORItemLineRet.ItemLineRet.DataExtRetList != null)
                            {
                                for(int i89 = 0; i89 < ORItemLineRet.ItemLineRet.DataExtRetList.Count; i89++)
                                {
                                    IDataExtRet DataExtRet = ORItemLineRet.ItemLineRet.DataExtRetList.GetAt(i89);
                                    //Get value of OwnerID
                                    if(DataExtRet.OwnerID != null)
                                    {
                                        string OwnerID90 = (string)DataExtRet.OwnerID.GetValue();
                                    }
                                    //Get value of DataExtName
                                    string DataExtName91 = (string)DataExtRet.DataExtName.GetValue();
                                    //Get value of DataExtType
                                    ENDataExtType DataExtType92 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                                    //Get value of DataExtValue
                                    string DataExtValue93 = (string)DataExtRet.DataExtValue.GetValue();
                                }
                            }
                        }
                    }
                    if(ORItemLineRet.ItemGroupLineRet != null)
                    {
                        if(ORItemLineRet.ItemGroupLineRet != null)
                        {
                            //Get value of TxnLineID
                            string TxnLineID94 = (string)ORItemLineRet.ItemGroupLineRet.TxnLineID.GetValue();
                            //Get value of ListID
                            if(ORItemLineRet.ItemGroupLineRet.ItemGroupRef.ListID != null)
                            {
                                string ListID95 = (string)ORItemLineRet.ItemGroupLineRet.ItemGroupRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if(ORItemLineRet.ItemGroupLineRet.ItemGroupRef.FullName != null)
                            {
                                string FullName96 = (string)ORItemLineRet.ItemGroupLineRet.ItemGroupRef.FullName.GetValue();
                            }
                            //Get value of Desc
                            if(ORItemLineRet.ItemGroupLineRet.Desc != null)
                            {
                                string Desc97 = (string)ORItemLineRet.ItemGroupLineRet.Desc.GetValue();
                            }
                            //Get value of Quantity
                            if(ORItemLineRet.ItemGroupLineRet.Quantity != null)
                            {
                                int Quantity98 = (int)ORItemLineRet.ItemGroupLineRet.Quantity.GetValue();
                            }
                            //Get value of UnitOfMeasure
                            if(ORItemLineRet.ItemGroupLineRet.UnitOfMeasure != null)
                            {
                                string UnitOfMeasure99 = (string)ORItemLineRet.ItemGroupLineRet.UnitOfMeasure.GetValue();
                            }
                            if(ORItemLineRet.ItemGroupLineRet.OverrideUOMSetRef != null)
                            {
                                //Get value of ListID
                                if(ORItemLineRet.ItemGroupLineRet.OverrideUOMSetRef.ListID != null)
                                {
                                    string ListID100 = (string)ORItemLineRet.ItemGroupLineRet.OverrideUOMSetRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if(ORItemLineRet.ItemGroupLineRet.OverrideUOMSetRef.FullName != null)
                                {
                                    string FullName101 = (string)ORItemLineRet.ItemGroupLineRet.OverrideUOMSetRef.FullName.GetValue();
                                }
                            }
                            //Get value of TotalAmount
                            double TotalAmount102 = (double)ORItemLineRet.ItemGroupLineRet.TotalAmount.GetValue();
                            if(ORItemLineRet.ItemGroupLineRet.ItemLineRetList != null)
                            {
                                for(int i103 = 0; i103 < ORItemLineRet.ItemGroupLineRet.ItemLineRetList.Count; i103++)
                                {
                                    IItemLineRet ItemLineRet = ORItemLineRet.ItemGroupLineRet.ItemLineRetList.GetAt(i103);
                                    //Get value of TxnLineID
                                    string TxnLineID104 = (string)ItemLineRet.TxnLineID.GetValue();
                                    if(ItemLineRet.ItemRef != null)
                                    {
                                        //Get value of ListID
                                        if(ItemLineRet.ItemRef.ListID != null)
                                        {
                                            string ListID105 = (string)ItemLineRet.ItemRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if(ItemLineRet.ItemRef.FullName != null)
                                        {
                                            string FullName106 = (string)ItemLineRet.ItemRef.FullName.GetValue();
                                        }
                                    }
                                    if(ItemLineRet.InventorySiteRef != null)
                                    {
                                        //Get value of ListID
                                        if(ItemLineRet.InventorySiteRef.ListID != null)
                                        {
                                            string ListID107 = (string)ItemLineRet.InventorySiteRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if(ItemLineRet.InventorySiteRef.FullName != null)
                                        {
                                            string FullName108 = (string)ItemLineRet.InventorySiteRef.FullName.GetValue();
                                        }
                                    }
                                    if(ItemLineRet.InventorySiteLocationRef != null)
                                    {
                                        //Get value of ListID
                                        if(ItemLineRet.InventorySiteLocationRef.ListID != null)
                                        {
                                            string ListID109 = (string)ItemLineRet.InventorySiteLocationRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if(ItemLineRet.InventorySiteLocationRef.FullName != null)
                                        {
                                            string FullName110 = (string)ItemLineRet.InventorySiteLocationRef.FullName.GetValue();
                                        }
                                    }
                                    if(ItemLineRet.ORSerialLotNumber != null)
                                    {
                                        if(ItemLineRet.ORSerialLotNumber.SerialNumber != null)
                                        {
                                            //Get value of SerialNumber
                                            if(ItemLineRet.ORSerialLotNumber.SerialNumber != null)
                                            {
                                                string SerialNumber111 = (string)ItemLineRet.ORSerialLotNumber.SerialNumber.GetValue();
                                            }
                                        }
                                        if(ItemLineRet.ORSerialLotNumber.LotNumber != null)
                                        {
                                            //Get value of LotNumber
                                            if(ItemLineRet.ORSerialLotNumber.LotNumber != null)
                                            {
                                                string LotNumber112 = (string)ItemLineRet.ORSerialLotNumber.LotNumber.GetValue();
                                            }
                                        }
                                    }
                                    //Get value of Desc
                                    if(ItemLineRet.Desc != null)
                                    {
                                        string Desc113 = (string)ItemLineRet.Desc.GetValue();
                                    }
                                    //Get value of Quantity
                                    if(ItemLineRet.Quantity != null)
                                    {
                                        int Quantity114 = (int)ItemLineRet.Quantity.GetValue();
                                    }
                                    //Get value of UnitOfMeasure
                                    if(ItemLineRet.UnitOfMeasure != null)
                                    {
                                        string UnitOfMeasure115 = (string)ItemLineRet.UnitOfMeasure.GetValue();
                                    }
                                    if(ItemLineRet.OverrideUOMSetRef != null)
                                    {
                                        //Get value of ListID
                                        if(ItemLineRet.OverrideUOMSetRef.ListID != null)
                                        {
                                            string ListID116 = (string)ItemLineRet.OverrideUOMSetRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if(ItemLineRet.OverrideUOMSetRef.FullName != null)
                                        {
                                            string FullName117 = (string)ItemLineRet.OverrideUOMSetRef.FullName.GetValue();
                                        }
                                    }
                                    //Get value of Cost
                                    if(ItemLineRet.Cost != null)
                                    {
                                        double Cost118 = (double)ItemLineRet.Cost.GetValue();
                                    }
                                    //Get value of Amount
                                    if(ItemLineRet.Amount != null)
                                    {
                                        double Amount119 = (double)ItemLineRet.Amount.GetValue();
                                    }
                                    if(ItemLineRet.CustomerRef != null)
                                    {
                                        //Get value of ListID
                                        if(ItemLineRet.CustomerRef.ListID != null)
                                        {
                                            string ListID120 = (string)ItemLineRet.CustomerRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if(ItemLineRet.CustomerRef.FullName != null)
                                        {
                                            string FullName121 = (string)ItemLineRet.CustomerRef.FullName.GetValue();
                                        }
                                    }
                                    if(ItemLineRet.ClassRef != null)
                                    {
                                        //Get value of ListID
                                        if(ItemLineRet.ClassRef.ListID != null)
                                        {
                                            string ListID122 = (string)ItemLineRet.ClassRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if(ItemLineRet.ClassRef.FullName != null)
                                        {
                                            string FullName123 = (string)ItemLineRet.ClassRef.FullName.GetValue();
                                        }
                                    }
                                    //Get value of BillableStatus
                                    if(ItemLineRet.BillableStatus != null)
                                    {
                                        ENBillableStatus BillableStatus124 = (ENBillableStatus)ItemLineRet.BillableStatus.GetValue();
                                    }
                                    if(ItemLineRet.SalesRepRef != null)
                                    {
                                        //Get value of ListID
                                        if(ItemLineRet.SalesRepRef.ListID != null)
                                        {
                                            string ListID125 = (string)ItemLineRet.SalesRepRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if(ItemLineRet.SalesRepRef.FullName != null)
                                        {
                                            string FullName126 = (string)ItemLineRet.SalesRepRef.FullName.GetValue();
                                        }
                                    }
                                    if(ItemLineRet.DataExtRetList != null)
                                    {
                                        for(int i127 = 0; i127 < ItemLineRet.DataExtRetList.Count; i127++)
                                        {
                                            IDataExtRet DataExtRet = ItemLineRet.DataExtRetList.GetAt(i127);
                                            //Get value of OwnerID
                                            if(DataExtRet.OwnerID != null)
                                            {
                                                string OwnerID128 = (string)DataExtRet.OwnerID.GetValue();
                                            }
                                            //Get value of DataExtName
                                            string DataExtName129 = (string)DataExtRet.DataExtName.GetValue();
                                            //Get value of DataExtType
                                            ENDataExtType DataExtType130 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                                            //Get value of DataExtValue
                                            string DataExtValue131 = (string)DataExtRet.DataExtValue.GetValue();
                                        }
                                    }
                                }
                            }
                            if(ORItemLineRet.ItemGroupLineRet.DataExtList != null)
                            {
                                for(int i132 = 0; i132 < ORItemLineRet.ItemGroupLineRet.DataExtList.Count; i132++)
                                {
                                    IDataExt DataExt = ORItemLineRet.ItemGroupLineRet.DataExtList.GetAt(i132);
                                    //Get value of OwnerID
                                    string OwnerID133 = (string)DataExt.OwnerID.GetValue();
                                    //Get value of DataExtName
                                    string DataExtName134 = (string)DataExt.DataExtName.GetValue();
                                    //Get value of DataExtValue
                                    string DataExtValue135 = (string)DataExt.DataExtValue.GetValue();
                                }
                            }
                        }
                    }
                }
            }
            //Get value of OpenAmount
            if(BillRet.OpenAmount != null)
            {
                double OpenAmount136 = (double)BillRet.OpenAmount.GetValue();
            }
            if(BillRet.DataExtRetList != null)
            {
                for(int i137 = 0; i137 < BillRet.DataExtRetList.Count; i137++)
                {
                    IDataExtRet DataExtRet = BillRet.DataExtRetList.GetAt(i137);
                    //Get value of OwnerID
                    if(DataExtRet.OwnerID != null)
                    {
                        string OwnerID138 = (string)DataExtRet.OwnerID.GetValue();
                    }
                    //Get value of DataExtName
                    string DataExtName139 = (string)DataExtRet.DataExtName.GetValue();
                    //Get value of DataExtType
                    ENDataExtType DataExtType140 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                    //Get value of DataExtValue
                    string DataExtValue141 = (string)DataExtRet.DataExtValue.GetValue();
                }
            }
        }

    }

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
    }
}
