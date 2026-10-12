using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using Microsoft.ApplicationBlocks.Data;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using System.Collections;
using System.Threading;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
using Sap.Data.Hana;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Linq;
using System.Text;
using System.Configuration;
using CrystalDecisions.CrystalReports.ViewerObjectModel;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Drawing.Printing;
using System.Net.Mail;
using System.Threading;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace LandedCost_SpareParts
{
    class ClsLandedCost : Base
    {

         #region VARIABLE DECLARATION

        private static HanaConnection _conn;
        private static HanaDataAdapter _da;
        //private static DataTable _dataTable;
        private static HanaParameter _parm1;
        private static HanaParameter _parm2;
        private static int _errors = 0;
        string _LastErrorMsg = " ", strcon = "";
        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        SAPbouiCOM.Button _Button;
        private bool flag;
        public SAPbouiCOM.EditText _EditText;
        public SAPbouiCOM.Matrix _Matrix;
        private SAPbouiCOM.Grid _Grid;
        private SAPbouiCOM.DataTable _GridDataTable, _GridDataTable1;
        private SAPbouiCOM.EditTextColumn _Column;
        private SAPbouiCOM.RowHeaders _RowHeaders;
        public static string LandedCostNo = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        SAPbouiCOM.ComboBox _ComboBox;
        private SAPbouiCOM.DBDataSource _DBDataSource;
        SAPbobsCOM.Recordset _RS = null, _RSSeries = null, _RS1 = null, _RSLandedCost = null, _Rs = null, _RSChk = null, _RsCal = null, _RSChkGetVal = null, _RSChk1 = null, _RSPriceList = null, _RSLandedC = null, _RsBP = null;
        string SQL1 = string.Empty, SQL = string.Empty, StrLandedCostSQL = string.Empty;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource1, _ChildDBDataSource2, _ChildDBDataSource3;
        SAPbobsCOM.UserTable UDTHeader, UDTLine1, UDTLine2;
        public static string _DocNum =string.Empty;
        public static string _DocEntry = string.Empty, _Branch = string.Empty;
        public static double TotalVAT = 0;
        public static string ControlAC = string.Empty,CustomOfficeTI = string.Empty,_DocDate,_DueDate,_PPNo;
        public static double TotalExciseDuty = 0;
        public static double TotalCustomDuty = 0;
        public static int count = 1;
        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;
        public bool _flag = true;
        string _ImportFileName = "";
        SAPbouiCOM.Form _FormNegative = null;

        #endregion

         #region CONSTRUCTOR & DISTRUCTOR
        public ClsLandedCost()
            : base()
        {
            //_FormNegative = Utilities.Application.SBO_Application.Forms.GetForm("-992", _Form.TypeCount);
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsLandedCost()
        {
        }
        #endregion

         #region ITEM EVENT
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            #region BEFORE ACTION ACTION
            if (pVal.Before_Action == true)
            {
                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode != BoFormMode.fm_FIND_MODE)
                        {
                            TotalCustomDuty = Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OIPF").GetValue("U_TotalCustomDuty", 0).ToString());
                            TotalExciseDuty = Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OIPF").GetValue("U_TotalExciseDuty", 0).ToString());
                            TotalVAT = Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OIPF").GetValue("U_TotalVATAmount", 0).ToString());
                            ControlAC = _Form.DataSources.DBDataSources.Item("OIPF").GetValue("U_ControlAccount", 0).ToString();
                            _DocNum = _Form.DataSources.DBDataSources.Item("OIPF").GetValue("DocNum", 0).ToString();
                            _DocEntry = _Form.DataSources.DBDataSources.Item("OIPF").GetValue("DocEntry", 0).ToString();
                            _Branch = _Form.DataSources.DBDataSources.Item("OIPF").GetValue("BPLId", 0).ToString();
                            CustomOfficeTI = _Form.DataSources.DBDataSources.Item("OIPF").GetValue("AgentCode", 0).ToString();
                            _DocDate = _Form.DataSources.DBDataSources.Item("OIPF").GetValue("DocDate", 0).ToString();
                            _DueDate = _Form.DataSources.DBDataSources.Item("OIPF").GetValue("DocDueDate", 0).ToString();
                            _PPNo = _Form.DataSources.DBDataSources.Item("OIPF").GetValue("U_PPNo", 0).ToString();
                            _flag = false;
                        }
                       

                        break;
                    #endregion

                     //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                     //   if (pVal.ItemUID == "U_TotalCustomDuty" && pVal.CharPressed != 9 && _Form.Mode != BoFormMode.fm_ADD_MODE)
                     //   {
                     //       BubbleEvent = false;
                     //       return;
                     //   }
                     //   break;

                }
            }
            #endregion

            #region After ACTION ACTION
            else if (pVal.Before_Action == false)
            {

                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && pVal.Action_Success == true && _Form.Mode != BoFormMode.fm_FIND_MODE)
                        {
                            if (_Form.DataSources.DBDataSources.Item("OIPF").GetValue("U_JENO", 0).ToString().Trim() == "" && _flag == false)
                            {

                                GeneralEntry(LandedCostNo);
                                _flag = true;

                            }
                            //UpdatePriceList(LandedCostNo);
                        }
                        break;

                    
                    #endregion
                }
            }
        }

            #endregion

            #region Post General Entry 

        public void GeneralEntry(string LandedCostNo)
        {
            string JENO = string.Empty;
            int lRetCode = 0;
            SAPbobsCOM.JournalEntries oJE = null;
            oJE = (SAPbobsCOM.JournalEntries)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            string strQuery = "SELECT \"U_VATINAC\",\"U_CustDtyAC\",\"U_ExDtyAC\",\"U_AdyAllAC\",\"U_SalesMrgn\" FROM \"@CSTMDSTPH\" ";
            Utilities.ExecuteSQL(ref _RS, strQuery);

            string strQuery12 = "select \"LaCAllcAcc\" from OALC where \"AlcCode\" = '06'";
            Utilities.ExecuteSQL(ref _RSLandedC, strQuery12);

            string StrQuery23 = "select Max(\"DocEntry\") as \"DocEntry\" from OIPF where \"DocNum\" = '" + _DocNum + "'";
            Utilities.ExecuteSQL(ref _RSChk, StrQuery23);

            string strQuery22 = "select \"CostSum\" from IPF2 where  \"DocEntry\" = '" + _RSChk.Fields.Item("DocEntry").Value.ToString() + "' AND \"AlcCode\" = '06' ";
            Utilities.ExecuteSQL(ref _RS1, strQuery22);

            string _Sql11 = "select \"AliasName\"  FROM OCRD where \"CardCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("43").Specific).Value.ToString().Trim() + "' ";
            Utilities.ExecuteSQL(ref _RsBP, _Sql11);
            string Remarks = _RsBP.Fields.Item("AliasName").Value.ToString() + _PPNo.ToString();


            //oJE.Lines.SetCurrentLine(0); //C00001 208010
            DateTime aa = Utilities.ToDate(DateTime.Now.ToString("yyyyMMdd"));
            oJE.TaxDate = Utilities.ToDate(_DocDate.ToString());
            oJE.DueDate = Utilities.ToDate(_DueDate.ToString());
            oJE.ReferenceDate = Utilities.ToDate(_DocDate.ToString());
            oJE.Memo = Remarks.ToString();
            oJE.UserFields.Fields.Item("U_LandedCostEntry").Value = _RSChk.Fields.Item("DocEntry").Value.ToString();
            oJE.Reference = _DocNum.ToString();
            oJE.Reference2 = Remarks.ToString();


            //Credit Custom Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(0);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_CustDtyAC").Value.ToString().Trim();
            oJE.Lines.Credit = TotalCustomDuty;
            oJE.Lines.Debit = 0;
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit Custom Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(1);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_AdyAllAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = TotalCustomDuty;
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Credit Excise Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(2);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_ExDtyAC").Value.ToString().Trim();
            oJE.Lines.Credit = TotalExciseDuty;
            oJE.Lines.Debit = 0;
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit Excise Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(3);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_AdyAllAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = TotalExciseDuty;
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Credit
            oJE.Lines.SetCurrentLine(4);
            oJE.Lines.ShortName = CustomOfficeTI;
            oJE.Lines.AccountCode = ControlAC;
            oJE.Lines.Credit = TotalVAT;
            oJE.Lines.Debit = 0;
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(5);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_VATINAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = TotalVAT;
            oJE.Lines.TaxCode = "VAT 13%";
            oJE.Lines.TaxGroup = "VAT 13%";
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());


            //-------------------//
            //Credit Custom Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(6);
            //oJE.Lines.AccountCode = ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString().Trim();
            oJE.Lines.ShortName = CustomOfficeTI;
            oJE.Lines.AccountCode = ControlAC;
            oJE.Lines.Credit = TotalCustomDuty;
            oJE.Lines.Debit = 0;
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit Custom Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(7);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_CustDtyAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = TotalCustomDuty;
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Credit Excise Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(8);
            oJE.Lines.ShortName = CustomOfficeTI;
            oJE.Lines.AccountCode = ControlAC;
            oJE.Lines.Credit = TotalExciseDuty;
            oJE.Lines.Debit = 0;
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit Excise Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(9);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_ExDtyAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = TotalExciseDuty;
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            ///////////////////////////////////////

            //Credit
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(10);
            oJE.Lines.ShortName = CustomOfficeTI;
            oJE.Lines.AccountCode = ControlAC;
            oJE.Lines.Credit = Convert.ToDouble(_RS1.Fields.Item("CostSum").Value.ToString().Trim());
            oJE.Lines.Debit = 0;
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(11);
            oJE.Lines.AccountCode = _RSLandedC.Fields.Item("LaCAllcAcc").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = Convert.ToDouble(_RS1.Fields.Item("CostSum").Value.ToString().Trim());
            oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());
            
            lRetCode = oJE.Add();
            if (lRetCode != 0)
            {
                int oErrCode;
                string oErrMsg;
                Utilities.Application.Company.GetLastError(out oErrCode, out oErrMsg);
            }
            else
            {
                Utilities.Application.Company.GetNewObjectCode(out JENO);
                string strQuery1 = "update \"OIPF\" set \"U_JENO\" = '" + JENO + "' where \"DocEntry\" = '" + _RSChk.Fields.Item("DocEntry").Value.ToString() + "'";
                Utilities.ExecuteSQL(ref _RS, strQuery1);
            }

        }

            #endregion

            #region Update Price List

        public void UpdatePriceList(string LandedCostNo)
        {
            SAPbobsCOM.Items _Items = (SAPbobsCOM.Items)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oItems);
            int lRetCode = 0;
            double MarginPer = 0;

            string strQuery = "SELECT \"U_SalesMrgn\" FROM \"@CSTMDSTPH\" ";
            Utilities.ExecuteSQL(ref _RS, strQuery);

            MarginPer = Convert.ToDouble(_RS.Fields.Item("U_SalesMrgn").Value.ToString());
            string strQueryPriceList = "select \"ItemCode\",\"PriceAtWH\" + ((\"PriceAtWH\" * " + MarginPer + ")/100) as \"MarginPrice\" FROM IPF1 where \"DocEntry\" = '" + _DocEntry + "'";//" + LandedCostNo + "
            Utilities.ExecuteSQL(ref _RSPriceList, strQueryPriceList);
            for (int i = 0; i < _RSPriceList.RecordCount; i++)
            {
                if (_Items.GetByKey(_RSPriceList.Fields.Item("ItemCode").Value.ToString().Trim()))
                {
                    _Items.PriceList.SetCurrentLine(0);
                    _Items.PriceList.Price = Convert.ToDouble(_RSPriceList.Fields.Item("MarginPrice").Value.ToString().Trim());
                }
                lRetCode = _Items.Update();
                _RSPriceList.MoveNext();
            }
        }

        #endregion

        }
        #endregion

    }

