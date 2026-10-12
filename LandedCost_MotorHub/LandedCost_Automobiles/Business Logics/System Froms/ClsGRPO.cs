using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LandedCost_Automobiles
{
    class ClsGRPO : Base
    {
       
        #region VARIABLE DECLARATION
        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        SAPbouiCOM.Button _Button;
        SAPbouiCOM.ComboBox _ComboBox;
        SAPbouiCOM.StaticText _StaticText;
        private bool flag;
        public SAPbouiCOM.EditText _EditText;
        public SAPbouiCOM.Matrix _Matrix;
        string _DocEntry = "", _DocCurency = "", _DocNum = "", _BPCode = "", _BPName = "", _Branch = "", _PostingDate = "", _DueDate = "", _ConSignmentType = "", _CountryCode = "", _CountryName = "",_DocumentType = "";
        double _DocTotal = 0, _DocTotal1 = 0;
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RSLandedCost = null, _Rs = null, _RSChk = null, _RsCal = null, _RSChkGetVal = null, _RSChk1 = null, _RSPriceList = null, _RSAPInvoice = null, _RSGRPO = null;
        SAPbouiCOM.Folder _Folder;
        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsGRPO()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsGRPO()
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
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemUID == "TxtExRate" && pVal.InnerEvent == false)
                        {
                            _Form.Freeze(true);
                            _Form.PaneLevel = 5;
                            double _CustomDuty = 0, _FinalCustomDuty = 0, _ExciseDuty = 0, _FinExciseDuty = 0;
                            _DocTotal1 = Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocTotalFC", 0).ToString()) - Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("VatSumFC", 0).ToString());
                            ((SAPbouiCOM.EditText)_Form.Items.Item("TxtInvVal").Specific).Value = _DocTotal1.ToString();
                            ((SAPbouiCOM.EditText)_Form.Items.Item("TxtValNRS").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtExRate").Specific).Value) * _DocTotal1).ToString();

                            string _SQL = "Select *  from \"@CSTMDSTPH\"";
                            Utilities.ExecuteSQL(ref _Rs, _SQL);

                            string _SQLGRPO = "select T0.\"LineTotal\",IFNULL(T3.\"U_CusDutyP\",0) as \"U_CusDutyP\",IFNULL(T3.\"U_ExDutyP\",0) as \"U_ExDutyP\" from PDN1 T0 " +
                            " inner join \"OPDN\" T4 ON T0.\"DocEntry\" = T4.\"DocEntry\" " +
                            " inner join \"OITM\" T1 ON T0.\"ItemCode\" = T1.\"ItemCode\"  " +
                            " left outer join \"@ITEMDUTYSETUPM\" T2 ON T2.\"U_ItemCode\" = T1.\"ItemCode\" " +
                            " left outer join \"@ITEMDUTYSETUPC\" T3 ON T2.\"Code\" = T3.\"Code\" and T4.\"U_CountryC\" = T3.\"U_CountryC\" " +
                            " where T0.\"DocEntry\" = '" + _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString() + "'";
                            Utilities.ExecuteSQL(ref _RSGRPO, _SQLGRPO);
                            if (_RSGRPO.RecordCount > 0)
                            {
                                for (int i = 0; i < _RSGRPO.RecordCount; i++)
                                {
                                    _CustomDuty = (_RSGRPO.Fields.Item("LineTotal").Value * _RSGRPO.Fields.Item("U_CusDutyP").Value) / 100;
                                    _ExciseDuty = ((_RSGRPO.Fields.Item("LineTotal").Value + _CustomDuty) * _RSGRPO.Fields.Item("U_ExDutyP").Value) / 100;
                                    _FinalCustomDuty = _FinalCustomDuty + _CustomDuty;
                                    _FinExciseDuty = _FinExciseDuty + _ExciseDuty;
                                }
                            }
                            ((SAPbouiCOM.EditText)_Form.Items.Item("TxtImpDuty").Specific).Value = _FinalCustomDuty.ToString();
                            ((SAPbouiCOM.EditText)_Form.Items.Item("TxtECS").Specific).Value = _FinExciseDuty.ToString();
                            ((SAPbouiCOM.EditText)_Form.Items.Item("TxtVAT").Specific).Value = ((Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtValNRS").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtImpDuty").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtECS").Specific).Value)) * Convert.ToDouble(_Rs.Fields.Item("U_VATPer").Value) / 100).ToString();
                            ((SAPbouiCOM.EditText)_Form.Items.Item("txtFndReq").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtImpDuty").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtECS").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtVAT").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtPragpn").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtDumrg").Specific).Value)).ToString();

                            _Form.Items.Item("TxtInvVal").Enabled = false;
                            _Form.Items.Item("TxtValNRS").Enabled = false;
                            _Form.Items.Item("TxtImpDuty").Enabled = false;
                            _Form.Items.Item("TxtECS").Enabled = false;
                            _Form.Items.Item("TxtVAT").Enabled = false;
                            //_Form.Items.Item("TxtPragpn").Enabled = false;
                            //_Form.Items.Item("TxtDumrg").Enabled = false;
                            _Form.Items.Item("TxtExRate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Form.Items.Item("txtFndReq").Enabled = false;
                            _Form.Freeze(false);
                        }
                        break;

                    #region ITEM PRESS EVENT

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "BtnLandC")
                        {
                            _DocCurency = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocCur", 0).ToString();
                            if (_DocCurency == "NPR")
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Document currency and local currency both are same.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            _DocEntry = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString();
                            string _SQL = "select OPDN.\"DocNum\" from OPDN inner join PDN1 ON OPDN.\"DocEntry\" = PDN1.\"DocEntry\" where OPDN.\"DocStatus\" = 'C' and PDN1.\"TargetType\" = 21 and PDN1.\"LineStatus\" = 'C' and OPDN.\"DocEntry\" = '" + _DocEntry + "';";
                            Utilities.ExecuteSQL(ref _RS, _SQL);

                            if (_RS.RecordCount > 0)
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Goods return Created for this GRPO.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                           
                        }

                        break;

                    #endregion


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
                        if (pVal.ItemUID == "FldCosting")
                        {
                            if (_Form.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                _Form.Freeze(true);
                                _Form.PaneLevel = 5;
                                double _CustomDuty = 0, _FinalCustomDuty = 0, _ExciseDuty = 0, _FinExciseDuty = 0;
                                _DocTotal1 = Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocTotalFC", 0).ToString()) - Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("VatSumFC", 0).ToString());
                                ((SAPbouiCOM.EditText)_Form.Items.Item("TxtInvVal").Specific).Value = _DocTotal1.ToString();
                                ((SAPbouiCOM.EditText)_Form.Items.Item("TxtExRate").Specific).Value = Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocRate", 0).ToString()).ToString();
                                ((SAPbouiCOM.EditText)_Form.Items.Item("TxtValNRS").Specific).Value = (Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocRate", 0)) * _DocTotal1).ToString();

                                string _SQL = "Select *  from \"@CSTMDSTPH\"";
                                Utilities.ExecuteSQL(ref _Rs, _SQL);

                                string _SQLGRPO = "select T0.\"LineTotal\",IFNULL(T3.\"U_CusDutyP\",0) as \"U_CusDutyP\",IFNULL(T3.\"U_ExDutyP\",0) as \"U_ExDutyP\" from PDN1 T0 " +
                                " inner join \"OPDN\" T4 ON T0.\"DocEntry\" = T4.\"DocEntry\" " +
                                " inner join \"OITM\" T1 ON T0.\"ItemCode\" = T1.\"ItemCode\"  " +
                                " left outer join \"@ITEMDUTYSETUPM\" T2 ON T2.\"U_ItemCode\" = T1.\"ItemCode\" " +
                                " left outer join \"@ITEMDUTYSETUPC\" T3 ON T2.\"Code\" = T3.\"Code\" and T4.\"U_CountryC\" = T3.\"U_CountryC\" " +
                                " where T0.\"DocEntry\" = '" + _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString() + "'";
                                Utilities.ExecuteSQL(ref _RSGRPO, _SQLGRPO);
                                if (_RSGRPO.RecordCount > 0)
                                {
                                    for (int i = 0; i < _RSGRPO.RecordCount; i++)
                                    {
                                        _CustomDuty = (_RSGRPO.Fields.Item("LineTotal").Value * _RSGRPO.Fields.Item("U_CusDutyP").Value) / 100;
                                        _ExciseDuty = ((_RSGRPO.Fields.Item("LineTotal").Value + _CustomDuty) * _RSGRPO.Fields.Item("U_ExDutyP").Value) / 100;
                                        _FinalCustomDuty = _FinalCustomDuty + _CustomDuty;
                                        _FinExciseDuty = _FinExciseDuty + _ExciseDuty;
                                    }
                                }
                                ((SAPbouiCOM.EditText)_Form.Items.Item("TxtImpDuty").Specific).Value = _FinalCustomDuty.ToString();
                                ((SAPbouiCOM.EditText)_Form.Items.Item("TxtECS").Specific).Value = _FinExciseDuty.ToString();
                                ((SAPbouiCOM.EditText)_Form.Items.Item("TxtVAT").Specific).Value = ((Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtValNRS").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtImpDuty").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("TxtECS").Specific).Value)) * Convert.ToDouble(_Rs.Fields.Item("U_VATPer").Value) / 100).ToString();
                                ((SAPbouiCOM.EditText)_Form.Items.Item("txtFndReq").Specific).Value = ((SAPbouiCOM.EditText)_Form.Items.Item("TxtImpDuty").Specific).Value + ((SAPbouiCOM.EditText)_Form.Items.Item("TxtECS").Specific).Value + ((SAPbouiCOM.EditText)_Form.Items.Item("TxtVAT").Specific).Value + ((SAPbouiCOM.EditText)_Form.Items.Item("TxtPragpn").Specific).Value + ((SAPbouiCOM.EditText)_Form.Items.Item("TxtDumrg").Specific).Value;

                                _Form.Items.Item("TxtInvVal").Enabled = false;
                                _Form.Items.Item("TxtValNRS").Enabled = false;
                                _Form.Items.Item("TxtImpDuty").Enabled = false;
                                _Form.Items.Item("TxtECS").Enabled = false;
                                _Form.Items.Item("TxtVAT").Enabled = false;
                                //_Form.Items.Item("TxtPragpn").Enabled = false;
                                //_Form.Items.Item("TxtDumrg").Enabled = false;
                                _Form.Items.Item("TxtExRate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                _Form.Items.Item("txtFndReq").Enabled = false;
                                _Form.Freeze(false);
                            }

                        }
                        if (pVal.ItemUID == "BtnLandC" && _Form.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            _DocEntry = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString();
                            string _SQL = "select TOP 1 PDN1.\"TrgetEntry\" from OPDN inner join PDN1 ON OPDN.\"DocEntry\" = PDN1.\"DocEntry\" where OPDN.\"LndCstNum\" = 0 AND OPDN.\"DocStatus\" = 'C' and PDN1.\"TargetType\" = 18 and PDN1.\"LineStatus\" = 'C' and OPDN.\"DocEntry\" = '" + _DocEntry + "' ";
                            Utilities.ExecuteSQL(ref _RS, _SQL);
                            if (_RS.RecordCount > 0)
                            {
                                string _SQLAPInvoice = "select \"DocEntry\",\"DocNum\",\"CardCode\",\"CardName\",\"DocCur\",\"BPLId\",\"DocDate\",\"DocDueDate\",\"DocTotalFC\",\"VatSumFC\",\"U_ConType\",\"U_CountryC\",\"U_CountryN\",'APInvoice' as \"DocumentType\" from OPCH WHERE OPCH.\"DocEntry\" = '" + _RS.Fields.Item("TrgetEntry").Value.ToString() + "' ";
                                Utilities.ExecuteSQL(ref _RSAPInvoice, _SQLAPInvoice);
                                if (_RS.RecordCount > 0)
                                {
                                    _DocEntry = _RSAPInvoice.Fields.Item("DocEntry").Value.ToString();
                                    _DocNum = _RSAPInvoice.Fields.Item("DocNum").Value.ToString();
                                    _BPCode = _RSAPInvoice.Fields.Item("CardCode").Value.ToString();
                                    _BPName = _RSAPInvoice.Fields.Item("CardName").Value.ToString();
                                    _DocCurency = _RSAPInvoice.Fields.Item("DocCur").Value.ToString();
                                    _Branch = _RSAPInvoice.Fields.Item("BPLId").Value.ToString();
                                    _PostingDate = _RSAPInvoice.Fields.Item("DocDate").Value.ToString("yyyyMMdd");
                                    _DueDate = _RSAPInvoice.Fields.Item("DocDueDate").Value.ToString("yyyyMMdd");
                                    _ConSignmentType = _RSAPInvoice.Fields.Item("U_ConType").Value.ToString();
                                    _CountryCode = _RSAPInvoice.Fields.Item("U_CountryC").Value.ToString();
                                    _CountryName = _RSAPInvoice.Fields.Item("U_CountryN").Value.ToString();
                                    _DocumentType = "APInvoice";
                                    _DocTotal = Convert.ToDouble(_RSAPInvoice.Fields.Item("DocTotalFC").Value.ToString()) - Convert.ToDouble(_RSAPInvoice.Fields.Item("VatSumFC").Value.ToString());
                                }
                            }
                            else
                            {
                                _DocEntry = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString();
                                _DocNum = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocNum", 0).ToString();
                                _BPCode = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("CardCode", 0).ToString();
                                _BPName = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("CardName", 0).ToString();
                                _DocCurency = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocCur", 0).ToString();
                                _Branch = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("BPLId", 0).ToString();
                                _PostingDate = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocDate", 0).ToString();
                                _DueDate = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocDueDate", 0).ToString();
                                _ConSignmentType = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("U_ConType", 0).ToString();
                                _CountryCode = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("U_CountryC", 0).ToString();
                                _CountryName = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("U_CountryN", 0).ToString();
                                _DocumentType = "GRPO";
                                _DocTotal = Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocTotalFC", 0).ToString()) - Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("VatSumFC", 0).ToString());
                            }
                            this._Object = new ClsLandedCost();
                            this.OpenChildForm(Constants.Forms.LANDED_COST, ref pVal, true);
                            ((ClsLandedCost)_Object).FormDefault(_DocEntry, _DocNum, _BPCode, _BPName, _DocCurency, _Branch, _PostingDate, _DueDate, _DocTotal, _ConSignmentType, _CountryCode, _CountryName, _DocumentType);
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (pVal.ItemUID == "CountryC")
                        {
                            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("CountryC").Specific;
                            _EditText = (SAPbouiCOM.EditText)_Form.Items.Item("CountryN").Specific;
                            _EditText.Value = _ComboBox.Selected.Description.ToString();

                        }
                        break;

                    #endregion

                   
                }
            }

            #endregion

        }
        #endregion

        #region Form Default
        internal void FormDefault()
        {
            try
            {
                string strSQL = "select \"U_LCostReq\" from OUSR where \"USER_CODE\" = '" + Utilities.Application.Company.UserName + "'";
                Utilities.ExecuteSQL(ref _RS, strSQL);
                string test = _RS.Fields.Item(0).Value.ToString().Trim();
                if (_RS.Fields.Item(0).Value.ToString().Trim() == "YES")
                {
                    _ExistItem = _Form.Items.Item("2");
                    _NewItem = _Form.Items.Add("BtnLandC", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
                    _NewItem.Width = _ExistItem.Width + 10;
                    _NewItem.Top = _ExistItem.Top;
                    _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                    _Button.Caption = "Landed Cost";
                }


                _ExistItem = _Form.Items.Item("70");
                _NewItem = _Form.Items.Add("CountryO", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Width = _ExistItem.Width - 5 ;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 45;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "Country Of Origin";

                _ExistItem = _Form.Items.Item("CountryO");
                _NewItem = _Form.Items.Add("CountryC", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.Top = _ExistItem.Top;
                _ComboBox = (SAPbouiCOM.ComboBox)_NewItem.Specific;
                _ComboBox.DataBind.SetBound(true, "OPDN", "U_CountryC");
                Utilities.FillCombo(ref _ComboBox, "select \"Code\",\"Name\" from OCRY");

                _ExistItem = _Form.Items.Item("CountryC");
                _NewItem = _Form.Items.Add("CountryN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.Top = _ExistItem.Top;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OPDN", "U_CountryN");

                /*---------------------------Start Cosing Folder--------------------------------*/

                _ExistItem = _Form.Items.Item("1320002137");
                _NewItem = _Form.Items.Add("FldCosting", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                _NewItem.Height = _ExistItem.Height;
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width;
                _NewItem.Width = _ExistItem.Width;
                _Folder = (SAPbouiCOM.Folder)_NewItem.Specific;
                _Folder.GroupWith("1320002137");
                _Folder.Caption = "Costing";

                _ExistItem = _Form.Items.Item("19");
                _NewItem = _Form.Items.Add("StInvVal", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top;
                _NewItem.Width = _ExistItem.Width + 100;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "Invoice Value";

                _ExistItem = _Form.Items.Item("StInvVal");
                _NewItem = _Form.Items.Add("TxtInvVal", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
                _NewItem.Top = _ExistItem.Top;
                _NewItem.Width = _ExistItem.Width;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OPDN", "U_InvVal");

                _ExistItem = _Form.Items.Item("StInvVal");
                _NewItem = _Form.Items.Add("StExRate", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "Exchange Rate";

                _ExistItem = _Form.Items.Item("TxtInvVal");
                _NewItem = _Form.Items.Add("TxtExRate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OPDN", "U_ExRate");

                _ExistItem = _Form.Items.Item("StExRate");
                _NewItem = _Form.Items.Add("StValNRS", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "Invoice Value(NPR)";

                _ExistItem = _Form.Items.Item("TxtExRate");
                _NewItem = _Form.Items.Add("TxtValNRS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OPDN", "U_ValNRS");

                _ExistItem = _Form.Items.Item("StValNRS");
                _NewItem = _Form.Items.Add("StImpDuty", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "Custom Duty";

                _ExistItem = _Form.Items.Item("TxtValNRS");
                _NewItem = _Form.Items.Add("TxtImpDuty", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OPDN", "U_ImpDuty");

                _ExistItem = _Form.Items.Item("StImpDuty");
                _NewItem = _Form.Items.Add("StECS", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "Excise Duty";

                _ExistItem = _Form.Items.Item("TxtImpDuty");
                _NewItem = _Form.Items.Add("TxtECS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OPDN", "U_ECSVal");

                _ExistItem = _Form.Items.Item("StECS");
                _NewItem = _Form.Items.Add("StVAT", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "VAT";

                _ExistItem = _Form.Items.Item("TxtECS");
                _NewItem = _Form.Items.Add("TxtVAT", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OPDN", "U_VatVal");

                _ExistItem = _Form.Items.Item("StVAT");
                _NewItem = _Form.Items.Add("StPragpn", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "PP Charges";

                _ExistItem = _Form.Items.Item("TxtVAT");
                _NewItem = _Form.Items.Add("TxtPragpn", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OPDN", "U_Pragyapan");

                _ExistItem = _Form.Items.Item("StPragpn");
                _NewItem = _Form.Items.Add("StDumrg", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "Dumurrage Charges";

                _ExistItem = _Form.Items.Item("TxtPragpn");
                _NewItem = _Form.Items.Add("TxtDumrg", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OPDN", "U_Dumurrage");

                _ExistItem = _Form.Items.Item("StDumrg");
                _NewItem = _Form.Items.Add("StFndReq", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "Total Fund Required";

                _ExistItem = _Form.Items.Item("TxtDumrg");
                _NewItem = _Form.Items.Add("txtFndReq", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.FromPane = 5;
                _NewItem.ToPane = 5;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.Width = _ExistItem.Width;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OPDN", "U_FndReq");

                //Utilities.Application.SBO_Application.Menus.Item("5909").String = "Test";

                /*---------------------------End Cosing Folder--------------------------------*/

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion

        public override void RightClick_Event(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {

            //if (eventInfo.BeforeAction)
            //{
            //    Utilities.Application.SBO_Application.Menus.Item("5909").String = "Test";

            //}
        }

    }
}
