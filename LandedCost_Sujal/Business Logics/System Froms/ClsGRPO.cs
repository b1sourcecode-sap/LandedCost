using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LandedCost
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
        string _DocEntry = "", _DocCurency = "", _DocNum = "", _BPCode = "", _BPName = "", _Branch = "", _PostingDate = "", _DueDate = "", _ConSignmentType = "", _CountryCode = "", _CountryName = "", _DocumentType = "",_DocType = "";
        double _DocTotal = 0;
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RSLandedCost = null, _Rs = null, _RSChk = null, _RsCal = null, _RSChkGetVal = null, _RSChk1 = null, _RSPriceList = null;

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
                        if (pVal.ItemUID == "BtnLandC" && _Form.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            _DocEntry = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString();
                            string _SQL = "select TOP 1 PDN1.\"TrgetEntry\" from OPDN inner join PDN1 ON OPDN.\"DocEntry\" = PDN1.\"DocEntry\" where OPDN.\"LndCstNum\" = 0 AND OPDN.\"DocStatus\" = 'C' and PDN1.\"TargetType\" = 18 and PDN1.\"LineStatus\" = 'C' and OPDN.\"DocEntry\" = '" + _DocEntry + "' ";
                            Utilities.ExecuteSQL(ref _RS, _SQL);
                            if (_RS.RecordCount > 0)
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
                                _DocTotal = Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocTotalFC", 0).ToString()) - Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("VatSumFC", 0).ToString());
                                _CountryCode = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("U_CountryC", 0).ToString();
                                _CountryName = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("U_CountryN", 0).ToString();
                                _DocumentType = "GRPO";

                                //string _SQLAPInvoice = "select \"DocEntry\",\"DocNum\",\"CardCode\",\"CardName\",\"DocCur\",\"BPLId\",\"DocDate\",\"DocDueDate\",\"DocTotalFC\",\"VatSumFC\",\"U_ConType\",\"U_CountryC\",\"U_CountryN\",'APInvoice' as \"DocumentType\" from OPCH WHERE OPCH.\"DocEntry\" = '" + _RS.Fields.Item("TrgetEntry").Value.ToString() + "' ";
                                //Utilities.ExecuteSQL(ref _RSAPInvoice, _SQLAPInvoice);
                                //if (_RS.RecordCount > 0)
                                //{
                                //    _DocEntry = _RSAPInvoice.Fields.Item("DocEntry").Value.ToString();
                                //    _DocNum = _RSAPInvoice.Fields.Item("DocNum").Value.ToString();
                                //    _BPCode = _RSAPInvoice.Fields.Item("CardCode").Value.ToString();
                                //    _BPName = _RSAPInvoice.Fields.Item("CardName").Value.ToString();
                                //    _DocCurency = _RSAPInvoice.Fields.Item("DocCur").Value.ToString();
                                //    _Branch = _RSAPInvoice.Fields.Item("BPLId").Value.ToString();
                                //    _PostingDate = _RSAPInvoice.Fields.Item("DocDate").Value.ToString("yyyyMMdd");
                                //    _DueDate = _RSAPInvoice.Fields.Item("DocDueDate").Value.ToString("yyyyMMdd");
                                //    _ConSignmentType = _RSAPInvoice.Fields.Item("U_ConType").Value.ToString();
                                //    _CountryCode = _RSAPInvoice.Fields.Item("U_CountryC").Value.ToString();
                                //    _CountryName = _RSAPInvoice.Fields.Item("U_CountryN").Value.ToString();
                                //    _DocumentType = "APInvoice";
                                //    _DocTotal = Convert.ToDouble(_RSAPInvoice.Fields.Item("DocTotalFC").Value.ToString()) - Convert.ToDouble(_RSAPInvoice.Fields.Item("VatSumFC").Value.ToString());
                                //}
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
                                _DocTotal = Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("DocTotalFC", 0).ToString()) - Convert.ToDouble(_Form.DataSources.DBDataSources.Item("OPDN").GetValue("VatSumFC", 0).ToString());
                                _CountryCode = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("U_CountryC", 0).ToString();
                                _CountryName = _Form.DataSources.DBDataSources.Item("OPDN").GetValue("U_CountryN", 0).ToString();
                                _DocumentType = "GRPO";
                            }
                            this._Object = new ClsLandedCost();
                            this.OpenChildForm(Constants.Forms.LANDED_COST, ref pVal, true);
                            ((ClsLandedCost)_Object).FormDefault(_DocEntry, _DocNum, _BPCode, _BPName, _DocCurency, _Branch, _PostingDate, _DueDate, _DocTotal, _ConSignmentType, _CountryCode, _CountryName, _DocumentType);
                        }

                        break;

                    #endregion

                    #region et_COMBO_SELECT

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
                _ExistItem = _Form.Items.Item("2");
                _NewItem = _Form.Items.Add("BtnLandC", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
                _NewItem.Width = _ExistItem.Width + 10;
                _NewItem.Top = _ExistItem.Top;
                _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                _Button.Caption = "Landed Cost";

                _ExistItem = _Form.Items.Item("70");
                _NewItem = _Form.Items.Add("CountryO", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Width = _ExistItem.Width - 5;
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


            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }

        #endregion

    }
}
