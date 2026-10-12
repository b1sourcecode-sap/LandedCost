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

namespace LandedCost
{
    class ClsItemMaster : Base
    {

         #region VARIABLE DECLARATION
        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        SAPbouiCOM.Button _Button;
        private bool flag;
        public SAPbouiCOM.Matrix _Matrix;
        SAPbouiCOM.StaticText _StaticText;
        string _DocEntry = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        private static HanaConnection _conn;
        private static HanaDataAdapter _da;
        private static DataTable _dataTable;
        private static HanaParameter _parm1;
        private static HanaParameter _parm2;
        private static int _errors = 0;
        string strcon, _Sql;
        private SAPbouiCOM.EditText _EditText;
        private SAPbouiCOM.ComboBox _ComboBox;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource;
        private SAPbobsCOM.Recordset _RS;
        private int RowID { get; set; }
        private string _Code { get; set; }
        private string _LineId { get; set; }
        private string _ActivityType = string.Empty;
        string _ItemCode = string.Empty, _ItemName = string.Empty;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsItemMaster()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsItemMaster()
        {
        }
        #endregion

        #region Item Event
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.Before_Action)
            {
                try
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                            if (pVal.ItemUID == "txtSbGrp1" && pVal.CharPressed == 9 && Form.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            {
                                _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                                _EditText = (SAPbouiCOM.EditText)_Form.Items.Item("txtSbGrp1").Specific;
                                if (_EditText.Value.ToString() == "" && pVal.ItemUID == "txtSbGrp1")
                                {
                                    DataSet ds = LoadCompanySetting();
                                    string ServerName = ds.Tables[0].Rows[0]["ServerName"].ToString().Trim();
                                    string UserID = ds.Tables[0].Rows[0]["UserID"].ToString().Trim();
                                    string Password = ds.Tables[0].Rows[0]["Password"].ToString().Trim();
                                    string CurrentSchema = ds.Tables[0].Rows[0]["CurrentSchema"].ToString().Trim();
                                    _conn = new HanaConnection("Server=" + ServerName + ";UserID=" + UserID + ";Password=" + Password + ";Current Schema=" + CurrentSchema + "");
                                    this._Object = new ClsChooseFromList();
                                    ((ClsChooseFromList)this._Object).FormTitle = "List of SubGroup1";
                                    ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref oBubbleEvent, "select \"U_SubGrpCodL1\",\"U_SubGrpNameL1\" from \"@ITM_SUBGROUPC1\" where \"U_Status\" = 'Y'", strcon, false);
                                }

                            }

                            if (pVal.ItemUID == "txtSbGrp2" && pVal.CharPressed == 9 && Form.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            {
                                _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                                _EditText = (SAPbouiCOM.EditText)_Form.Items.Item("txtSbGrp2").Specific;
                                if (_EditText.Value.ToString() == "" && pVal.ItemUID == "txtSbGrp2")
                                {
                                    DataSet ds = LoadCompanySetting();
                                    string ServerName = ds.Tables[0].Rows[0]["ServerName"].ToString().Trim();
                                    string UserID = ds.Tables[0].Rows[0]["UserID"].ToString().Trim();
                                    string Password = ds.Tables[0].Rows[0]["Password"].ToString().Trim();
                                    string CurrentSchema = ds.Tables[0].Rows[0]["CurrentSchema"].ToString().Trim();
                                    _conn = new HanaConnection("Server=" + ServerName + ";UserID=" + UserID + ";Password=" + Password + ";Current Schema=" + CurrentSchema + "");
                                    this._Object = new ClsChooseFromList();
                                    ((ClsChooseFromList)this._Object).FormTitle = "List of SubGroup1";
                                    ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref oBubbleEvent, "select \"U_SubGrpCodL2\",\"U_SubGrpNameL2\" from \"@ITM_SUBGROUP2\" where \"U_Status\" = 'Y' and \"U_SubGrpCodL1\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtSbGrp1").Specific).Value.ToString() + "'", strcon, false);
                                }

                            }

                            break;
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
            #endregion

            #region After Action
            else
            {
                try
                {
                    switch (pVal.EventType)
                    {

                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            if (pVal.ItemUID == "btnDtSetup")
                            {
                                _ItemCode = Form.DataSources.DBDataSources.Item("OITM").GetValue("ItemCode", 0).ToString().Trim();
                                _ItemName = Form.DataSources.DBDataSources.Item("OITM").GetValue("ItemName", 0).ToString();
                                this._Object = new ClsDutySetup();
                                Utilities.LoadForm(ref this._Object, Constants.Forms.DUTY_SETUP);
                                ((ClsDutySetup)this._Object).FormDefault(_ItemCode, _ItemName);
                            }
                            break;
                            
                            
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
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
                _MasterDBDataSource = Form.DataSources.DBDataSources.Item("OITM");
                _ExistItem = _Form.Items.Item("148");
                _NewItem = _Form.Items.Add("LblExDuty", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.FromPane = 1;
                _NewItem.ToPane = 1;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "Excise Duty";

                _ExistItem = _Form.Items.Item("149");
                _NewItem = _Form.Items.Add("TxtExDuty", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.FromPane = 1;
                _NewItem.ToPane = 1;
                _NewItem.Height = _ExistItem.Height;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OITM", "U_ExDPer");

                _ExistItem = _Form.Items.Item("150");
                _NewItem = _Form.Items.Add("LblPer", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.FromPane = 1;
                _NewItem.ToPane = 1;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "%";

                //------------------------------------------//
                _ExistItem = _Form.Items.Item("LblExDuty");
                _NewItem = _Form.Items.Add("LblVtDuty", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.FromPane = 1;
                _NewItem.ToPane = 1;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "VAT Duty";

                _ExistItem = _Form.Items.Item("TxtExDuty");
                _NewItem = _Form.Items.Add("TxtVtDuty", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.FromPane = 1;
                _NewItem.ToPane = 1;
                _NewItem.Height = _ExistItem.Height;
                _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
                _EditText.DataBind.SetBound(true, "OITM", "U_VATDPer");

                _ExistItem = _Form.Items.Item("LblPer");
                _NewItem = _Form.Items.Add("LblPerVT", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height + 2;
                _NewItem.FromPane = 1;
                _NewItem.ToPane = 1;
                _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
                _StaticText.Caption = "%";

                _ExistItem = _Form.Items.Item("2");
                _NewItem = _Form.Items.Add("btnDtSetup", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 400;
                _NewItem.Width = _ExistItem.Width + 50;
                _NewItem.Top = _ExistItem.Top;
                _NewItem.FromPane = _ExistItem.FromPane;
                _NewItem.ToPane = _ExistItem.ToPane;
                _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                _Button.Caption = "Item Duty Setup";

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion

        #region LOAD COMPANY SETTINGS
        public static DataSet LoadCompanySetting()
        {
            XmlDataDocument xmlDoc = new XmlDataDocument();
            DataSet dataSet = new DataSet();
            dataSet.ReadXml("AppConfig.xml");
            //dataSet.ReadXml("App.Config");
            return dataSet;

        }
        #endregion

        #region CHOOSE FROM LIST EVENT

        public override void ChooseFromList_Event()
        {
            try
            {
                _Form.Freeze(true);
                if (this._CFLEvent.ItemUID == "txtSbGrp1")
                {
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtSbGrp1").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                  
                }
                if (this._CFLEvent.ItemUID == "txtSbGrp2")
                {
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtSbGrp2").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();

                }
               
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.ToString());
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        #endregion


    }
}
