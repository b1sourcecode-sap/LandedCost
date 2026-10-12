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
using System.Configuration;
using CrystalDecisions.CrystalReports.ViewerObjectModel;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Drawing.Printing;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;


namespace LandedCost_SujalFoods
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
        public static string _DocEntry = "", _DocCurency = "", _DocNum = "", _BPCode = "", _BPName = "", _Branch = "", _PostingDate = "", _DueDate = "", _ConSignmentType = "", _CountryCode = "", _CountryName = "", _DocumentType;
        public static double _DocTotal = 0;
        public static string LandedCostNo = "";
        public static string Remarks = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        SAPbouiCOM.ComboBox _ComboBox;
        private SAPbouiCOM.DBDataSource _DBDataSource;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RSLandedCost = null, _Rs = null, _RSChk = null, _RsCal = null, _RSChkGetVal = null, _RSSeries = null, _RsBP = null, _RSLandedC = null, _RSUpdate = null;
        string SQL1 = string.Empty, SQL = string.Empty, StrLandedCostSQL = string.Empty;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource1, _ChildDBDataSource2, _ChildDBDataSource3;
        SAPbobsCOM.UserTable UDTHeader, UDTLine1, UDTLine2;
        public static int PvalCount = 0;
        public static double SubTotal1 = 0;
        public static double SubTotal2 = 0;
        public static double ExciseDuty = 0;
        public static double CustomDuty = 0;
        SAPbouiCOM.ComboBoxColumn _ComboBoxColumn;
        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;
        string _strDocCurency = string.Empty;
        SAPbouiCOM.CommonSetting _CommonSetting = null;
        public static int count = 1;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsLandedCost()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsLandedCost()
        {
        }
        #endregion

        #region Item Events

        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.Before_Action)
            {
                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        try
                        {
                            if (pVal.ColUID == "RowsHeader" && pVal.ItemUID == "405")
                            {
                                _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                                _Grid.CommonSetting.SetRowBackColor(count, -1);//-41
                                count = pVal.Row + 1;
                                _Grid.CommonSetting.SetRowBackColor(pVal.Row + 1, 3198718);

                            }
                            if (pVal.ColUID == "RowsHeader" && pVal.ItemUID == "404")
                            {
                                _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
                                _Grid.CommonSetting.SetRowBackColor(count, -1);//-41
                                count = pVal.Row + 1;
                                _Grid.CommonSetting.SetRowBackColor(pVal.Row + 1, 3198718);

                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        break;

                    #region et_CHOOSE_FROM_LIST

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.ItemUID == "41")
                        {
                            this.FilteGLAccount("");
                        }
                        if (pVal.ItemUID == "52")
                        {
                            this.FilteGLAccount1("");
                        }
                        break;

                    #endregion

                    #region et_ITEM_PRESSED Before action

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            for (int i = 0; i < _Grid.Rows.Count; i++)
                            {
                                if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(i).Value == "")
                                {
                                    Utilities.Message("You have not calculate the cost please do the same....", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    oBubbleEvent = false;
                                    return;
                                }

                            }
                        }
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value == "")
                            {
                                Utilities.Message("Custom Office account can not left blank.......", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        //if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        //{
                        //    if (((SAPbouiCOM.EditText)_Form.Items.Item("52").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Control account can not left blank.......", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //}
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("54").Specific).Value == "")
                            {
                                Utilities.Message("PP No. can not left blank.......", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("49").Specific).Value == "" || ((SAPbouiCOM.EditText)_Form.Items.Item("50").Specific).Value == "")
                            {
                                Utilities.Message("Country of origin can not left blank.......", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        if (pVal.ItemUID == "36")
                        {
                            _Form.Freeze(true);
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _Grid.DataTable.Columns.Item("Amount").Cells.Item(1).Value = "";
                            for (int i = 3; i < _Grid.Rows.Count; i++)
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(i).Value = "";

                            }
                            _Form.Freeze(false);
                        }

                        if (pVal.ItemUID == "21")
                        {
                            _Form.PaneLevel = 1;
                        }
                        else if (pVal.ItemUID == "22")
                        {
                            _Form.PaneLevel = 2;
                        }
                        else if (pVal.ItemUID == "23")
                        {
                            _Form.PaneLevel = 3;
                        }

                        if (pVal.ItemUID == "1" && (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                        {
                            try
                            {

                                Utilities.Message("Please Wait Operation is in process....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                if (!Utilities.Application.Company.InTransaction)
                                    Utilities.Application.Company.StartTransaction();

                                if (SaveData() == false)
                                {
                                    oBubbleEvent = false;
                                    return;
                                }

                                if (_DocumentType == "GRPO")
                                {
                                    string strQuery1 = "update \"OPDN\" set \"U_PPNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("54").Specific).Value.ToString().Trim() + "' where \"DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value.ToString().Trim() + "' and \"DocEntry\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("1000003").Specific).Value.ToString().Trim() + "'";
                                    Utilities.ExecuteSQL(ref _RSUpdate, strQuery1);
                                }
                                else if (_DocumentType == "APInvoice")
                                {
                                    string strQuery1 = "update \"OPCH\" set \"U_PPNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("54").Specific).Value.ToString().Trim() + "' where \"DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value.ToString().Trim() + "' and \"DocEntry\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("1000003").Specific).Value.ToString().Trim() + "'";
                                    Utilities.ExecuteSQL(ref _RSUpdate, strQuery1);
                                }

                                if (Utilities.Application.Company.InTransaction)
                                    Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                                Utilities.Message("Operation completed sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }

                        break;
                    #endregion

                    #region et_VALIDATE

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemUID == "404" && pVal.ColUID == "IDB Value" && pVal.InnerEvent == false)
                        {
                            _Form.Freeze(true);

                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
                            _Grid.DataTable.Columns.Item("Custom Duty").Cells.Item(pVal.Row).Value = Math.Ceiling((Convert.ToDouble(_Grid.DataTable.Columns.Item("IDB Value").Cells.Item(pVal.Row).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Custom Duty %").Cells.Item(pVal.Row).Value)) / 100);
                            _Grid.DataTable.Columns.Item("Base Value For Excise Duty").Cells.Item(pVal.Row).Value = (Convert.ToDouble(_Grid.DataTable.Columns.Item("IDB Value").Cells.Item(pVal.Row).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Custom Duty").Cells.Item(pVal.Row).Value));
                            _Grid.DataTable.Columns.Item("Excise Duty").Cells.Item(pVal.Row).Value = Math.Ceiling(((Convert.ToDouble(_Grid.DataTable.Columns.Item("IDB Value").Cells.Item(pVal.Row).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Custom Duty").Cells.Item(pVal.Row).Value)) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Excise Duty %").Cells.Item(pVal.Row).Value)) / 100);
                            _Grid.DataTable.Columns.Item("Base Value For VAT").Cells.Item(pVal.Row).Value = (Convert.ToDouble(_Grid.DataTable.Columns.Item("Excise Duty").Cells.Item(pVal.Row).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Base Value For Excise Duty").Cells.Item(pVal.Row).Value));
                            _Grid.DataTable.Columns.Item("VAT Amount").Cells.Item(pVal.Row).Value = Math.Ceiling((Convert.ToDouble(_Grid.DataTable.Columns.Item("Base Value For VAT").Cells.Item(pVal.Row).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("VAT %").Cells.Item(pVal.Row).Value))/100);

                            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("IDB Value")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Base Value For VAT")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Base Value For Excise Duty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Custom Duty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Excise Duty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("LineTotal")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("VAT Amount")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                            _Grid.SetCellFocus(pVal.Row, 14);
                            _Grid.SetCellFocus(pVal.Row, 8);

                            _Form.Freeze(false);
                        }
                        if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.InnerEvent == false)
                        {
                            _Form.Freeze(true);
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(pVal.Row).Value == "")
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(pVal.Row).Value = "0.00";
                            _Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(pVal.Row).Value = (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(pVal.Row).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(pVal.Row).Value)).ToString();
                            _Form.Freeze(false);
                        }
                        break;

                    #endregion

                    #region et_KEY_DOWN

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:

                        if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 0 && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 9 && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 10 && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 11 && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 12 && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 13 && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        //else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 14 && pVal.CharPressed != 9)
                        //{
                        //    oBubbleEvent = false;
                        //    return;
                        //}
                        //else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 15 && pVal.CharPressed != 9)
                        //{
                        //    oBubbleEvent = false;
                        //    return;
                        //}
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 36 && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 37 && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 38 && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Exchange Rate" && pVal.Row == 1 && pVal.CharPressed != 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Currency").Cells.Item(1).Value == "")
                            {
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Exchange Rate" && pVal.Row == 2 && pVal.CharPressed != 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Currency").Cells.Item(2).Value == "")
                            {
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Exchange Rate" && pVal.Row == 3 && pVal.CharPressed != 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Currency").Cells.Item(3).Value == "")
                            {
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Exchange Rate" && pVal.Row == 4 && pVal.CharPressed != 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Currency").Cells.Item(4).Value == "")
                            {
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Exchange Rate" && pVal.Row == 5 && pVal.CharPressed != 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Currency").Cells.Item(5).Value == "")
                            {
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Exchange Rate" && pVal.Row == 6 && pVal.CharPressed != 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Currency").Cells.Item(6).Value == "")
                            {
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Exchange Rate" && pVal.Row == 7 && pVal.CharPressed != 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Currency").Cells.Item(7).Value == "")
                            {
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Exchange Rate" && pVal.Row == 8 && pVal.CharPressed != 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Currency").Cells.Item(8).Value == "")
                            {
                                oBubbleEvent = false;
                                return;
                            }
                        }

                        break;

                    #endregion

                }
            }
            #endregion

            #region After Action
            else
            {
                switch (pVal.EventType)
                {

                    #region et_COMBO_SELECT

                    case BoEventTypes.et_COMBO_SELECT:
                        if (pVal.ItemUID == "405" && pVal.ColUID == "Currency" && pVal.Row == 1)
                        {

                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _strDocCurency = string.Empty;
                            _Grid.Columns.Item("Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            _strDocCurency = _ComboBoxColumn.GetSelectedValue(pVal.Row).Value.ToString();
                            string _SQL = "select \"Rate\" FROM \"ORTT\" where \"RateDate\" = '" + _PostingDate + "' and \"Currency\" = '" + _strDocCurency + "'";
                            Utilities.ExecuteSQL(ref _Rs, _SQL);
                                if (_strDocCurency != "NPR")
                                {
                                    _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(1).Value = _Rs.Fields.Item(0).Value.ToString();
                                }
                                else
                                {
                                    _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(1).Value = 1;
                                }
                        }
                        if (pVal.ItemUID == "405" && pVal.ColUID == "Currency" && pVal.Row == 2)
                        {

                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _strDocCurency = string.Empty;
                            _Grid.Columns.Item("Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            _strDocCurency = _ComboBoxColumn.GetSelectedValue(pVal.Row).Value.ToString();
                            string _SQL = "select \"Rate\" FROM \"ORTT\" where \"RateDate\" = '" + _PostingDate + "' and \"Currency\" = '" + _strDocCurency + "'";
                            Utilities.ExecuteSQL(ref _Rs, _SQL);
                            if (_strDocCurency != "NPR")
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(2).Value = _Rs.Fields.Item(0).Value.ToString();
                            }
                            else
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(2).Value = 1;
                            }
                        }
                        if (pVal.ItemUID == "405" && pVal.ColUID == "Currency" && pVal.Row == 3)
                        {

                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _strDocCurency = string.Empty;
                            _Grid.Columns.Item("Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            _strDocCurency = _ComboBoxColumn.GetSelectedValue(pVal.Row).Value.ToString();
                            string _SQL = "select \"Rate\" FROM \"ORTT\" where \"RateDate\" = '" + _PostingDate + "' and \"Currency\" = '" + _strDocCurency + "'";
                            Utilities.ExecuteSQL(ref _Rs, _SQL);
                            if (_strDocCurency != "NPR")
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(3).Value = _Rs.Fields.Item(0).Value.ToString();
                            }
                            else
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(3).Value = 1;
                            }
                        }
                        if (pVal.ItemUID == "405" && pVal.ColUID == "Currency" && pVal.Row == 4)
                        {

                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _strDocCurency = string.Empty;
                            _Grid.Columns.Item("Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            _strDocCurency = _ComboBoxColumn.GetSelectedValue(pVal.Row).Value.ToString();
                            string _SQL = "select \"Rate\" FROM \"ORTT\" where \"RateDate\" = '" + _PostingDate + "' and \"Currency\" = '" + _strDocCurency + "'";
                            Utilities.ExecuteSQL(ref _Rs, _SQL);
                            if (_strDocCurency != "NPR")
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(4).Value = _Rs.Fields.Item(0).Value.ToString();
                            }
                            else
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(4).Value = 1;
                            }
                        }
                        if (pVal.ItemUID == "405" && pVal.ColUID == "Currency" && pVal.Row == 5)
                        {

                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _strDocCurency = string.Empty;
                            _Grid.Columns.Item("Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            _strDocCurency = _ComboBoxColumn.GetSelectedValue(pVal.Row).Value.ToString();
                            string _SQL = "select \"Rate\" FROM \"ORTT\" where \"RateDate\" = '" + _PostingDate + "' and \"Currency\" = '" + _strDocCurency + "'";
                            Utilities.ExecuteSQL(ref _Rs, _SQL);
                            if (_strDocCurency != "NPR")
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(5).Value = _Rs.Fields.Item(0).Value.ToString();
                            }
                            else
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(5).Value = 1;
                            }
                        }
                        if (pVal.ItemUID == "405" && pVal.ColUID == "Currency" && pVal.Row == 6)
                        {

                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _strDocCurency = string.Empty;
                            _Grid.Columns.Item("Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            _strDocCurency = _ComboBoxColumn.GetSelectedValue(pVal.Row).Value.ToString();
                            string _SQL = "select \"Rate\" FROM \"ORTT\" where \"RateDate\" = '" + _PostingDate + "' and \"Currency\" = '" + _strDocCurency + "'";
                            Utilities.ExecuteSQL(ref _Rs, _SQL);
                            if (_strDocCurency != "NPR")
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(6).Value = _Rs.Fields.Item(0).Value.ToString();
                            }
                            else
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(6).Value = 1;
                            }
                        }
                        if (pVal.ItemUID == "405" && pVal.ColUID == "Currency" && pVal.Row == 7)
                        {

                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _strDocCurency = string.Empty;
                            _Grid.Columns.Item("Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            _strDocCurency = _ComboBoxColumn.GetSelectedValue(pVal.Row).Value.ToString();
                            string _SQL = "select \"Rate\" FROM \"ORTT\" where \"RateDate\" = '" + _PostingDate + "' and \"Currency\" = '" + _strDocCurency + "'";
                            Utilities.ExecuteSQL(ref _Rs, _SQL);
                            if (_strDocCurency != "NPR")
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(7).Value = _Rs.Fields.Item(0).Value.ToString();
                            }
                            else
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(7).Value = 1;
                            }
                        }
                        if (pVal.ItemUID == "405" && pVal.ColUID == "Currency" && pVal.Row == 8)
                        {

                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _strDocCurency = string.Empty;
                            _Grid.Columns.Item("Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            _strDocCurency = _ComboBoxColumn.GetSelectedValue(pVal.Row).Value.ToString();
                            string _SQL = "select \"Rate\" FROM \"ORTT\" where \"RateDate\" = '" + _PostingDate + "' and \"Currency\" = '" + _strDocCurency + "'";
                            Utilities.ExecuteSQL(ref _Rs, _SQL);
                            if (_strDocCurency != "NPR")
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(8).Value = _Rs.Fields.Item(0).Value.ToString();
                            }
                            else
                            {
                                _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(8).Value = 1;
                            }
                        }
                       

                        break;

                    #endregion

                    #region Set CFL Values
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "41" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_CRAccount", 0, _SysCFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString().Trim());

                        }
                        if (pVal.ItemUID == "52" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_CNTRAcont", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());

                        }
                        break;
                    #endregion

                    #region et_ITEM_PRESSED
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "46")
                        {
                            this._Object = new ClsCostingSheet();
                            this.OpenChildForm(Constants.Forms.COSTING_SHEET_REGISTER, ref pVal, true);
                            string _LandedCostNum = "";
                            _LandedCostNum = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString();
                            ((ClsCostingSheet)_Object).FormDefault(_LandedCostNum);

                        }
                        if (pVal.ItemUID == "47")
                        {
                            Thread myThread = new Thread(new ThreadStart(ShowReport));
                            myThread.SetApartmentState(ApartmentState.STA);
                            myThread.Start();

                        }
                        if (pVal.ItemUID == "1" && (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                        {  
                            _Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            Utilities.Message("Operation completed sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        if (pVal.ItemUID == "37" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {

                            

                            if (Utilities.Application.SBO_Application.MessageBox("Are you sure want to post landed cost document?", 1, "OK", "Cancel") == 1)
                            {
                                try
                                {
                                    Utilities.Message("Please Wait Operation is in process....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                    if (!Utilities.Application.Company.InTransaction)
                                        Utilities.Application.Company.StartTransaction();

                                    landedcosts(_BPCode, Convert.ToInt32(_DocEntry.ToString()), Utilities.ToDate(_DueDate.ToString()), Utilities.ToDate(_PostingDate.ToString()), _DocumentType);
                                    GeneralEntry(LandedCostNo);

                                    if (Utilities.Application.Company.InTransaction)
                                        Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                                    Utilities.Message("Operation completed sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                                catch (Exception ex)
                                {
                                    if (Utilities.Application.Company.InTransaction)
                                        Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    throw ex;
                                }
                               
                            }
                            else
                            {
                                oBubbleEvent = false;
                                return;
                            }

                        }
                        break;
                    #endregion

                    #region et_KEY_DOWN
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 1 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(1).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(1).Value = "0.00";
                            }

                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 2 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(2).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(2).Value = "0.00";
                            }

                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 3 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(3).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(3).Value = "0.00";
                            }

                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 4 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(4).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(4).Value = "0.00";
                            }
                            
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 5 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(5).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(5).Value = "0.00";
                            }

                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 6 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(6).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(6).Value = "0.00";
                            }

                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 7 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(7).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(7).Value = "0.00";
                            }
                            
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 8 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(8).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(8).Value = "0.00";
                            }
                            double CIFIndia = 0;
                            CIFIndia = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(1).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(1).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(2).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(2).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(3).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(3).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(4).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(4).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(5).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(5).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(6).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(6).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(7).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(7).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(8).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(8).Value);
                            _Grid.DataTable.Columns.Item("Amount").Cells.Item(9).Value = Math.Round(CIFIndia, 2).ToString();
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 9 && pVal.CharPressed == 9)
                        {
                            _Form.Freeze(true);
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(9).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(9).Value = "0.00";
                            }
                            double IDBValue = 0;
                            double SubTotalIDBValue = 0;
                            double ValForAddCost = 0;
                            double ValForInvoiceValue = 0;
                            double AdditionalCost = 0;
                            ValForInvoiceValue = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(0).Value);
                            ValForAddCost = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(1).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(2).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(3).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(4).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(5).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(6).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(7).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(8).Value);
                            AdditionalCost = ValForAddCost / ValForInvoiceValue;
                            _Form.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
                            for (int i = 0; i < _Grid.Rows.Count; i++)
                            {
                                IDBValue = Convert.ToDouble(_Grid.DataTable.Columns.Item("IDB Value").Cells.Item(i).Value);
                                SubTotalIDBValue = IDBValue + SubTotalIDBValue;

                                _Grid.DataTable.Columns.Item("Additional Cost").Cells.Item(i).Value = AdditionalCost * Convert.ToDouble(_Grid.DataTable.Columns.Item("LineTotal").Cells.Item(i).Value);
                            }
                            _Form.Items.Item("22").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _Grid.DataTable.Columns.Item("Amount").Cells.Item(10).Value = Math.Round(SubTotalIDBValue, 2).ToString();
                            _Grid.SetCellFocus(10, 5);
                            _Form.Freeze(false);

                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 10 && pVal.CharPressed == 9)
                        {
                            _Form.Freeze(true);
                            CustomDuty = 0;
                            SubTotal1 = 0;
                            CustomDuty = 0;
                            _Form.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
                            for (int i = 0; i < _Grid.Rows.Count; i++)
                            {
                                CustomDuty = Convert.ToDouble(_Grid.DataTable.Columns.Item("Custom Duty").Cells.Item(i).Value);
                                SubTotal1 = CustomDuty + SubTotal1;
                            }
                         
                            _Grid.SetCellFocus(0, 7);
                            _Form.Items.Item("22").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _Grid.DataTable.Columns.Item("Amount").Cells.Item(11).Value = Math.Round(SubTotal1, 2).ToString();
                            _Grid.SetCellFocus(11, 5);
                            _Form.Freeze(false);

                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 11 && pVal.CharPressed == 9)
                        {
                            _Form.Freeze(true);
                            ExciseDuty = 0;
                            SubTotal2 = 0;
                            double SubTotalTemp2 = 0;
                            _Form.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
                            for (int i = 0; i < _Grid.Rows.Count; i++)
                            {
                                ExciseDuty = Convert.ToDouble(_Grid.DataTable.Columns.Item("Excise Duty").Cells.Item(i).Value);
                                SubTotal2 = ExciseDuty + SubTotal2;
                            }
                          
                            _Grid.SetCellFocus(0, 11);
                            _Form.Items.Item("22").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value = Math.Round(SubTotal2, 2).ToString();
                            _Grid.SetCellFocus(12,5);
                            _Form.Freeze(false);

                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 12 && pVal.CharPressed == 9)
                        {
                            _Form.Freeze(true);
                            ExciseDuty = 0;
                            double SubTotal3 = 0;
                            double SubTotalTemp3 = 0;
                            _Form.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
                            for (int i = 0; i < _Grid.Rows.Count; i++)
                            {
                                SubTotalTemp3 = Convert.ToDouble(_Grid.DataTable.Columns.Item("VAT Amount").Cells.Item(i).Value);
                                SubTotal3 = SubTotal3 + SubTotalTemp3;
                            }
                            _Grid.SetCellFocus(0, 11);
                            _Form.Items.Item("22").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                            _Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value = Math.Round(SubTotal3, 2).ToString();
                            _Grid.SetCellFocus(13, 5);
                            _Form.Freeze(false);

                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 13 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value = "0.00";
                            }
                            string strQuery = "SELECT \"U_VATPer\" FROM \"@CSTMDSTPH\" ";
                            Utilities.ExecuteSQL(ref _RS, strQuery);
                            double VatAmount = 0, VatAmountCon = 0;
                            VatAmount = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(11).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(10).Value);
                            VatAmountCon = (VatAmount * Convert.ToDouble(_RS.Fields.Item("U_VATPer").Value)) / 100;
                            _Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value = Math.Round(VatAmountCon - Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value), 2).ToString();
                        }

                      
                        //else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 12 && pVal.CharPressed == 9)
                        //{
                        //    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value == "")
                        //    {
                        //        _Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value = "0.00";
                        //    }
                        //    double VatAmount = 0, VatAmountCon = 0;
                        //    VatAmount = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(11).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(10).Value);// + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(9).Value)
                        //    VatAmountCon = (VatAmount * 13) / 100;
                        //    _Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value = Math.Round(VatAmountCon, 2).ToString();
                        //}
                       
                        //else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 13 && pVal.CharPressed == 9)
                        //{
                        //    _Form.Freeze(true);
                        //    ExciseDuty = 0;
                        //    double SubTotal3 = 0;
                        //    double SubTotalTemp3 = 0;
                        //    _Form.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //    _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
                        //    for (int i = 0; i < _Grid.Rows.Count; i++)
                        //    {
                        //        SubTotalTemp3 = Convert.ToDouble(_Grid.DataTable.Columns.Item("VAT Amount").Cells.Item(i).Value);
                        //        SubTotal3 = SubTotal3 + SubTotalTemp3;

                        //    }
                        //    _Grid.SetCellFocus(0, 11);
                        //    _Form.Items.Item("22").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //    _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                        //    _Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value = Math.Round(SubTotal3, 2).ToString();
                        //    _Grid.SetCellFocus(14, 5);
                        //    _Form.Freeze(false);

                         
                        //}
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 14 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value = "0.00";
                            }
                            //double TotalDutyBeforeCustom = 0;
                            //double VATAmount = 0;
                           
                            //VATAmount = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value) - Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value);
                            //_Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value = Math.Round(VATAmount, 2).ToString();
                            //_Grid.SetCellFocus(15, 5);
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 15 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(15).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(15).Value = "0.00";
                            }
                            double TotalDuctyBeforeCustom = 0;
                            TotalDuctyBeforeCustom = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(11).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value) + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(15).Value);
                            _Grid.DataTable.Columns.Item("Amount").Cells.Item(16).Value = Math.Round(TotalDuctyBeforeCustom, 2).ToString();
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 16 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(16).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(16).Value = "0.00";
                            }
                           
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 17 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(17).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(17).Value = "0.00";
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 18 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(18).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(18).Value = "0.00";
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 19 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(19).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(19).Value = "0.00";
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 20 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(20).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(20).Value = "0.00";
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 21 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(21).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(21).Value = "0.00";
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 22 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(22).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(22).Value = "0.00";
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 23 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(23).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(23).Value = "0.00";
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 24 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(24).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(24).Value = "0.00";
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 25 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(25).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(25).Value = "0.00";
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 26 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(26).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(26).Value = "0.00";
                            }
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 27 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(27).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(27).Value = "0.00";
                            }
                            double TotalAdditionalCostWithVAT = 0;

                            TotalAdditionalCostWithVAT = (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(9).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(11).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value)) + 
                                (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(15).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(17).Value)) +
                                (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(18).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(19).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(20).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(21).Value)) +
                                (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(22).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(23).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(24).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(25).Value)) +
                                (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(26).Value)) + (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(27).Value));
                       
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(28).Value = Math.Round(TotalAdditionalCostWithVAT, 2).ToString();
                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 28 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(28).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(28).Value = "0.00";
                            }
                            double TotalLandedCost  = 0;

                            TotalLandedCost = (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(28).Value)) - (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value));
                            _Grid.DataTable.Columns.Item("Amount").Cells.Item(29).Value = Math.Round(TotalLandedCost, 2).ToString();

                        }
                        else if (pVal.ItemUID == "405" && pVal.ColUID == "Amount" && pVal.Row == 29 && pVal.CharPressed == 9)
                        {
                            if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(29).Value == "")
                            {
                                _Grid.DataTable.Columns.Item("Amount").Cells.Item(29).Value = "0.00";
                            }
                            double LandingFactor  = 0;

                            LandingFactor = (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(29).Value)) / (Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(0).Value));
                            _Grid.DataTable.Columns.Item("Amount").Cells.Item(30).Value = Math.Round(LandingFactor, 2).ToString();

                        }
                        break;
                    #endregion

                }
            }

            #endregion
        }
        
        #endregion

        #region Form Default

        internal void FormDefault(string DocEntry, string DocNum, string BPCode, string BPName, string DocCurency, string Branch, string PostingDate, string DueDate, double DocTotal, string ConSignmentType,string CountryCode,string CountryName,string DocumentType)
        {
            _DocEntry = DocEntry;
            _DocCurency = DocCurency;
            _DocNum = DocNum;
            _BPCode = BPCode;
            _BPName = BPName;
            _Branch = Branch;
            _PostingDate = PostingDate;
            _DueDate = DueDate;
            _DocTotal = DocTotal;
            _ConSignmentType = ConSignmentType;
            _CountryCode = CountryCode;
            _CountryName = CountryName;
            _DocumentType = DocumentType;

            _MasterDBDataSource = Form.DataSources.DBDataSources.Item("@LANDEDCOSTH");
            _ChildDBDataSource3 = Form.DataSources.DBDataSources.Item("@LANDEDCOSTL3");

            _Form.DataSources.DataTables.Add("dtReviewList");
            _Form.DataSources.DataTables.Add("dtReviewList1");

            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific;
            Utilities.FillCombo(ref _ComboBox, "SELECT \"CurrCode\",\"CurrName\" FROM \"OCRN\"");
            _ComboBox.Select(_DocCurency, SAPbouiCOM.BoSearchKey.psk_ByValue);

            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("8").Specific;
            Utilities.FillCombo(ref _ComboBox, "SELECT \"BPLId\",\"BPLName\" FROM \"OBPL\"");
            _ComboBox.Select(_Branch, SAPbouiCOM.BoSearchKey.psk_ByValue);

            string series = "";

            if (_DocumentType == "GRPO")
            {
                series = "select Top 1 \"Series\" from NNM1 where \"ObjectCode\" = 69 " +
                               " and \"Indicator\" = (select \"Indicator\" from OFPR  where (Select \"DocDate\" from OPDN where \"DocEntry\" = '" + _DocEntry + "') " +
                               " between \"F_RefDate\" and \"T_RefDate\") ";
            }
            else
            {
                series = "select Top 1 \"Series\" from NNM1 where \"ObjectCode\" = 69 " +
                                   " and \"Indicator\" = (select \"Indicator\" from OFPR  where (Select \"DocDate\" from OPCH where \"DocEntry\" = '" + _DocEntry + "') " +
                                   " between \"F_RefDate\" and \"T_RefDate\") ";
            }
            Utilities.ExecuteSQL(ref _RSSeries, series);
            ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = _RSSeries.Fields.Item("Series").Value.ToString().Trim();

            string _Sql = "SELECT \"U_DocNum\" FROM \"@LANDEDCOSTH\" where \"U_GRPODocEntry\" = '" + _DocEntry + "' AND \"U_GRPODocNum\" = '" + _DocNum + "'";
            Utilities.ExecuteSQL(ref _RS, _Sql);
            if (_RS.RecordCount > 0)
            {
                _Form.Freeze(true);
                FillHeaderDataFromUDT(_DocEntry);
                GRPODataLoadFromUDT(_DocEntry);
                LandedCostHeadsLoadFromUDT(_DocEntry);
                _Form.Freeze(false);
                _Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                _Form.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                string _Sql1 = "SELECT * from \"IPF1\" where \"BaseEntry\" = '" + _DocEntry + "' ";
                Utilities.ExecuteSQL(ref _RSChk, _Sql1);
                if (_RSChk.RecordCount > 0)
                {
                    _Form.Items.Item("36").Visible = false;
                    _Form.Items.Item("37").Visible = false;
                    _Form.Items.Item("405").Enabled = false;
                }
                
                
            }
            else
            {
                ((SAPbouiCOM.EditText)_Form.Items.Item("32").Specific).Value = _BPCode.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value = _BPName.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value = _DocNum.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = Utilities.getMaxColumnValueNum("@LANDEDCOSTH", "U_DocNum");
                ((SAPbouiCOM.EditText)_Form.Items.Item("1000003").Specific).Value = _DocEntry.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value = _PostingDate.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("16").Specific).Value = _DueDate.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("49").Specific).Value = _CountryCode.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("50").Specific).Value = _CountryName.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = _DocumentType.ToString();

                _Form.Items.Item("32").Enabled = false;
                _Form.Items.Item("4").Enabled = false;
                _Form.Items.Item("30").Enabled = false;
                _Form.Items.Item("10").Enabled = false;
                _Form.Items.Item("1000003").Enabled = false;
                _Form.Items.Item("32").Enabled = false;
                _Form.Items.Item("4").Enabled = false;
                _Form.Items.Item("1000001").Enabled = false;
                _Form.Items.Item("49").Enabled = false;
                _Form.Items.Item("50").Enabled = false;
                _Form.Items.Item("888").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("20").Enabled = false;
                _Form.Items.Item("8").Enabled = false;
                _Form.Items.Item("12").Enabled = false;
                _Form.Items.Item("18").Enabled = false;

                GRPODataLoadinItemsTab(_DocEntry);
                LandedCostHeadsLoadCostTab(_DocEntry);
                _Form.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

            }

        }
       
        #endregion

        #region GRPODataLoadinItemsTab

        internal void GRPODataLoadinItemsTab(string DocEntry)
        {
            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
            _GridDataTable = _Form.DataSources.DataTables.Item("dtReviewList");


            //SQL1 = "Select 0 as \"Code\",T0.\"ItemCode\",T0.\"Dscription\",T0.\"Quantity\",T0.\"Price\",(T0.\"LineTotal\" + T0.\"DistribSum\") as \"LineTotal\",0.00 as \"Additional Cost\",T2.\"TotalTax\" as \"Custom Duty %\" ,(((T0.\"LineTotal\" + T0.\"DistribSum\") * T2.\"TotalTax\")/100) as \"Custom Duty\", " +
            //        " ((((T0.\"LineTotal\" + T0.\"DistribSum\") * T2.\"TotalTax\")/100) + ((T0.\"LineTotal\" + T0.\"DistribSum\"))) as \"Base Value For Excise Duty\" ,T1.\"U_ExDPer\" as \"Excise Duty %\",  " +
            //         " ((((((T0.\"LineTotal\" + T0.\"DistribSum\") * T2.\"TotalTax\")/100) + ((T0.\"LineTotal\" + T0.\"DistribSum\"))) * T1.\"U_ExDPer\")/100) as \"Excise Duty\",  " +
            //         " ((((((T0.\"LineTotal\" + T0.\"DistribSum\") * T2.\"TotalTax\")/100) + ((T0.\"LineTotal\" + T0.\"DistribSum\"))) * T1.\"U_ExDPer\")/100) + ((((T0.\"LineTotal\" + T0.\"DistribSum\") * T2.\"TotalTax\")/100) + ((T0.\"LineTotal\" + T0.\"DistribSum\"))) as \"Base Value For VAT\"  " +
            //         ",T1.\"U_VATDPer\" as \"VAT %\",0.00 as \"VAT Amount\" " +
            //        " from \"PDN1\" T0 " +
            //         " inner join \"OITM\" T1 ON T0.\"ItemCode\" = T1.\"ItemCode\" " +
            //         " inner join \"OARG\" T2 ON T1.\"CstGrpCode\" = T2.\"CstGrpCode\" " +
            //        " where T0.\"DocEntry\" = '" + DocEntry + "'";

            SQL1 = " Select 0 as \"Code\",T0.\"ItemCode\",T0.\"Dscription\",T0.\"Quantity\",T0.\"Price\",(T0.\"LineTotal\" + T0.\"DistribSum\") as \"LineTotal\", " +
                      " 0.00 as \"Additional Cost\",0.00 as \"IDB Value\",IFNULL(T4.\"U_CusDutyP\",0) as \"Custom Duty %\" ,0.00 as \"Custom Duty\", " +
                      " 0.00 as \"Base Value For Excise Duty\" , " +
                      " IFNULL(T4.\"U_ExDutyP\",0) as \"Excise Duty %\",  " +
                      " 0.00 as \"Excise Duty\", " +
                      " 0.00 as \"Base Value For VAT\" " +
                      ",IFNULL(T4.\"U_VATP\",0) as \"VAT %\",0.00 as \"VAT Amount\" " +
                      " from \"PDN1\" T0 " +
                      " inner join \"OPDN\" T5 ON T0.\"DocEntry\" = T5.\"DocEntry\" " +
                      " inner join \"OITM\" T1 ON T0.\"ItemCode\" = T1.\"ItemCode\"  " +
                      " inner join \"OARG\" T2 ON T1.\"CstGrpCode\" = T2.\"CstGrpCode\" " +
                      " left outer join \"@ITEMDUTYSETUPM\" T3 ON T3.\"U_ItemCode\" = T1.\"ItemCode\" " +
                      " left outer join \"@ITEMDUTYSETUPC\" T4 ON T3.\"Code\" = T4.\"Code\" and T5.\"U_CountryC\" = T4.\"U_CountryC\" " +
                       " where T0.\"DocEntry\" = '" + DocEntry + "' order by T0.\"LineNum\" ";

            _GridDataTable.ExecuteQuery(SQL1);
            _Grid.DataTable = _GridDataTable;
            _RowHeaders = _Grid.RowHeaders;
            _RowHeaders.TitleObject.Caption = "#";

            for (int i = 0; i < _Grid.Rows.Count; i++)
            {
                _RowHeaders.SetText(i, (i + 1).ToString());

            }
            _Grid.Columns.Item("Code").Width = 0;
            _Column = (SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(1);
            _Column.LinkedObjectType = "4";

            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Base Value For VAT")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Base Value For Excise Duty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Custom Duty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Excise Duty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;

            _Grid.Columns.Item(0).Editable = false;
            _Grid.Columns.Item(1).Editable = false;
            _Grid.Columns.Item(2).Editable = false;
            _Grid.Columns.Item(3).Editable = false;
            _Grid.Columns.Item(4).Editable = false;
            _Grid.Columns.Item(5).Editable = false;

        }

        #endregion

        #region GRPODataLoadFromUDT

        internal void GRPODataLoadFromUDT(string DocEntry)
        {
            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
            _GridDataTable = _Form.DataSources.DataTables.Item("dtReviewList");

            SQL1 = "Select \"Code\", \"U_ItemCode\" as \"ItemCode\",\"U_ItemName\" as \"Dscription\",\"U_Quantity\" as \"Quantity\",\"U_Price\" as \"Price\",\"U_LineTotal\" as \"LineTotal\",\"U_AddCost\" as \"Additional Cost\",\"U_IDBValue\" as \"IDB Value\" ,\"U_CusDutyPer\" as \"Custom Duty %\",\"U_SubTotal1\" as \"Custom Duty\", " +
                   "\"U_ExDutyPer\" as \"Excise Duty %\",\"U_TotCusDuty\" as \"Base Value For Excise Duty\" , " +
                    "\"U_SubTotal2\" as \"Excise Duty\",\"U_TotExDuty\" as \"Base Value For VAT\",\"U_VatPer\" as \"VAT %\",\"U_VatAmt\" as \"VAT Amount\"  from \"@LANDEDCOSTL1\" where \"U_GRPODocEntry\" = '" + DocEntry + "'";

            _GridDataTable.ExecuteQuery(SQL1);
            _Grid.DataTable = _GridDataTable;
            _RowHeaders = _Grid.RowHeaders;
            _RowHeaders.TitleObject.Caption = "#";

            for (int i = 0; i < _Grid.Rows.Count; i++)
            {
                _RowHeaders.SetText(i, (i + 1).ToString());

            }
            _Grid.Columns.Item("Code").Width = 0;
            _Column = (SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(1);
            _Column.LinkedObjectType = "4";

            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Additional Cost")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("IDB Value")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Base Value For VAT")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Base Value For Excise Duty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Custom Duty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("Excise Duty")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("LineTotal")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("VAT Amount")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
    
            

            _Grid.Columns.Item(0).Editable = false;
            _Grid.Columns.Item(1).Editable = false;
            _Grid.Columns.Item(2).Editable = false;
            _Grid.Columns.Item(3).Editable = false;
            _Grid.Columns.Item(4).Editable = false;
            _Grid.Columns.Item(5).Editable = false;
            
            string _Sql1 = "SELECT * from \"IPF1\" where \"BaseEntry\" = '" + DocEntry + "' ";
            Utilities.ExecuteSQL(ref _RSChk, _Sql1);
            if (_RSChk.RecordCount > 0)
            {
                _Grid.Columns.Item(6).Editable = false;
                _Grid.Columns.Item(7).Editable = false;
                _Grid.Columns.Item(8).Editable = false;
                _Grid.Columns.Item(9).Editable = false;
                _Grid.Columns.Item(10).Editable = false;
                _Grid.Columns.Item(11).Editable = false;
                _Grid.Columns.Item(12).Editable = false;
                _Grid.Columns.Item(13).Editable = false;
                _Grid.Columns.Item(14).Editable = false;
                _Grid.Columns.Item(15).Editable = false;
            }

        }

        #endregion

        #region LandedCostHeadsLoadCostTab

        internal void LandedCostHeadsLoadCostTab(string DocEntry)
        {
            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
            _GridDataTable = _Form.DataSources.DataTables.Item("dtReviewList1");
            StrLandedCostSQL = "select  0 as \"Code\",\"U_SeqNo\" as \"LandedCost Code\" , \"U_LandedCost\" as \"LandedCost Name\",CAST('' as VARCHAR(10)) as \"Currency\",CAST('' as VARCHAR(10)) as \"Exchange Rate\",\"U_Amount\" as \"Amount\",\"U_AmountNPR\" as \"Amount(NPR)\" FROM \"@LANDCOSTSETUP\" Order by  CAST(\"U_SeqNo\" as Numeric)";//,\"Name\" as \"LandedCost Code\"
            _GridDataTable.ExecuteQuery(StrLandedCostSQL);
            _Grid.DataTable = _GridDataTable;
            _RowHeaders = _Grid.RowHeaders;
            _RowHeaders.TitleObject.Caption = "#";

            for (int i = 0; i < _Grid.Rows.Count; i++)
            {
                _RowHeaders.SetText(i, (i + 1).ToString());

            }
            _Grid.Columns.Item("Code").Width = 0;
            _Grid.Columns.Item("Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
            //SAPbouiCOM.ComboBoxColumn _ComboBoxColumn;
            
            _ComboBoxColumn = (SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Currency");
            _ComboBoxColumn.ValidValues.Add("-", "-");
            _ComboBoxColumn.ValidValues.Add("AUD","Australian Dollar");
            _ComboBoxColumn.ValidValues.Add("BND", "Brunei Dollar");
            _ComboBoxColumn.ValidValues.Add("CNY", "Chinese Yuan");
            _ComboBoxColumn.ValidValues.Add("EUR", "Euro");
            _ComboBoxColumn.ValidValues.Add("GBP", "British Pound");
            _ComboBoxColumn.ValidValues.Add("INR", "Indian Rupee");
            _ComboBoxColumn.ValidValues.Add("MYR", "Malaysian Ringgit");
            _ComboBoxColumn.ValidValues.Add("NPR", "Nepalese Rupee");
            _ComboBoxColumn.ValidValues.Add("SGD", "Singapore Dollar");
            _ComboBoxColumn.ValidValues.Add("USD", "US Dollar");

            /*----------------------------*/
          

            _ComboBoxColumn.SetSelectedValue(0, _ComboBoxColumn.ValidValues.Item(6));

            string _SQL = "select \"Rate\" FROM \"ORTT\" where \"RateDate\" = '" + _PostingDate + "' and \"Currency\" = '" + _DocCurency + "'";
            Utilities.ExecuteSQL(ref _Rs, _SQL);
            if (_Rs.RecordCount > 0)
            _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(0).Value = _Rs.Fields.Item(0).Value.ToString();

            _Grid.DataTable.Columns.Item("Amount").Cells.Item(0).Value = _DocTotal.ToString();
            _Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(0).Value = (_DocTotal * _Rs.Fields.Item(0).Value).ToString();
            _Grid.Columns.Item(6).Editable = false;
            for (int j = 1; j <= 30; j++)
            {
                if (j < 9)
                {
                    _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(j).Value = "0.0000";
                }
                else
                {
                    _ComboBoxColumn.SetSelectedValue(j, _ComboBoxColumn.ValidValues.Item(8));
                    _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(j).Value = "1.0000";
                }
            }

            _Grid.Columns.Item(0).Editable = false;
            _Grid.Columns.Item(1).Editable = false;
            _Grid.Columns.Item(2).Editable = false;



            _Grid.CommonSetting.SetCellEditable(10, 4, false);
            _Grid.CommonSetting.SetCellEditable(11, 4, false);
            _Grid.CommonSetting.SetCellEditable(12, 4, false);
            _Grid.CommonSetting.SetCellEditable(13, 4, false);
            _Grid.CommonSetting.SetCellEditable(14, 4, false);
            _Grid.CommonSetting.SetCellEditable(15, 4, false);
            _Grid.CommonSetting.SetCellEditable(16, 4, false);
            _Grid.CommonSetting.SetCellEditable(17, 4, false);
            _Grid.CommonSetting.SetCellEditable(18, 4, false);
            _Grid.CommonSetting.SetCellEditable(19, 4, false);
            _Grid.CommonSetting.SetCellEditable(20, 4, false);
            _Grid.CommonSetting.SetCellEditable(21, 4, false);
            _Grid.CommonSetting.SetCellEditable(22, 4, false);
            _Grid.CommonSetting.SetCellEditable(23, 4, false);
            _Grid.CommonSetting.SetCellEditable(24, 4, false);
            _Grid.CommonSetting.SetCellEditable(25, 4, false);
            _Grid.CommonSetting.SetCellEditable(26, 4, false);
            _Grid.CommonSetting.SetCellEditable(27, 4, false);
            _Grid.CommonSetting.SetCellEditable(28, 4, false);
            _Grid.CommonSetting.SetCellEditable(29, 4, false);
            _Grid.CommonSetting.SetCellEditable(30, 4, false);
            _Grid.CommonSetting.SetCellEditable(31, 4, false);


            _Grid.CommonSetting.SetCellEditable(10, 5, false);
            _Grid.CommonSetting.SetCellEditable(11, 5, false);
            _Grid.CommonSetting.SetCellEditable(12, 5, false);
            _Grid.CommonSetting.SetCellEditable(13, 5, false);
            _Grid.CommonSetting.SetCellEditable(14, 5, false);
            _Grid.CommonSetting.SetCellEditable(15, 5, false);
            _Grid.CommonSetting.SetCellEditable(16, 5, false);
            _Grid.CommonSetting.SetCellEditable(17, 5, false);
            _Grid.CommonSetting.SetCellEditable(18, 5, false);
            _Grid.CommonSetting.SetCellEditable(19, 5, false);
            _Grid.CommonSetting.SetCellEditable(20, 5, false);
            _Grid.CommonSetting.SetCellEditable(21, 5, false);
            _Grid.CommonSetting.SetCellEditable(22, 5, false);
            _Grid.CommonSetting.SetCellEditable(23, 5, false);
            _Grid.CommonSetting.SetCellEditable(24, 5, false);
            _Grid.CommonSetting.SetCellEditable(25, 5, false);
            _Grid.CommonSetting.SetCellEditable(26, 5, false);
            _Grid.CommonSetting.SetCellEditable(27, 5, false);
            _Grid.CommonSetting.SetCellEditable(28, 5, false);
            _Grid.CommonSetting.SetCellEditable(29, 5, false);
            _Grid.CommonSetting.SetCellEditable(30, 5, false);
            _Grid.CommonSetting.SetCellEditable(31, 5, false);
            
        }

        #endregion

        #region LandedCostHeadsLoadFromUDT

        internal void LandedCostHeadsLoadFromUDT(string DocEntry)
        {
            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
            _GridDataTable = _Form.DataSources.DataTables.Item("dtReviewList1");
            StrLandedCostSQL = "select  \"Code\",\"U_LandCstCod\" as \"LandedCost Code\",\"U_LandedCst\" as \"LandedCost Name\",\"U_Currency\" as \"Currency\",\"U_ExRate\" as \"Exchange Rate\",\"U_Amount\" as \"Amount\",\"U_AmountNPR\" as \"Amount(NPR)\" from \"@LANDEDCOSTL2\" where \"U_GRPODocEntry\" = '" + DocEntry + "' Order by  CAST(\"U_LandCstCod\" as Numeric)";
            _GridDataTable.ExecuteQuery(StrLandedCostSQL);
            _Grid.DataTable = _GridDataTable;
            _RowHeaders = _Grid.RowHeaders;
            _RowHeaders.TitleObject.Caption = "#";

            for (int i = 0; i < _Grid.Rows.Count; i++)
            {
                _RowHeaders.SetText(i, (i + 1).ToString());

            }
            _Grid.Columns.Item("Code").Width = 0;
            _Grid.Columns.Item(0).Editable = false;
            _Grid.Columns.Item(1).Editable = false;
            _Grid.Columns.Item(2).Editable = false;
            //_Grid.Columns.Item(3).Editable = false;
            string _Sql1 = "SELECT * from \"IPF1\" where \"BaseEntry\" = '" + DocEntry + "' ";
            Utilities.ExecuteSQL(ref _RSChk, _Sql1);
            if (_RSChk.RecordCount > 0)
            {
               _Grid.Columns.Item(3).Editable = false;
            }

            /*-----------------*/

            _Grid.Columns.Item("Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
            //SAPbouiCOM.ComboBoxColumn _ComboBoxColumn;

            _ComboBoxColumn = (SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Currency");
            _ComboBoxColumn.ValidValues.Add("-", "-");
            _ComboBoxColumn.ValidValues.Add("AUD", "Australian Dollar");
            _ComboBoxColumn.ValidValues.Add("BND", "Brunei Dollar");
            _ComboBoxColumn.ValidValues.Add("CNY", "Chinese Yuan");
            _ComboBoxColumn.ValidValues.Add("EUR", "Euro");
            _ComboBoxColumn.ValidValues.Add("GBP", "British Pound");
            _ComboBoxColumn.ValidValues.Add("INR", "Indian Rupee");
            _ComboBoxColumn.ValidValues.Add("MYR", "Malaysian Ringgit");
            _ComboBoxColumn.ValidValues.Add("NPR", "Nepalese Rupee");
            _ComboBoxColumn.ValidValues.Add("SGD", "Singapore Dollar");
            _ComboBoxColumn.ValidValues.Add("USD", "US Dollar");

            /*----------------------------*/


            _ComboBoxColumn.SetSelectedValue(0, _ComboBoxColumn.ValidValues.Item(6));

            /*------------------*/

            _Grid.CommonSetting.SetCellEditable(10, 4, false);
            _Grid.CommonSetting.SetCellEditable(11, 4, false);
            _Grid.CommonSetting.SetCellEditable(12, 4, false);
            _Grid.CommonSetting.SetCellEditable(13, 4, false);
            _Grid.CommonSetting.SetCellEditable(14, 4, false);
            _Grid.CommonSetting.SetCellEditable(15, 4, false);
            _Grid.CommonSetting.SetCellEditable(16, 4, false);
            _Grid.CommonSetting.SetCellEditable(17, 4, false);
            _Grid.CommonSetting.SetCellEditable(18, 4, false);
            _Grid.CommonSetting.SetCellEditable(19, 4, false);
            _Grid.CommonSetting.SetCellEditable(20, 4, false);
            _Grid.CommonSetting.SetCellEditable(21, 4, false);
            _Grid.CommonSetting.SetCellEditable(22, 4, false);
            _Grid.CommonSetting.SetCellEditable(23, 4, false);
            _Grid.CommonSetting.SetCellEditable(24, 4, false);
            _Grid.CommonSetting.SetCellEditable(25, 4, false);
            _Grid.CommonSetting.SetCellEditable(26, 4, false);
            _Grid.CommonSetting.SetCellEditable(27, 4, false);
            _Grid.CommonSetting.SetCellEditable(28, 4, false);
            _Grid.CommonSetting.SetCellEditable(29, 4, false);
            _Grid.CommonSetting.SetCellEditable(30, 4, false);
            _Grid.CommonSetting.SetCellEditable(31, 4, false);


            _Grid.CommonSetting.SetCellEditable(10, 5, false);
            _Grid.CommonSetting.SetCellEditable(11, 5, false);
            _Grid.CommonSetting.SetCellEditable(12, 5, false);
            _Grid.CommonSetting.SetCellEditable(13, 5, false);
            _Grid.CommonSetting.SetCellEditable(14, 5, false);
            _Grid.CommonSetting.SetCellEditable(15, 5, false);
            _Grid.CommonSetting.SetCellEditable(16, 5, false);
            _Grid.CommonSetting.SetCellEditable(17, 5, false);
            _Grid.CommonSetting.SetCellEditable(18, 5, false);
            _Grid.CommonSetting.SetCellEditable(19, 5, false);
            _Grid.CommonSetting.SetCellEditable(20, 5, false);
            _Grid.CommonSetting.SetCellEditable(21, 5, false);
            _Grid.CommonSetting.SetCellEditable(22, 5, false);
            _Grid.CommonSetting.SetCellEditable(23, 5, false);
            _Grid.CommonSetting.SetCellEditable(24, 5, false);
            _Grid.CommonSetting.SetCellEditable(25, 5, false);
            _Grid.CommonSetting.SetCellEditable(26, 5, false);
            _Grid.CommonSetting.SetCellEditable(27, 5, false);
            _Grid.CommonSetting.SetCellEditable(28, 5, false);
            _Grid.CommonSetting.SetCellEditable(29, 5, false);
            _Grid.CommonSetting.SetCellEditable(30, 5, false);
            _Grid.CommonSetting.SetCellEditable(31, 5, false);

            _Grid.Columns.Item(6).Editable = false;
        }

        #endregion

        #region FillHeaderDataFromUDT

        internal void FillHeaderDataFromUDT(string DocEntry)
        {
            DateTime PostingDate;

            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
            _GridDataTable = _Form.DataSources.DataTables.Item("dtReviewList1");
            StrLandedCostSQL = "SELECT \"U_DocNum\",\"U_Series\",\"U_PostingDate\",\"U_DueDate\",\"U_RefNo\",\"U_FileNo\",\"U_SuppCode\",\"U_SuppName\",\"U_Curency\",\"U_Branch\",\"U_GRPODocNum\",\"U_GRPODocEntry\",\"U_ConType\",\"U_CRAccount\",\"U_Remarks\",\"U_CountryC\",\"U_CountryN\",\"U_CNTRAcont\",\"U_PPNo\" FROM \"@LANDEDCOSTH\" where \"U_GRPODocEntry\" = '" + _DocEntry + "' ";
            Utilities.ExecuteSQL(ref _RS, StrLandedCostSQL);
            ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = _RS.Fields.Item("U_DocNum").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = _RS.Fields.Item("U_Series").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value = Convert.ToDateTime(_RS.Fields.Item("U_PostingDate").Value.ToString()).ToString("yyyyMMdd");
            ((SAPbouiCOM.EditText)_Form.Items.Item("16").Specific).Value = Convert.ToDateTime(_RS.Fields.Item("U_DueDate").Value.ToString()).ToString("yyyyMMdd");
            ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value = _RS.Fields.Item("U_RefNo").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = _RS.Fields.Item("U_FileNo").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("32").Specific).Value = _RS.Fields.Item("U_SuppCode").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value = _RS.Fields.Item("U_SuppName").Value.ToString();
            ((SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific).Select(_RS.Fields.Item("U_Curency").Value.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
            //((SAPbouiCOM.ComboBox)_Form.Items.Item("8").Specific).Select(_RS.Fields.Item("U_Branch").Value.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
            ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value = _RS.Fields.Item("U_GRPODocNum").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("1000003").Specific).Value = _RS.Fields.Item("U_GRPODocEntry").Value.ToString();
           // ((SAPbouiCOM.ComboBox)_Form.Items.Item("39").Specific).Select(_RS.Fields.Item("U_ConType").Value.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
            ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value = _RS.Fields.Item("U_CRAccount").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("42").Specific).Value = _RS.Fields.Item("U_Remarks").Value.ToString();

            ((SAPbouiCOM.EditText)_Form.Items.Item("49").Specific).Value = _RS.Fields.Item("U_CountryC").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("50").Specific).Value = _RS.Fields.Item("U_CountryN").Value.ToString();
           // ((SAPbouiCOM.EditText)_Form.Items.Item("52").Specific).Value = _RS.Fields.Item("U_CNTRAcont").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("54").Specific).Value = _RS.Fields.Item("U_PPNo").Value.ToString();

            string _Sql1 = "SELECT * from \"IPF1\" where \"BaseEntry\" = '" + _DocEntry + "' ";
            Utilities.ExecuteSQL(ref _RSChk, _Sql1);
            if (_RSChk.RecordCount > 0)
            {
                _Form.Items.Item("10").Enabled = false;
                _Form.Items.Item("12").Enabled = false;
                _Form.Items.Item("14").Enabled = false;
                _Form.Items.Item("16").Enabled = false;
                _Form.Items.Item("18").Enabled = false;
                _Form.Items.Item("20").Enabled = false;
                _Form.Items.Item("32").Enabled = false;
                _Form.Items.Item("4").Enabled = false;
                _Form.Items.Item("1000001").Enabled = false;
                _Form.Items.Item("8").Enabled = false;
                _Form.Items.Item("30").Enabled = false;
                _Form.Items.Item("888").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("1000003").Enabled = false;
                _Form.Items.Item("41").Enabled = false;
                _Form.Items.Item("49").Enabled = false;
                _Form.Items.Item("50").Enabled = false;
                _Form.Items.Item("54").Enabled = false;

            }
            else
            {
                _Form.Items.Item("32").Enabled = false;
                _Form.Items.Item("4").Enabled = false;
                _Form.Items.Item("30").Enabled = false;
                _Form.Items.Item("10").Enabled = false;
                _Form.Items.Item("1000003").Enabled = false;
                _Form.Items.Item("32").Enabled = false;
                _Form.Items.Item("4").Enabled = false;
                _Form.Items.Item("1000001").Enabled = false;
                _Form.Items.Item("49").Enabled = false;
                _Form.Items.Item("50").Enabled = false;
                _Form.Items.Item("888").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("20").Enabled = false;
                _Form.Items.Item("8").Enabled = false;
                _Form.Items.Item("12").Enabled = false;
                _Form.Items.Item("18").Enabled = false;
            }


        }

        #endregion

        #region SaveData

        private bool SaveData()
        {
            int intLineID = 0;
            try
            {
                //Form.Freeze(true);
                UDTHeader = Utilities.Application.Company.UserTables.Item("LANDEDCOSTH");
                UDTLine1 = Utilities.Application.Company.UserTables.Item("LANDEDCOSTL1");
                UDTLine2 = Utilities.Application.Company.UserTables.Item("LANDEDCOSTL2");

                int j = 0;
                int k = 0;
                int l = 0;
                #region Save Header Data
                string LandedCostNoH = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString().Trim();
                string ChkGetVal = "select \"Code\" from \"@LANDEDCOSTH\" where \"U_DocNum\" = '"+ LandedCostNoH +"'";
                Utilities.ExecuteSQL(ref _RSChkGetVal, ChkGetVal);
                string test = _RSChkGetVal.Fields.Item(0).Value.ToString().Trim();

                string series = "select Top 1 \"Series\" from NNM1 where \"ObjectCode\" = 69  " + //and \"BPLId\" = '" + ((SAPbouiCOM.ComboBox)_Form.Items.Item("8").Specific).Selected.Value.ToString() + "'
                              " and \"Indicator\" = (select \"Indicator\" from OFPR  where '" + ((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value.ToString() + "' " +
                              " between \"F_RefDate\" and \"T_RefDate\") ";
                Utilities.ExecuteSQL(ref _RSSeries, series);

                if (UDTHeader.GetByKey(_RSChkGetVal.Fields.Item(0).Value.ToString().Trim())) // Code 
                {
                    UDTHeader.UserFields.Fields.Item("U_Series").Value = _RSSeries.Fields.Item("Series").Value.ToString().Trim();//((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_PostingDate").Value = Utilities.ToDate(((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value.ToString());
                    UDTHeader.UserFields.Fields.Item("U_DueDate").Value = Utilities.ToDate(((SAPbouiCOM.EditText)_Form.Items.Item("16").Specific).Value.ToString());
                    UDTHeader.UserFields.Fields.Item("U_RefNo").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_FileNo").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_SuppCode").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("32").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_SuppName").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_Curency").Value = ((SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific).Selected.Value.ToString();
                  //  UDTHeader.UserFields.Fields.Item("U_Branch").Value = ((SAPbouiCOM.ComboBox)_Form.Items.Item("8").Specific).Selected.Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_CRAccount").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_Remarks").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("42").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_CountryC").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("49").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_CountryN").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("50").Specific).Value.ToString();
                   // UDTHeader.UserFields.Fields.Item("U_CNTRAcont").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("52").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_PPNo").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("54").Specific).Value.ToString();

                    k = UDTHeader.Update();
                    if (k != 0)
                    {
                        return false;
                    }
                }
                else
                {
                    UDTHeader.Code = Utilities.getMaxColumnValue("@LANDEDCOSTH", "Code").ToString();
                    UDTHeader.Name = Utilities.getMaxColumnValue("@LANDEDCOSTH", "Code").ToString();
                    UDTHeader.UserFields.Fields.Item("U_DocNum").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_Series").Value = _RSSeries.Fields.Item("Series").Value.ToString().Trim();//((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_PostingDate").Value = Utilities.ToDate(((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value.ToString());
                    UDTHeader.UserFields.Fields.Item("U_DueDate").Value = Utilities.ToDate(((SAPbouiCOM.EditText)_Form.Items.Item("16").Specific).Value.ToString());
                    UDTHeader.UserFields.Fields.Item("U_RefNo").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_FileNo").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_SuppCode").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("32").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_SuppName").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_Curency").Value = ((SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific).Selected.Value.ToString();
                  //  UDTHeader.UserFields.Fields.Item("U_Branch").Value = ((SAPbouiCOM.ComboBox)_Form.Items.Item("8").Specific).Selected.Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_GRPODocNum").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_GRPODocEntry").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("1000003").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_CRAccount").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_Remarks").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("42").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_CountryC").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("49").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_CountryN").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("50").Specific).Value.ToString();
                 //   UDTHeader.UserFields.Fields.Item("U_CNTRAcont").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("52").Specific).Value.ToString();
                    UDTHeader.UserFields.Fields.Item("U_PPNo").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("54").Specific).Value.ToString();

                    k = UDTHeader.Add();
                    if (k != 0)
                    {
                        return false;
                    }
                }

                #endregion

                #region Save Line1 Data
                _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
                for (int i = 0; i < _Grid.Rows.Count; i++)
                {
                    if (UDTLine1.GetByKey(_Grid.DataTable.Columns.Item("Code").Cells.Item(i).Value.ToString().Trim())) // Code 
                    {
                        UDTLine1.UserFields.Fields.Item("U_DocNum").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_ItemCode").Value = _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_ItemName").Value = _Grid.DataTable.Columns.Item("Dscription").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_Quantity").Value = _Grid.DataTable.Columns.Item("Quantity").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_Price").Value = _Grid.DataTable.Columns.Item("Price").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_LineTotal").Value = _Grid.DataTable.Columns.Item("LineTotal").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_SubTotal1").Value = _Grid.DataTable.Columns.Item("Custom Duty").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_ExDutyPer").Value = _Grid.DataTable.Columns.Item("Excise Duty %").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_TotExDuty").Value = _Grid.DataTable.Columns.Item("Base Value For VAT").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_SubTotal2").Value = _Grid.DataTable.Columns.Item("Excise Duty").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_CusDutyPer").Value = _Grid.DataTable.Columns.Item("Custom Duty %").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_TotCusDuty").Value = _Grid.DataTable.Columns.Item("Base Value For Excise Duty").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_AddCost").Value = _Grid.DataTable.Columns.Item("Additional Cost").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_VatPer").Value = _Grid.DataTable.Columns.Item("VAT %").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_VatAmt").Value = _Grid.DataTable.Columns.Item("VAT Amount").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_GRPODocNum").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_GRPODocEntry").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("1000003").Specific).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_IDBValue").Value = _Grid.DataTable.Columns.Item("IDB Value").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_FileNo").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value.ToString();

                        j = UDTLine1.Update();
                        if (j != 0)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        string _Code = Utilities.getMaxColumnValue("@LANDEDCOSTL1", "Code");
                        UDTLine1.Code = _Code.ToString();
                        UDTLine1.Name = _Code.ToString();
                        UDTLine1.UserFields.Fields.Item("U_DocNum").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_ItemCode").Value = _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_ItemName").Value = _Grid.DataTable.Columns.Item("Dscription").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_Quantity").Value = _Grid.DataTable.Columns.Item("Quantity").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_Price").Value = _Grid.DataTable.Columns.Item("Price").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_LineTotal").Value = _Grid.DataTable.Columns.Item("LineTotal").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_SubTotal1").Value = _Grid.DataTable.Columns.Item("Custom Duty").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_ExDutyPer").Value = _Grid.DataTable.Columns.Item("Excise Duty %").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_TotExDuty").Value = _Grid.DataTable.Columns.Item("Base Value For VAT").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_SubTotal2").Value = _Grid.DataTable.Columns.Item("Excise Duty").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_CusDutyPer").Value = _Grid.DataTable.Columns.Item("Custom Duty %").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_TotCusDuty").Value = _Grid.DataTable.Columns.Item("Base Value For Excise Duty").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_AddCost").Value = _Grid.DataTable.Columns.Item("Additional Cost").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_VatPer").Value = _Grid.DataTable.Columns.Item("VAT %").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_VatAmt").Value = _Grid.DataTable.Columns.Item("VAT Amount").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_GRPODocNum").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_GRPODocEntry").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("1000003").Specific).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_IDBValue").Value = _Grid.DataTable.Columns.Item("IDB Value").Cells.Item(i).Value.ToString();
                        UDTLine1.UserFields.Fields.Item("U_FileNo").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value.ToString();
                        j = UDTLine1.Add();
                        if (j != 0)
                        {
                            return false;
                        }
                    }

                }
                #endregion

                #region Save Line1 Data
                _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                for (int i = 0; i < _Grid.Rows.Count; i++)
                {
                    if (UDTLine2.GetByKey(_Grid.DataTable.Columns.Item("Code").Cells.Item(i).Value.ToString().Trim())) // Code 
                    {
                        UDTLine2.UserFields.Fields.Item("U_DocNum").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_LandCstCod").Value = _Grid.DataTable.Columns.Item("LandedCost Code").Cells.Item(i).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_LandedCst").Value = _Grid.DataTable.Columns.Item("LandedCost Name").Cells.Item(i).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_Amount").Value = _Grid.DataTable.Columns.Item("Amount").Cells.Item(i).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_AmountNPR").Value = _Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(i).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_GRPODocNum").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_GRPODocEntry").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("1000003").Specific).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_Currency").Value = _Grid.DataTable.Columns.Item("Currency").Cells.Item(i).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_ExRate").Value = _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(i).Value.ToString();
                        l = UDTLine2.Update();
                        if (l != 0)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        string _Code = Utilities.getMaxColumnValue("@LANDEDCOSTL2", "Code");
                        UDTLine2.Code = _Code.ToString();
                        UDTLine2.Name = _Code.ToString();
                        UDTLine2.UserFields.Fields.Item("U_DocNum").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_LandCstCod").Value = _Grid.DataTable.Columns.Item("LandedCost Code").Cells.Item(i).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_LandedCst").Value = _Grid.DataTable.Columns.Item("LandedCost Name").Cells.Item(i).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_Amount").Value = _Grid.DataTable.Columns.Item("Amount").Cells.Item(i).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_AmountNPR").Value = _Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(i).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_GRPODocNum").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_GRPODocEntry").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("1000003").Specific).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_Currency").Value = _Grid.DataTable.Columns.Item("Currency").Cells.Item(i).Value.ToString();
                        UDTLine2.UserFields.Fields.Item("U_ExRate").Value = _Grid.DataTable.Columns.Item("Exchange Rate").Cells.Item(i).Value.ToString();
                        l = UDTLine2.Add();
                        if (l != 0)
                        {
                            return false;
                        }
                    }

                }
                #endregion
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                return false;
            }
            finally
            {
                //Form.Freeze(false);
            }
            return true;
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
                if (this._CFLEvent.ColUID == "Amount" && PvalCount == 8)
                {
                    _Form.Freeze(true);
                    _Grid.DataTable.Columns.Item("Amount").Cells.Item(8).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                    _Grid.DataTable.Columns.Item("Amount").Cells.Item(9).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][1].ToString();
                    //_Grid.SetCellFocus(8, 2);
                    _Form.Freeze(false);
                }
              
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.ToString());
            }
        }

        #endregion

        #region LANDED COSTS

        private void landedcosts(string CardCode, int Docentry, DateTime DueDate, DateTime PostingDate, string DocumentType)
        {
            SAPbobsCOM.LandedCostsService oLandedCostsService = null;
            SAPbobsCOM.LandedCost oLandedCost = null;
            SAPbobsCOM.LandedCost_ItemLine oLandedCost_ItemLine = null;
            SAPbobsCOM.LandedCost_CostLine oLandedCost_CostLine = null;
            SAPbobsCOM.LandedCostParams oLandedCostParams = null;
            SAPbobsCOM.LandedCostParams oLandedCostParams1 = null;

            string _Sql = "";
            try
            {
                string _Sql11 = "select \"AliasName\"  FROM OCRD where \"CardCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString().Trim() + "' ";
                Utilities.ExecuteSQL(ref _RsBP, _Sql11);

                Remarks = _RsBP.Fields.Item("AliasName").Value.ToString() + ((SAPbouiCOM.EditText)_Form.Items.Item("54").Specific).Value.ToString().Trim();

                if (DocumentType == "GRPO")
                {
                    _Sql = "SELECT \"LineNum\" FROM PDN1 where \"DocEntry\" = " + Docentry + "";
                }
                else
                {
                    _Sql = "SELECT \"LineNum\" FROM PCH1 where \"DocEntry\" = " + Docentry + "";
                }
                Utilities.ExecuteSQL(ref _Rs, _Sql);

                string _Sql1 = "select \"U_SubTotal1\" + \"U_SubTotal2\" + \"U_VatAmt\" FROM \"@LANDEDCOSTL1\" where \"U_GRPODocEntry\" = " + Docentry + " and \"U_FileNo\" = '" + DocumentType + "' ";
                Utilities.ExecuteSQL(ref _RsCal, _Sql1);

                double test = 0;
                if (_Rs != null)
                {
                    string series = "select Top 1 \"Series\" from NNM1 where \"ObjectCode\" = 69  " +//and \"BPLId\" = '" + ((SAPbouiCOM.ComboBox)_Form.Items.Item("8").Specific).Selected.Value.ToString() + "'
                             " and \"Indicator\" = (select \"Indicator\" from OFPR  where '" + ((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value.ToString() + "' " +
                             " between \"F_RefDate\" and \"T_RefDate\") ";

                    Utilities.ExecuteSQL(ref _RSSeries, series);

                    oLandedCostsService = (SAPbobsCOM.LandedCostsService)Utilities.Application.Company.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.LandedCostsService);
                    oLandedCost = (SAPbobsCOM.LandedCost)oLandedCostsService.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCost);

                    long oLandedCostEntry = 0;
                    //header part

                    oLandedCost.VendorCode = CardCode;
                    oLandedCost.DocumentCurrency = "NPR";
                    oLandedCost.DueDate = Utilities.ToDate(((SAPbouiCOM.EditText)_Form.Items.Item("16").Specific).Value.ToString().Trim());
                    oLandedCost.PostingDate = Utilities.ToDate(((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value.ToString().Trim());
                    oLandedCost.Series = Convert.ToInt16(_RSSeries.Fields.Item("Series").Value.ToString().Trim());//Convert.ToInt16(((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString().Trim());
                    oLandedCost.Remarks = Remarks.ToString();
                    //oLandedCost.FileNumber = Remarks.ToString();
                    oLandedCost.Broker = ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString().Trim();

                    while (!_Rs.EoF)
                    {
                        // item line
                        test = Convert.ToDouble(_RsCal.Fields.Item(0).Value.ToString());

                        oLandedCost_ItemLine = (SAPbobsCOM.LandedCost_ItemLine)oLandedCost.LandedCost_ItemLines.Add();
                        if (DocumentType == "GRPO")
                        {
                            oLandedCost_ItemLine.BaseDocumentType = SAPbobsCOM.LandedCostBaseDocumentTypeEnum.asGoodsReceiptPO;
                        }
                        else
                        {
                            oLandedCost_ItemLine.BaseDocumentType = SAPbobsCOM.LandedCostBaseDocumentTypeEnum.asPurchaseInvoice;
                        }
                        oLandedCost_ItemLine.BaseEntry = Docentry;
                        oLandedCost_ItemLine.BaseLine = Convert.ToInt32(_Rs.Fields.Item(0).Value.ToString());
                        oLandedCost_ItemLine.CustomsValue = Convert.ToDouble(_RsCal.Fields.Item(0).Value.ToString());
                        _Rs.MoveNext();
                        _RsCal.MoveNext();
                    }

                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(1).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "01";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(1).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(2).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "02";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(2).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(3).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "03";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(3).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(4).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "04";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(4).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(5).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "05";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(5).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(6).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "06";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(6).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(7).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "07";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(7).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(8).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "08";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(8).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(15).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "09";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(15).Value)) , 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(17).Value != "0.00")
                    {
                        string aa = _Grid.DataTable.Columns.Item("Amount").Cells.Item(17).Value;
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "10";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(17).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(18).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "11";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(18).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(19).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "12";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(19).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(20).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "13";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(20).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(21).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "14";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(21).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(22).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "15";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(22).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(23).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "16";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(23).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(24).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "17";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(24).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(25).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "18";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(25).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(26).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "19";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(26).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(27).Value != "0.00")
                    {
                        oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                        oLandedCost_CostLine.LandedCostCode = "20";
                        oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount(NPR)").Cells.Item(27).Value)), 2);
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueAfterCustoms;
                    }
                    //if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(28).Value != "0.00")
                    //{
                    //    oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                    //    oLandedCost_CostLine.LandedCostCode = "21";
                    //    oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(28).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(28).Value)), 2);
                    //    oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    //}
                    //if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(29).Value != "0.00")
                    //{
                    //    oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                    //    oLandedCost_CostLine.LandedCostCode = "22";
                    //    oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(29).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(29).Value)), 2);
                    //    oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    //}
                    //if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(30).Value != "0.00")
                    //{
                    //    oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                    //    oLandedCost_CostLine.LandedCostCode = "23";
                    //    oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(30).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(30).Value)), 2);
                    //    oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    //}
                    //if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(31).Value != "0.00")
                    //{
                    //    oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                    //    oLandedCost_CostLine.LandedCostCode = "24";
                    //    oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(31).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(31).Value)), 2);
                    //    oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    //}
                    //if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(32).Value != "0.00")
                    //{
                    //    oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                    //    oLandedCost_CostLine.LandedCostCode = "25";
                    //    oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(32).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(32).Value)), 2);
                    //    oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    //}
                    //if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(33).Value != "0.00")
                    //{
                    //    oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                    //    oLandedCost_CostLine.LandedCostCode = "26";
                    //    oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(33).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(33).Value)), 2);
                    //    oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    //}
                    //if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(34).Value != "0.00")
                    //{
                    //    oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                    //    oLandedCost_CostLine.LandedCostCode = "27";
                    //    oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(34).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(34).Value)), 2);
                    //    oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    //}
                    //if (_Grid.DataTable.Columns.Item("Amount").Cells.Item(35).Value != "0.00")
                    //{
                    //    oLandedCost_CostLine = (SAPbobsCOM.LandedCost_CostLine)oLandedCost.LandedCost_CostLines.Add();
                    //    oLandedCost_CostLine.LandedCostCode = "28";
                    //    oLandedCost_CostLine.amount = Math.Round((Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(35).Value) * Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(35).Value)), 2);
                    //    oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asCashValueBeforeCustoms;
                    //}

                    oLandedCostParams = (SAPbobsCOM.LandedCostParams)oLandedCostsService.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCostParams);
                    oLandedCostParams = oLandedCostsService.AddLandedCost(oLandedCost);
                    oLandedCostEntry = oLandedCostParams.LandedCostNumber;
                    LandedCostNo = oLandedCostEntry.ToString();
                }
            }
            catch
            {
                string _errMsg = string.Empty;
                int _errCode = 0;
                Utilities.Application.Company.GetLastError(out _errCode, out _errMsg);
            }
            finally
            {
                _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
                string _Sql1 = "update OIPF set \"U_PPNo\" = '" + Remarks.ToString() +"' , \"U_TotalCustomDuty\" = '" + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(11).Value) + "',\"U_TotalExciseDuty\" = '" + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value) + "',\"U_TotalVATAmount\" = '" + Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value) + "' where \"DocEntry\" = '" + LandedCostNo + "'";
                Utilities.ExecuteSQL(ref _RsCal, _Sql1);
                
                if (oLandedCostsService != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oLandedCostsService);
                    oLandedCostsService = null;
                }
                if (oLandedCost != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oLandedCost);
                    oLandedCost = null;

                }
                if (oLandedCost_ItemLine != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oLandedCost_ItemLine);
                    oLandedCost_ItemLine = null;

                }

                if (oLandedCost_CostLine != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oLandedCost_CostLine);
                    oLandedCost_CostLine = null;

                }
                if (oLandedCostParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oLandedCostParams);
                    oLandedCostParams = null;

                }
                GC.Collect();
               
            }
        }

        #endregion

        #region FILTER BP Type
        internal void FilteGLAccount(string CFLId)
        {

            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();


                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "CardType";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "S";
                _Con.BracketCloseNum = 1;
                _Form.ChooseFromLists.Item("CFL_1").SetConditions(_Cons);

            }


            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region FILTER BP Type
        internal void FilteGLAccount1(string CFLId)
        {

            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();

                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "Postable";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "Y";
                _Con.BracketCloseNum = 1;

                _Con.Relationship = BoConditionRelationship.cr_AND;
                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "LocManTran";
                _Con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                _Con.CondVal = "Y";
                _Con.BracketCloseNum = 1;

                _Form.ChooseFromLists.Item("CFL_2").SetConditions(_Cons);

            }


            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
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
            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("405").Specific;
            string strQuery = "SELECT \"U_VATINAC\",\"U_CustDtyAC\",\"U_ExDtyAC\",\"U_AdyAllAC\" FROM \"@CSTMDSTPH\" ";
            Utilities.ExecuteSQL(ref _RS, strQuery);

            string strQuery12 = "select \"LaCAllcAcc\" from OALC where \"AlcCode\" = '09'";
            Utilities.ExecuteSQL(ref _RSLandedC, strQuery12);

            string _Sql11 = "select \"AliasName\"  FROM OCRD where \"CardCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString().Trim() + "' ";
            Utilities.ExecuteSQL(ref _RsBP, _Sql11);
            string Remarks = _RsBP.Fields.Item("AliasName").Value.ToString() + ((SAPbouiCOM.EditText)_Form.Items.Item("54").Specific).Value.ToString().Trim();

            DateTime aa = Utilities.ToDate(((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value.ToString().Trim());
            oJE.TaxDate = Utilities.ToDate(((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value.ToString().Trim());//Utilities.ToDate(DateTime.Now.ToString("yyyyMMdd"));
            oJE.DueDate = Utilities.ToDate(((SAPbouiCOM.EditText)_Form.Items.Item("16").Specific).Value.ToString().Trim());//Utilities.ToDate(DateTime.Now.ToString("yyyyMMdd"));
            oJE.ReferenceDate = Utilities.ToDate(((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value.ToString().Trim());//Utilities.ToDate(DateTime.Now.ToString("yyyyMMdd"));
            oJE.Memo = Remarks.ToString();
            oJE.Reference = LandedCostNo.ToString();
            oJE.Reference2 = Remarks.ToString();

            //Credit Custom Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(0);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_CustDtyAC").Value.ToString().Trim();
            oJE.Lines.Credit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(11).Value);
            oJE.Lines.Debit = 0;
           // oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit Custom Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(1);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_AdyAllAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(11).Value);
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());


            //Credit Excise Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(2);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_ExDtyAC").Value.ToString().Trim();
            oJE.Lines.Credit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value);
            oJE.Lines.Debit = 0;
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit Excise Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(3);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_AdyAllAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value);
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Credit VAT Amount
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(4);
            oJE.Lines.ShortName = ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString().Trim();
            oJE.Lines.Credit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value);
            oJE.Lines.Debit = 0;
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit VAT Amount
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(5);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_AdyAllAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(13).Value);
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Credit
            oJE.Lines.SetCurrentLine(6);
            oJE.Lines.ShortName = ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString().Trim();
           // oJE.Lines.AccountCode = ((SAPbouiCOM.EditText)_Form.Items.Item("52").Specific).Value.ToString().Trim();//New
            oJE.Lines.Credit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value);
            oJE.Lines.Debit = 0;
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(7);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_VATINAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(14).Value);
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //-------------------//
            //Credit Custom Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(8);
            //oJE.Lines.AccountCode = ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString().Trim();
            oJE.Lines.ShortName = ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString().Trim();
          //  oJE.Lines.AccountCode = ((SAPbouiCOM.EditText)_Form.Items.Item("52").Specific).Value.ToString().Trim();//New
            oJE.Lines.Credit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(11).Value);
            oJE.Lines.Debit = 0;
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit Custom Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(9);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_CustDtyAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(11).Value);
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Credit Excise Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(10);
            oJE.Lines.ShortName = ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString().Trim();
          //  oJE.Lines.AccountCode = ((SAPbouiCOM.EditText)_Form.Items.Item("52").Specific).Value.ToString().Trim();//New
            oJE.Lines.Credit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value);
            oJE.Lines.Debit = 0;
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit Excise Duty
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(11);
            oJE.Lines.AccountCode = _RS.Fields.Item("U_ExDtyAC").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(12).Value);
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            ///////////////////////////////////////

            //Credit
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(12);
            oJE.Lines.ShortName = ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value.ToString().Trim();
          //  oJE.Lines.AccountCode = ((SAPbouiCOM.EditText)_Form.Items.Item("52").Specific).Value.ToString().Trim();//New
            oJE.Lines.Credit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(15).Value);
            oJE.Lines.Debit = 0;
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

            //Debit
            oJE.Lines.Add();
            oJE.Lines.SetCurrentLine(13);
            oJE.Lines.AccountCode = _RSLandedC.Fields.Item("LaCAllcAcc").Value.ToString().Trim();
            oJE.Lines.Credit = 0;
            oJE.Lines.Debit = Convert.ToDouble(_Grid.DataTable.Columns.Item("Amount").Cells.Item(15).Value);
            //oJE.Lines.BPLID = Convert.ToInt32(_Branch.ToString());

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
                string strQuery1 = "update \"@LANDEDCOSTH\" set \"U_RefNo\" = '" + JENO + "' where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString().Trim() + "'";
                Utilities.ExecuteSQL(ref _RS, strQuery1);

                string _Sql1 = "update OIPF set \"U_JENO\" = '" + JENO + "' where \"DocEntry\" = '" + LandedCostNo + "'";
                Utilities.ExecuteSQL(ref _RsCal, _Sql1);
            }
            
        }

        #endregion

        #region ShowReport

        public void ShowReport()
        {
            try
            {
                /*Web Config Data Read Start*/
                DataSet ds = LoadCompanySetting();

                string Server = ds.Tables[0].Rows[0]["ServerName"].ToString().Trim();
                string UserName = ds.Tables[0].Rows[0]["UserID"].ToString().Trim();
                string Password = ds.Tables[0].Rows[0]["Password"].ToString().Trim();

                /*Web Config Data Read End*/
                _DocEntry = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString().Trim();
                objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
                CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();
                ParameterFields paramFields = new ParameterFields();
                ParameterField paramField = new ParameterField();
                ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
                ParameterRangeValue rangeVal = new ParameterRangeValue();

                ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "CostSheetSummary.rpt";
                objCryRep.Load(ReportPath);

                objConInfo.ServerName = Server.ToString();//ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;

                objConInfo.DatabaseName = Utilities.Application.Company.CompanyDB;

                objConInfo.UserID = UserName.ToString();//ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
                objConInfo.Password = Password.ToString();//ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;
                objConInfo.Type = ConnectionInfoType.CRQE;


                ReportDisplay Report = new ReportDisplay();
                Report.Refresh();

                foreach (CrystalDecisions.CrystalReports.Engine.Table objTable in objCryRep.Database.Tables)
                {
                    objTabLoginInfo = objTable.LogOnInfo;
                    objTabLoginInfo.ConnectionInfo = objConInfo;

                    objTable.ApplyLogOnInfo(objTabLoginInfo);
                    objTabLoginInfo.ConnectionInfo.AllowCustomConnection = true;

                }


                objCryRep.Refresh();
                //  objCryRep.VerifyDatabase();
                //objCryRep.SetParameterValue("ExtParam@", _ExtParam);
                objCryRep.SetParameterValue("DocKey@", _DocEntry);
                //objCryRep.SetParameterValue("SuppSerial@", strMnfSerial);

                Report.crystalReportViewer1.ShowGroupTreeButton = false;
                Report.crystalReportViewer1.ReportSource = objCryRep;
                Report.crystalReportViewer1.Refresh();
                Report.crystalReportViewer1.Show();

                Report.ShowDialog();
                // Report.crystalReportViewer1.
                Report.crystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None;


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion
        
    }
}
