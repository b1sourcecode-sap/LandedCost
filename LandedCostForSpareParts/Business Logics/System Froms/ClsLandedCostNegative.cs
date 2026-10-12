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
    class ClsLandedCostNegative : Base
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
        public static string _DocNum = string.Empty;
        public static string _DocEntry = string.Empty, _Branch = string.Empty;
        public static double TotalVAT = 0;
        public static string ControlAC = string.Empty, CustomOfficeTI = string.Empty, _DocDate, _DueDate, _PPNo;
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
        public ClsLandedCostNegative()
            : base()
        {
            //_FormNegative = Utilities.Application.SBO_Application.Forms.GetForm("-992", _Form.TypeCount);
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsLandedCostNegative()
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
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "U_TotalCustomDuty" && pVal.CharPressed != 9 && _Form.Mode != BoFormMode.fm_ADD_MODE)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "U_TotalExciseDuty" && pVal.CharPressed != 9 && _Form.Mode != BoFormMode.fm_ADD_MODE)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "U_TotalVATAmount" && pVal.CharPressed != 9 && _Form.Mode != BoFormMode.fm_ADD_MODE)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "U_ControlAccount" && pVal.CharPressed != 9 && _Form.Mode != BoFormMode.fm_ADD_MODE)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "U_JENO" && pVal.CharPressed != 9 && _Form.Mode != BoFormMode.fm_ADD_MODE)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "U_PPNo" && pVal.CharPressed != 9 && _Form.Mode != BoFormMode.fm_ADD_MODE)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        break;

                }
            }
            #endregion

            #region After ACTION ACTION
            else if (pVal.Before_Action == false)
            {

                switch (pVal.EventType)
                {

                }
            }

            #endregion

        }
        #endregion
    }
}
