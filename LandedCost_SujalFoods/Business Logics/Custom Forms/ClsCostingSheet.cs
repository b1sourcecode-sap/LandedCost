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

namespace LandedCost_SujalFoods
{
    class ClsCostingSheet : Base
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
        public static string _DocEntry = "", _DocCurency = "", _DocNum = "", _BPCode = "", _BPName = "", _Branch = "", _PostingDate = "", _DueDate = "", _ConSignmentType = "";
        public static double _DocTotal = 0;
        public static string LandedCostNo = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        SAPbouiCOM.ComboBox _ComboBox;
        private SAPbouiCOM.DBDataSource _DBDataSource;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RSLandedCost = null, _Rs = null, _RSChk = null, _RsCal = null, _RSChkGetVal = null, _RSChk1 = null, _RSPriceList = null;
        string SQL1 = string.Empty, SQL = string.Empty, StrLandedCostSQL = string.Empty;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource1, _ChildDBDataSource2, _ChildDBDataSource3;
        SAPbobsCOM.UserTable UDTHeader, UDTLine1, UDTLine2;
        public static int PvalCount = 0;
        public static double SubTotal1 = 0;
        public static double SubTotal2 = 0;
        public static double ExciseDuty = 0;
        public static double CustomDuty = 0;
        public static int count = 1;
        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsCostingSheet()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsCostingSheet()
        {
        }
        #endregion

        #region GRPODataLoadinItemsTab

        internal void FormDefault(string _LandedCstNum)
        {
            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("404").Specific;
            _Form.DataSources.DataTables.Add("dtReviewList");
            _GridDataTable = _Form.DataSources.DataTables.Item("dtReviewList");

            SQL1 = "select T4.\"DocNum\" as \"LandedCostNum\", T1.\"U_ItemCode\" as \"ItemCode\", "+
                   "T1.\"U_ItemName\" as \"Item Name\",T1.\"U_Quantity\" as \"Quantity\",\"U_Price\" as \"Price\",T1.\"U_LineTotal\" as \"LineTotal\", "+
                   "T1.\"U_AddCost\" as \"Additional Cost\", T1.\"U_IDBValue\" as \"IDB Value\" ,T1.\"U_SubTotal1\" as \"CustomeDuty\", "+
                   "T1.\"U_SubTotal2\" as \"Excise Duty\", T1.\"U_VatAmt\" as \"VAT Amount\", "+
                   "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 16 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"DECLARATION FEES\",  " +
                   "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 17 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"Total Duty Paid At Custom\",  " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 18 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"L/C OPENING COMMISSION\", " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 19 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"L/C OPENING CHARGES\",  " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 20 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"L/C CONFIRMATION CHARGES\",  " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 21 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"OTHER BANK CHARGES\",  " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 22 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"SWIFT CHARGES\",  " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 23 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"AMENDMENT CHARGES\",  " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 24 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"CHA CHARGES - NEPAL CUSTOM\",  " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 25 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"CHA CHARGES - NEPAL AGENT\", " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 26 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"WAREHOUSE CHARGES\",  " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 27 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"FREIGHT-NEPAL CUSTOM - FACTORY\",  " +
                    "(select Round(((REPLACE(T2.\"U_AmountNPR\",',', '') / T3.\"DocTotal\") * T1.\"U_LineTotal\"),2) from \"@LANDEDCOSTL2\" T2  where T2.\"U_LandCstCod\" = 28 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\") as \"OTHER CHARGES - NEPAL\" " +
                   
                    "from \"@LANDEDCOSTH\" T0  Inner join \"@LANDEDCOSTL1\" T1 ON T0.\"U_GRPODocNum\" = T1.\"U_GRPODocNum\"  "+
                    "inner join \"OPDN\" T3 ON T3.\"DocEntry\" = T0.\"U_GRPODocEntry\"  inner join \"OIPF\" T4 ON T4.\"DocEntry\" = T3.\"LndCstNum\"  "+
                    "where T1.\"U_DocNum\" = '" + _LandedCstNum + "'";

            _GridDataTable.ExecuteQuery(SQL1);
            _Grid.DataTable = _GridDataTable;
            _RowHeaders = _Grid.RowHeaders;
            _RowHeaders.TitleObject.Caption = "#";

            for (int i = 0; i < _Grid.Rows.Count; i++)
            {
                _RowHeaders.SetText(i, (i + 1).ToString());

            }
            //_Grid.Columns.Item("Code").Width = 0;
            //_Column = (SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(1);
            //_Column.LinkedObjectType = "4";

            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(5)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(6)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(7)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(8)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(9)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(10)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(11)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(12)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(13)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(14)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(15)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(16)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(17)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(18)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(19)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(20)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(21)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(22)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(23)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;

            _Grid.Columns.Item(0).Editable = false;
            _Grid.Columns.Item(1).Editable = false;
            _Grid.Columns.Item(2).Editable = false;
            _Grid.Columns.Item(3).Editable = false;
            _Grid.Columns.Item(4).Editable = false;
            _Grid.Columns.Item(5).Editable = false;
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
            _Grid.Columns.Item(16).Editable = false;
            _Grid.Columns.Item(17).Editable = false;
            _Grid.Columns.Item(18).Editable = false;
            _Grid.Columns.Item(19).Editable = false;
            _Grid.Columns.Item(20).Editable = false;
            _Grid.Columns.Item(21).Editable = false;
            _Grid.Columns.Item(22).Editable = false;
            _Grid.Columns.Item(23).Editable = false;
            //_Grid.Columns.Item(24).Editable = false;
            //_Grid.Columns.Item(25).Editable = false;
            //_Grid.Columns.Item(26).Editable = false;
          
        }

        #endregion
    }
}
