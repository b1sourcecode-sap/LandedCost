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

namespace LandedCost_Automobiles
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

          string SQL2 = "select *,Round(\"Total Cost\"/\"Quantity\",2) as \"Per Unit Cost\", " +
                  " Round(((Round(\"Total Cost\"/\"Quantity\",2) * \"Sales Margin\")/100 + Round(\"Total Cost\"/\"Quantity\",2)),2) AS \"SalesPrice with Margin of %\" from( " +
                  " select *,Round(\"LineTotal\"+\"Additional Cost\"+\"Custom Duty\"+\"Excise Duty\"+\"Dummarge Charges\"+\"PP Charges\" " +
                  " +\"Road Development Tax\"+\"Registration Expenses\"+\"Intrest Cost\"+\"Handling Cost\"+\"C&F Dealer/ Dist.\"+\"C&F Dealer/Distr. Misc.\" " +
                  " +\"Insurance Transit Nepal\"+\"LC & Bank Comission Charges\"+\"Storage 1 Charges\"+\"Storage 2 Charges\",2) AS \"Total Cost\" FROM " +
                  " (select T0.\"U_RefNo\" as \"JE No\",T4.\"DocNum\" as \"LandedCostNum\",T0.\"U_SalesMrgn\" as \"Sales Margin\", " +
                  " T1.\"U_ItemCode\" as \"ItemCode\",T1.\"U_ItemName\" as \"Item Name\",T1.\"U_Quantity\" as \"Quantity\",\"U_Price\" as \"Price\",T1.\"U_LineTotal\" as \"LineTotal\",T1.\"U_AddCost\" as \"Additional Cost\", " +
                  " T1.\"U_SubTotal1\" as \"Custom Duty\",T1.\"U_SubTotal2\" as \"Excise Duty\", " +
                  " (select Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 13 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"Dummarge Charges\", " +
                  " (select Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 15 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"PP Charges\", " +
                  " (select  Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 18 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"Road Development Tax\", " +
                  " (select  Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 19 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"Registration Expenses\", " +
                  " (select  Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 20 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"Intrest Cost\", " +
                  " (select  Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 21 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"Handling Cost\", " +
                  " (select  Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 22 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"C&F Dealer/ Dist.\", " +
                  " (select  Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 23 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"C&F Dealer/Distr. Misc.\", " +
                  " (select  Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 24 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"Insurance Transit Nepal\", " +
                  " (select Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 25 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"LC & Bank Comission Charges\", " +
                  " (select Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 26 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"Storage 1 Charges\", " +
                  " (select Round((REPLACE(T2.\"U_Amount\",',', '') * T1.\"U_LineTotal\")/ T3.\"DocTotal\",2) from \"@LANDEDCOSTL2\" T2 where T2.\"U_LandCstCod\" = 27 and T1.\"U_GRPODocNum\" = T2.\"U_GRPODocNum\" AND T1.\"U_GRPODocEntry\" = T2.\"U_GRPODocEntry\") as \"Storage 2 Charges\" " +
                  " from \"@LANDEDCOSTH\" T0 " +
                  " Inner join \"@LANDEDCOSTL1\" T1 ON T0.\"U_GRPODocEntry\" = T1.\"U_GRPODocEntry\" " +
                  " inner join \"OPDN\" T3 ON T3.\"DocEntry\" = T0.\"U_GRPODocEntry\" " +
                  " inner join \"OIPF\" T4 ON T4.\"DocEntry\" = T3.\"LndCstNum\" " +
                  " where T1.\"U_DocNum\" = '" + _LandedCstNum + "')A)B";

            SQL1 = "select * from(Select *,(\"Per Unit Freight Import\" + \"Per Unit Insurance Import\" + \"Per Unit C&F KOLKATTA TO INDIA BORDER\" + \"Per Unit Insurance Transit India\" " +
                   " +\"Per Unit Custom Duty\" +\"Per Unit Excise Duty\"+ \"Per Unit Registration Expenses\"+ \"Per Unit Dummarge Charges\"+ \"Per Unit PP Charges\"+ " +
                   " \"Per Unit Road Dev Tax\"+\"Per Unit Interest Cost Transit Period\"+\"Per Unit Handling Cost\"+ \"Per Unit C&F CUSTOM - DEALER / DISTR\"+ \"Per Unit C&F CUSTOM - DEALER /DIST MISC\"+ \"Per Unit Insurance Transit Nepal\" + \"Per Unit LC and Bank Commission Charges\"+ \"Per Unit Storage Charges 1\" " +
                   " +\"Per Unit Vehicle Upgradation Charges\"+ \"Per Unit Miscellaneous Expenses\") as \"Per Unit Additional Cost\" , " +
                   " (\"Per Unit Freight Import\" + \"Per Unit Insurance Import\" + \"Per Unit C&F KOLKATTA TO INDIA BORDER\" + \"Per Unit Insurance Transit India\" " +
                   " +\"Per Unit Custom Duty\" +\"Per Unit Excise Duty\"+ \"Per Unit Registration Expenses\"+ \"Per Unit Dummarge Charges\"+ \"Per Unit PP Charges\"+ " +
                   " \"Per Unit Road Dev Tax\"+\"Per Unit Interest Cost Transit Period\"+\"Per Unit Handling Cost\"+ \"Per Unit C&F CUSTOM - DEALER / DISTR\"+ \"Per Unit C&F CUSTOM - DEALER /DIST MISC\"+ \"Per Unit Insurance Transit Nepal\" + \"Per Unit LC and Bank Commission Charges\"+ \"Per Unit Storage Charges 1\" " +
                   " +\"Per Unit Vehicle Upgradation Charges\"+ \"Per Unit Miscellaneous Expenses\" + \"PerUnit LineTotal\") as \"Per Unit Landing Cost\",\"CostTotal\" as \"SAP Loaded Cost\", " +
                   " (\"Per Unit Freight Import\" + \"Per Unit Insurance Import\" + \"Per Unit C&F KOLKATTA TO INDIA BORDER\" + \"Per Unit Insurance Transit India\" " +
                   " +\"Per Unit Custom Duty\" +\"Per Unit Excise Duty\"+ \"Per Unit Registration Expenses\"+ \"Per Unit Dummarge Charges\"+ \"Per Unit PP Charges\"+ " +
                   " \"Per Unit Road Dev Tax\"+\"Per Unit Interest Cost Transit Period\"+\"Per Unit Handling Cost\"+ \"Per Unit C&F CUSTOM - DEALER / DISTR\"+ \"Per Unit C&F CUSTOM - DEALER /DIST MISC\"+ \"Per Unit Insurance Transit Nepal\" + \"Per Unit LC and Bank Commission Charges\"+ \"Per Unit Storage Charges 1\" " +
                   " +\"Per Unit Vehicle Upgradation Charges\"+ \"Per Unit Miscellaneous Expenses\" + \"PerUnit LineTotal\") - \"CostTotal\" as \"Per Unit Difference\", " +
                   " ((\"CostTotal\" * \"Sales Margin Per\") / 100) as \"Sales Margin Value\" ,\"CostTotal\" + ((\"CostTotal\" * \"Sales Margin Per\") / 100) as \"Expected Sales Price\"  " +
                   "  from(select distinct " +
                   " \"DocNum\" as \"GRPO No\",\"DocEntry\" as \"GRPO DocEntry\",\"JE No\",\"LandedCostNum\",\"Sales Margin Per\",\"DocDate\" as \"GRPO Date\",\"Supplier Code\", " +
                   " \"Supplier Name\",\"ItemCode\" as \"Item No\",\"Item Name\", " +
                   " \"Country of Origin\",\"DistNumber\" as \"Engine No\",\"Chasis No\",\"CostTotal\",\"LC No\",\"LC Start Date\",\"LC End Date\", " +
                   " \"Issuing Bank\",\"Landed Cost Num\", \"Costing Date\",\"PP No\",\"Invoice Value (FC)\",\"Invoice Value\" as \"Invoice Value (NP)\"," +
                   " \"DocCur\" as \"Document Currency\",\"DocRate\" as \"Exchange Rate GRPO\", " +
                   " \"Invoice Value\" as \"Total Invoice Value\", \"LineTotal\", \"PerUnit LineTotal\", " +
                   " Round(Replace( (\"Freight Import\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Freight Import\", " +
                   " Round(Replace( (\"Insurance Import\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Insurance Import\", " +
                   " Round(Replace( ( \"C&F KOLKATTA TO INDIA BORDER\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit C&F KOLKATTA TO INDIA BORDER\", " +
                   " Round(Replace( (\"Insurance Transit India\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Insurance Transit India\"," +
                   " Round(Replace( (\"Custom Duty\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Custom Duty\"," +
                   " Round(Replace( (\"Excise Duty\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Excise Duty\"," +
                   " Round(Replace( (\"Dummarge Charges\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Dummarge Charges\"," +
                   " Round(Replace( (\"PP Charges\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit PP Charges\"," +
                   " Round(Replace( (\"Total Road Dev Tax\"/case when \"U_Quantity\" > 0 then \"U_Quantity\" else 1 END),',',''),2) As \"Per Unit Road Dev Tax\"," +
                   " Round(Replace( (\"Registration Expenses\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Registration Expenses\"," +
                   " Round(Replace( (\"Interest Cost Transit Period\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Interest Cost Transit Period\"," +
                   " Round(Replace( (\"Handling Cost\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Handling Cost\"," +
                   " Round(Replace( ( \"C&F CUSTOM - DEALER / DISTR\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit C&F CUSTOM - DEALER / DISTR\"," +
                   " Round(Replace( ( \"C&F CUSTOM - DEALER /DIST MISC\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit C&F CUSTOM - DEALER /DIST MISC\"," +
                   " Round(Replace( ( \"Insurance Transit Nepal\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Insurance Transit Nepal\"," +
                   " Round(Replace( (\"LC and Bank Commission Charges\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit LC and Bank Commission Charges\"," +
                   " Round(Replace( ( \"Storage Charges 1\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Storage Charges 1\"," +
                   " Round(Replace( ( \"Vehicle Upgradation Charges\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Vehicle Upgradation Charges\"," +
                   " Round(Replace( (\"Miscellaneous Expenses\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Miscellaneous Expenses\"," +
                   " Round(Replace( (\"Total VAT Amount\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit VAT Amount\",\"Total VAT Amount\" " +
                   " from(SELECT distinct T0.\"DocDate\",T0.\"DocCur\",T0.\"DocRate\",T0.\"DocNum\",T0.\"DocEntry\", " +
                   " (T1.\"LineTotal\"+T1.\"DistribSum\") as \"LineTotal\", " +
                   " ((T1.\"LineTotal\"+T1.\"DistribSum\")/T1.\"Quantity\") as \"PerUnit LineTotal\",T0.\"NumAtCard\" as \"Vendor Ref No\", " +
                   " T0.\"U_CountryN\" as \"Country of Origin\", " +
                   " T0.\"DocTotal\",T0.\"CardCode\" as \"Supplier Code\",T0.\"CardName\" as \"Supplier Name\",T9.\"U_DocNum\" as \"Landed Cost Num\", " +
                   " T5.\"ItemCode\",T5.\"itemName\" as \"Item Name\",T11.\"ItmsGrpNam\" as \"Item Group\", T5.\"DistNumber\", T5.\"MnfSerial\" as \"Chasis No\",T5.\"CostTotal\", " +
                   " T7.\"StartDate\" as \"LC Start Date\", T7.\"EndDate\" as \"LC End Date\", " +
                   " T7.\"Descript\" as \"LC No\",T7.\"U_BankName\" as \"Issuing Bank\",T8.\"U_DocNum\" as \"Landed Cost No\",T8.\"U_PostingDate\" as \"Costing Date\", " +
                   " T8.\"U_PPNo\" as \"PP No\",ifnull(T9.\"U_TotExDuty\",0) as \"Total Excise Duty\",ifnull(T9.\"U_TotCusDuty\",0) as \"Total Custom Duty\", " +
                   " ifnull(T9.\"U_Quantity\",0) as \"U_Quantity\", " +
                   " ifnull(T9.\"U_TotRDTax\",0) as \"Total Road Dev Tax\"," +
                   " (Select (ROUND(REPLACE(\"U_AmountNPR\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='2' and \"U_DocNum\" = T10.\"U_DocNum\") as \"Freight Import\"," +
                   " ifnull((Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='1' and \"U_DocNum\"=T9.\"U_DocNum\"),1) as \"Invoice Value (FC)\"," +
                   " ifnull((Select (ROUND(REPLACE(\"U_AmountNPR\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='1' and \"U_DocNum\"=T9.\"U_DocNum\"),1) as \"Invoice Value\"," +
                   " (Select (ROUND(REPLACE(\"U_AmountNPR\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='17' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Registration Expenses\"," +
                   " (Select (ROUND(REPLACE(\"U_AmountNPR\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='11' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Dummarge Charges\"," +
                   " (Select (ROUND(REPLACE(\"U_AmountNPR\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='23' and \"U_DocNum\"=T9.\"U_DocNum\") as \"LC and Bank Commission Charges\"," +
                   " (Select ifnull(\"U_AmountNPR\",0) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='9' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Custom Duty\"," +
                   " (Select ifnull(\"U_AmountNPR\",0) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='10' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Excise Duty\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_AmountNPR\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('20') and \"U_DocNum\"=T9.\"U_DocNum\") as \"C&F CUSTOM - DEALER / DISTR\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_AmountNPR\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('21') and \"U_DocNum\"=T9.\"U_DocNum\") as \"C&F CUSTOM - DEALER /DIST MISC\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_AmountNPR\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('5') and \"U_DocNum\"=T9.\"U_DocNum\") as \"C&F KOLKATTA TO INDIA BORDER\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_AmountNPR\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('4') and \"U_DocNum\"=T9.\"U_DocNum\") as \"CIF KOLKOTTA\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_AmountNPR\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('24') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Storage Charges 1\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_AmountNPR\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('25') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Vehicle Upgradation Charges\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_AmountNPR\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('3') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Insurance Import\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_AmountNPR\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('6') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Insurance Transit India\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_AmountNPR\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('22') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Insurance Transit Nepal\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_AmountNPR\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('18') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Interest Cost Transit Period\"," +
                   " (Select (ROUND(REPLACE(\"U_AmountNPR\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='27' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Total Direct Cost\"," +
                   " (Select (ROUND(REPLACE(\"U_AmountNPR\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='26' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Miscellaneous Expenses\"," +
                   " (Select (ROUND(REPLACE(\"U_AmountNPR\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='19' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Handling Cost\"," +
                   " (Select (ROUND(REPLACE(\"U_AmountNPR\",',',''),2))  from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='13' and \"U_DocNum\"=T9.\"U_DocNum\") as \"PP Charges\"," +
                   " (Select (ROUND(REPLACE(\"U_AmountNPR\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='12' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Total VAT Amount\",T8.\"U_RefNo\" as \"JE No\",T12.\"DocNum\" as \"LandedCostNum\",T8.\"U_SalesMrgn\" as \"Sales Margin Per\" " +
                   " from OPDN T0 " +
                   " INNER JOIN PDN1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                   " INNER JOIN OITM T2 ON T1.\"ItemCode\" = T2.\"ItemCode\"" +
                   " LEFT JOIN OITL T3 ON T3.\"DocEntry\" = T0.\"DocEntry\" AND T3.\"DocType\" = T0.\"ObjType\" AND T3.\"DocLine\" = T1.\"LineNum\"" +
                   " LEFT JOIN ITL1 T4 ON T3.\"LogEntry\" = T4.\"LogEntry\"" +
                   " LEFT JOIN OSRN T5 ON T4.\"MdAbsEntry\" = T5.\"AbsEntry\"" +
                   " LEFT JOIN PDN3 T6 on T1.\"DocEntry\"=T6.\"DocEntry\" " +
                   " LEFT JOIN OOAT T7 ON T1.\"AgrNo\"=T7.\"AbsID\"" +
                   " LEFT JOIN \"@LANDEDCOSTH\" T8 on T0.\"DocEntry\"=T8.\"U_GRPODocEntry\" and T0.\"DocNum\"=T8.\"U_GRPODocNum\" AND T8.\"Code\" >= '00002895'" +
                   " LEFT JOIN \"@LANDEDCOSTL1\" T9 on T8.\"U_DocNum\"=T9.\"U_DocNum\"  and T9.\"U_ItemCode\"=T1.\"ItemCode\" AND T9.\"Code\" >= '00004535'" +
                   " LEFT JOIN \"@LANDEDCOSTL2\" T10 on T9.\"U_DocNum\"=T10.\"U_DocNum\" and T1.\"DocEntry\"=T10.\"U_GRPODocEntry\" and T0.\"DocNum\"=T10.\"U_GRPODocNum\" AND T10.\"Code\" >= '00087361'" +
                   " INNER JOIN OITB T11 on T2.\"ItmsGrpCod\"=T11.\"ItmsGrpCod\"" +
                    "inner join \"OIPF\" T12 ON T12.\"DocEntry\" = T0.\"LndCstNum\" " +
                   " where  T0.\"BPLId\"='2' AND  T9.\"U_DocNum\" = '" + _LandedCstNum + "'" +
                   " ORDER BY T5.\"MnfSerial\") A)B where IFNULL(\"Total VAT Amount\",0) > 0)C" +
                   " UNION ALL" +
                   " select * from(Select *,(\"Per Unit Freight Import\" + \"Per Unit Insurance Import\" + \"Per Unit C&F KOLKATTA TO INDIA BORDER\" + \"Per Unit Insurance Transit India\" " +
                   " +\"Per Unit Custom Duty\" +\"Per Unit Excise Duty\"+ \"Per Unit Registration Expenses\"+ \"Per Unit Dummarge Charges\"+ \"Per Unit PP Charges\"+ " +
                   " \"Per Unit Road Dev Tax\"+\"Per Unit Interest Cost Transit Period\"+\"Per Unit Handling Cost\"+ \"Per Unit C&F CUSTOM - DEALER / DISTR\"+ \"Per Unit C&F CUSTOM - DEALER /DIST MISC\"+ \"Per Unit Insurance Transit Nepal\" + \"Per Unit LC and Bank Commission Charges\"+ \"Per Unit Storage Charges 1\" " +
                   " +\"Per Unit Vehicle Upgradation Charges\"+ \"Per Unit Miscellaneous Expenses\") as \"Per Unit Additional Cost\"," +
                   " (\"Per Unit Freight Import\" + \"Per Unit Insurance Import\" + \"Per Unit C&F KOLKATTA TO INDIA BORDER\" + \"Per Unit Insurance Transit India\" " +
                   " +\"Per Unit Custom Duty\" +\"Per Unit Excise Duty\"+ \"Per Unit Registration Expenses\"+ \"Per Unit Dummarge Charges\"+ \"Per Unit PP Charges\"+ " +
                   " \"Per Unit Road Dev Tax\"+\"Per Unit Interest Cost Transit Period\"+\"Per Unit Handling Cost\"+ \"Per Unit C&F CUSTOM - DEALER / DISTR\"+ \"Per Unit C&F CUSTOM - DEALER /DIST MISC\"+ \"Per Unit Insurance Transit Nepal\" + \"Per Unit LC and Bank Commission Charges\"+ \"Per Unit Storage Charges 1\" " +
                   " +\"Per Unit Vehicle Upgradation Charges\"+ \"Per Unit Miscellaneous Expenses\" + \"PerUnit LineTotal\") as \"Per Unit Landing Cost\",\"CostTotal\" as \"SAP Loaded Cost\"," +
                   " (\"Per Unit Freight Import\" + \"Per Unit Insurance Import\" + \"Per Unit C&F KOLKATTA TO INDIA BORDER\" + \"Per Unit Insurance Transit India\" " +
                   " +\"Per Unit Custom Duty\" +\"Per Unit Excise Duty\"+ \"Per Unit Registration Expenses\"+ \"Per Unit Dummarge Charges\"+ \"Per Unit PP Charges\"+ " +
                   " \"Per Unit Road Dev Tax\"+\"Per Unit Interest Cost Transit Period\"+\"Per Unit Handling Cost\"+ \"Per Unit C&F CUSTOM - DEALER / DISTR\"+ \"Per Unit C&F CUSTOM - DEALER /DIST MISC\"+ \"Per Unit Insurance Transit Nepal\" + \"Per Unit LC and Bank Commission Charges\"+ \"Per Unit Storage Charges 1\" " +
                   " +\"Per Unit Vehicle Upgradation Charges\"+ \"Per Unit Miscellaneous Expenses\" + \"PerUnit LineTotal\") - \"CostTotal\" as \"Per Unit Difference\", " +
                   " ((\"CostTotal\" * \"Sales Margin Per\") / 100) as \"Sales Margin Value\" ,\"CostTotal\" + ((\"CostTotal\" * \"Sales Margin Per\") / 100) as \"Expected Sales Price\"  " +
                   " from (select distinct \"DocNum\" as \"GRPO No\",\"DocEntry\" as \"GRPO DocEntry\",\"JE No\",\"LandedCostNum\",\"Sales Margin Per\",\"DocDate\" as \"GRPO Date\",\"Supplier Code\"," +
                   " \"Supplier Name\",\"ItemCode\" as \"Item No\",\"Item Name\",\"Country of Origin\",\"DistNumber\" as \"Engine No\",\"Chasis No\",\"CostTotal\",\"LC No\",\"LC Start Date\",\"LC End Date\"," +
                   " \"Issuing Bank\",\"Landed Cost Num\",\"Costing Date\",\"PP No\",\"Invoice Value (FC)\",\"Invoice Value\" as \"Invoice Value (NP)\"," +
                   " \"DocCur\" as \"Document Currency\",\"DocRate\" as \"Exchange Rate GRPO\",\"Invoice Value\" as \"Total Invoice Value\", \"LineTotal\",\"PerUnit LineTotal\", " +
                   " Round(Replace( (\"Freight Import\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Freight Import\"," +
                   " Round(Replace( (\"Insurance Import\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Insurance Import\"," +
                   " Round(Replace( ( \"C&F KOLKATTA TO INDIA BORDER\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit C&F KOLKATTA TO INDIA BORDER\"," +
                   " Round(Replace( (\"Insurance Transit India\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Insurance Transit India\"," +
                   " Round(Replace( (\"Custom Duty\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Custom Duty\"," +
                   " Round(Replace( (\"Excise Duty\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Excise Duty\"," +
                   " Round(Replace( (\"Dummarge Charges\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Dummarge Charges\"," +
                   " Round(Replace( (\"PP Charges\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit PP Charges\"," +
                   " Round(Replace( (\"Road Development Tax\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Road Dev Tax\", " +
                   " Round(Replace( (\"Registration Expenses\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Registration Expenses\"," +
                   " Round(Replace( (\"Interest Cost Transit Period\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Interest Cost Transit Period\"," +
                   " Round(Replace( (\"Handling Cost\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Handling Cost\"," +
                   " Round(Replace( ( \"C&F CUSTOM - DEALER / DISTR\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit C&F CUSTOM - DEALER / DISTR\"," +
                   " Round(Replace( ( \"C&F CUSTOM - DEALER /DIST MISC\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit C&F CUSTOM - DEALER /DIST MISC\"," +
                   " Round(Replace( ( \"Insurance Transit Nepal\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Insurance Transit Nepal\"," +
                   " Round(Replace( (\"LC and Bank Commission Charges\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit LC and Bank Commission Charges\"," +
                   " Round(Replace( ( \"Storage Charges 1\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Storage Charges 1\"," +
                   " Round(Replace( ( \"Vehicle Upgradation Charges\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Vehicle Upgradation Charges\"," +
                   " Round(Replace( (\"Miscellaneous Expenses\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit Miscellaneous Expenses\"," +
                   " Round(Replace( (\"Total VAT Amount\"/case when IFNULL(\"Invoice Value\",0) > 0 then \"Invoice Value\" else 1 END)*(\"PerUnit LineTotal\"),',',''),2) As \"Per Unit VAT Amount\" ,\"Total VAT Amount\" " +
                   " from (SELECT distinct T0.\"DocDate\",T0.\"DocCur\",T0.\"DocRate\",T0.\"DocNum\",T0.\"DocEntry\",(T1.\"LineTotal\"+T1.\"DistribSum\") as \"LineTotal\", " +
                   " ((T1.\"LineTotal\"+T1.\"DistribSum\")/T1.\"Quantity\") as \"PerUnit LineTotal\",T0.\"NumAtCard\" as \"Vendor Ref No\", " +
                   " T0.\"U_CountryN\" as \"Country of Origin\", " +
                   " T0.\"DocTotal\",T0.\"CardCode\" as \"Supplier Code\",T0.\"CardName\" as \"Supplier Name\",T9.\"U_DocNum\" as \"Landed Cost Num\", " +
                   " T5.\"ItemCode\",T5.\"itemName\" as \"Item Name\",T11.\"ItmsGrpNam\" as \"Item Group\", T5.\"DistNumber\", T5.\"MnfSerial\" as \"Chasis No\",T5.\"CostTotal\", " +
                   " T7.\"StartDate\" as \"LC Start Date\", T7.\"EndDate\" as \"LC End Date\", " +
                   " T7.\"Descript\" as \"LC No\",T7.\"U_BankName\" as \"Issuing Bank\",T8.\"U_DocNum\" as \"Landed Cost No\",T8.\"U_PostingDate\" as \"Costing Date\", " +
                   " T8.\"U_PPNo\" as \"PP No\", " +
                   " ifnull(T9.\"U_TotExDuty\",0) as \"Total Excise Duty\", " +
                   " ifnull(T9.\"U_TotCusDuty\",0) as \"Total Custom Duty\", " +
                   " ifnull(T9.\"U_Quantity\",0) as \"U_Quantity\", " +
                   " ifnull(T9.\"U_TotRDTax\",0) as \"Total Road Dev Tax\", " +
                   " (Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='4' and \"U_DocNum\" = T10.\"U_DocNum\") as \"Freight Import\", " +
                   " ifnull((Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='3' and \"U_DocNum\"=T9.\"U_DocNum\"),1) as \"Invoice Value (FC)\", " +
                   " ifnull((Select (ROUND(REPLACE(\"U_Amount\",',',''),2)*T0.\"DocRate\") from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='3' and \"U_DocNum\"=T9.\"U_DocNum\"),1) as \"Invoice Value\", " +
                   " (Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='19' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Registration Expenses\", " +
                   "(Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='18' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Road Development Tax\"," +
                   " (Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='13' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Dummarge Charges\", " +
                   " (Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='25' and \"U_DocNum\"=T9.\"U_DocNum\") as \"LC and Bank Commission Charges\", " +
                   " (Select ifnull(\"U_Amount\",0) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='11' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Custom Duty\", " +
                   " (Select ifnull(\"U_Amount\",0) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='12' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Excise Duty\", " +
                   " (Select (SUM(ROUND(REPLACE(\"U_Amount\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('22') and \"U_DocNum\"=T9.\"U_DocNum\") as \"C&F CUSTOM - DEALER / DISTR\", " +
                   " (Select (SUM(ROUND(REPLACE(\"U_Amount\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('23') and \"U_DocNum\"=T9.\"U_DocNum\") as \"C&F CUSTOM - DEALER /DIST MISC\", " +
                   " (Select (SUM(ROUND(REPLACE(\"U_Amount\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('7') and \"U_DocNum\"=T9.\"U_DocNum\") as \"C&F KOLKATTA TO INDIA BORDER\", " +
                   " (Select (SUM(ROUND(REPLACE(\"U_Amount\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('6') and \"U_DocNum\"=T9.\"U_DocNum\") as \"CIF KOLKOTTA\", " +
                   " (Select (SUM(ROUND(REPLACE(\"U_Amount\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('26') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Storage Charges 1\", " +
                   " (Select (SUM(ROUND(REPLACE(\"U_Amount\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('27') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Vehicle Upgradation Charges\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_Amount\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('5') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Insurance Import\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_Amount\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('8') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Insurance Transit India\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_Amount\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('24') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Insurance Transit Nepal\"," +
                   " (Select (SUM(ROUND(REPLACE(\"U_Amount\",',',''),2))) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\" in ('20') and \"U_DocNum\"=T9.\"U_DocNum\") as \"Interest Cost Transit Period\"," +
                   " (Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='29' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Total Direct Cost\"," +
                   " (Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='28' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Miscellaneous Expenses\"," +
                   " (Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='21' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Handling Cost\"," +
                   " (Select (ROUND(REPLACE(\"U_Amount\",',',''),2))  from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='15' and \"U_DocNum\"=T9.\"U_DocNum\") as \"PP Charges\"," +
                   " (Select (ROUND(REPLACE(\"U_Amount\",',',''),2)) from \"@LANDEDCOSTL2\" where \"U_LandCstCod\"='14' and \"U_DocNum\"=T9.\"U_DocNum\") as \"Total VAT Amount\",T8.\"U_RefNo\" as \"JE No\",T12.\"DocNum\" as \"LandedCostNum\",T8.\"U_SalesMrgn\" as \"Sales Margin Per\" " +
                   " from OPDN T0 " +
                   " INNER JOIN PDN1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                   " INNER JOIN OITM T2 ON T1.\"ItemCode\" = T2.\"ItemCode\" " +
                   " LEFT JOIN OITL T3 ON T3.\"DocEntry\" = T0.\"DocEntry\" AND T3.\"DocType\" = T0.\"ObjType\" AND T3.\"DocLine\" = T1.\"LineNum\" " +
                   " LEFT JOIN ITL1 T4 ON T3.\"LogEntry\" = T4.\"LogEntry\" " +
                   " LEFT JOIN OSRN T5 ON T4.\"MdAbsEntry\" = T5.\"AbsEntry\" " +
                   " LEFT JOIN PDN3 T6 on T1.\"DocEntry\"=T6.\"DocEntry\"  " +
                   " LEFT JOIN OOAT T7 ON T1.\"AgrNo\"=T7.\"AbsID\" " +
                   " LEFT JOIN \"@LANDEDCOSTH\" T8 on T0.\"DocEntry\"=T8.\"U_GRPODocEntry\" and T0.\"DocNum\"=T8.\"U_GRPODocNum\" AND T8.\"Code\" < '00002895' " +
                   " LEFT JOIN \"@LANDEDCOSTL1\" T9 on T8.\"U_DocNum\"=T9.\"U_DocNum\"  and T9.\"U_ItemCode\"=T1.\"ItemCode\" AND T9.\"Code\" < '00004535' " +
                   " LEFT JOIN \"@LANDEDCOSTL2\" T10 on T9.\"U_DocNum\"=T10.\"U_DocNum\" and T1.\"DocEntry\"=T10.\"U_GRPODocEntry\" and T0.\"DocNum\"=T10.\"U_GRPODocNum\" AND T10.\"Code\" < '00087361' " +
                   " INNER JOIN OITB T11 on T2.\"ItmsGrpCod\"=T11.\"ItmsGrpCod\" " +
                   "inner join \"OIPF\" T12 ON T12.\"DocEntry\" = T0.\"LndCstNum\" "+
                   " where  T0.\"BPLId\"='2' AND  T9.\"U_DocNum\" = '" + _LandedCstNum + "' " +
                   " ORDER BY T5.\"MnfSerial\") A  )B where IFNULL(\"Total VAT Amount\",0) > 0)C order by \"GRPO No\" ";

            _GridDataTable.ExecuteQuery(SQL1);
            _Grid.DataTable = _GridDataTable;
            _RowHeaders = _Grid.RowHeaders;
            _RowHeaders.TitleObject.Caption = "#";

            for (int i = 0; i < _Grid.Rows.Count; i++)
            {
                _RowHeaders.SetText(i, (i + 1).ToString());

            }
            _Form.Mode = BoFormMode.fm_VIEW_MODE;
            //_Grid.Columns.Item("Code").Width = 0;
            _Column = (SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(1);
            _Column.LinkedObjectType = "20";
            _Column = (SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(2);
            _Column.LinkedObjectType = "30";
            _Column = (SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(3);
            _Column.LinkedObjectType = "69";
            //((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(25)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(26)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(27)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(28)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(29)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(30)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(31)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(32)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(33)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(34)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(35)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(36)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(37)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(38)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(39)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(40)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(41)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(42)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(43)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(44)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(45)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(46)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(47)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(48)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(49)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(50)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(51)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
            ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(52)).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
        }

        #endregion

    }
}
