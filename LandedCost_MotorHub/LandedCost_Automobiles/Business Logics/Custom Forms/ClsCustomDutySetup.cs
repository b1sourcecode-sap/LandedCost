using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Sap.Data.Hana;
using SAPbobsCOM;
using SAPbouiCOM;
namespace LandedCost_Automobiles
{
    class ClsCustomDutySetup : Base
    {

        #region Variable Declaration
        private SAPbouiCOM.Matrix _Matrix;
        private SAPbouiCOM.DBDataSource _DBDataSource, _DBDataSourceH;
        private SAPbobsCOM.Recordset _Rs = null;
        string SQL = string.Empty;
        private SAPbouiCOM.EditText _EditText;
        int _LastErrorCode = 0;
        string _LastErrorMsg = " ";
        public static string _PONum = string.Empty;
        public static string _PODocEntry = string.Empty;
        public static string _AfterLandingCode = string.Empty;
        SAPbouiCOM.Condition oCon;
        SAPbouiCOM.Conditions oCons;
        SAPbouiCOM.ComboBox _ComboBox;

        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;

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
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            if (pVal.ItemUID == "1" && (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                            {
                                if (this.SaveData() == false)
                                {
                                    oBubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    oBubbleEvent = false;
                                }

                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            if (pVal.ItemUID == "8")
                            {
                                this.FilteGLAccount("");
                            }
                            if (pVal.ItemUID == "10")
                            {
                                this.FilteGLAccount("");
                            }
                            if (pVal.ItemUID == "13")
                            {
                                this.FilteGLAccount("");
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
                switch (pVal.EventType)
                {
                    #region Set CFL Values
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "8" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _DBDataSourceH.SetValue("U_VATINAC", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());

                        }
                        if (pVal.ItemUID == "10" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _DBDataSourceH.SetValue("U_CustDtyAC", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());

                        }
                        if (pVal.ItemUID == "13" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _DBDataSourceH.SetValue("U_ExDtyAC", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());

                        }
                        

                        break;
                    #endregion
                }
            }
            #endregion
        }

        #endregion

        #region From Default
        internal void FormDefault()
        {
            try
            {
                _Form.Freeze(true);
                _DBDataSourceH = this._Form.DataSources.DBDataSources.Item("@CSTMDSTPH");

                //Header Value Check
                string _SQL = "Select *  from \"@CSTMDSTPH\"";
                Utilities.ExecuteSQL(ref _Rs, _SQL);
                if (_Rs.RecordCount > 0)
                {
                    ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = _Rs.Fields.Item("U_VATPer").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("8").Specific).Value = _Rs.Fields.Item("U_VATINAC").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = _Rs.Fields.Item("U_CustDtyAC").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("13").Specific).Value = _Rs.Fields.Item("U_ExDtyAC").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value = _Rs.Fields.Item("U_AdyAllAC").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("15").Specific).Value = _Rs.Fields.Item("U_SalesMrgn").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("144").Specific).Value = _Rs.Fields.Item("U_RDTAXAC").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("145").Specific).Value = _Rs.Fields.Item("U_INTPDCS").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("146").Specific).Value = _Rs.Fields.Item("U_PPCHRGAC").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("147").Specific).Value = _Rs.Fields.Item("U_DUMRAGAC").Value.ToString();
                }

                //End region

                if (dataexist() == true)
                {
                  
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                }
                
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
                _Form.Visible = true;
            }
        }
        #endregion

        #region Check Data in table
        private bool dataexist()
        {
            try
            {
                SQL = "Select Count(*)  from \"@CSTMDSTPL\"";
                Utilities.ExecuteSQL(ref _Rs, SQL);
                if (_Rs.Fields.Item(0).Value.ToString() != "" && Convert.ToInt32(_Rs.Fields.Item(0).Value.ToString()) > 0)
                {
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        #endregion

        #region  Conditional Reload data in Form
        private void ConditionalReloadData()
        {
            try
            {
                //_Form.Freeze(true);
                oCons = new SAPbouiCOM.Conditions();
                oCon = oCons.Add();
                _DBDataSource.Query(null);
                _Matrix.LoadFromDataSource();
                _DBDataSource.Clear();
                _Matrix.AddRow(1, _Matrix.RowCount);

               
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                // _Form.Freeze(false); 
            }
        }
        #endregion

        #region Save Data

        private bool SaveData()
        {
            SAPbobsCOM.UserTable oUserTable = Utilities.Application.Company.UserTables.Item("CSTMDSTPL");
            SAPbobsCOM.UserTable oUserTableH = Utilities.Application.Company.UserTables.Item("CSTMDSTPH");

            //Header Text box Values
            string MaxValueH = "00001";
            if (oUserTableH.GetByKey(MaxValueH)) // Code 
            {
                oUserTableH.UserFields.Fields.Item("U_VATPer").Value = _DBDataSourceH.GetValue("U_VATPer", 0).ToString().Trim();
                oUserTableH.UserFields.Fields.Item("U_VATINAC").Value = _DBDataSourceH.GetValue("U_VATINAC", 0).ToString().Trim();
                oUserTableH.UserFields.Fields.Item("U_CustDtyAC").Value = _DBDataSourceH.GetValue("U_CustDtyAC", 0).ToString().Trim();
                oUserTableH.UserFields.Fields.Item("U_ExDtyAC").Value = _DBDataSourceH.GetValue("U_ExDtyAC", 0).ToString().Trim();
                oUserTableH.UserFields.Fields.Item("U_AdyAllAC").Value = _DBDataSourceH.GetValue("U_AdyAllAC", 0).ToString().Trim();
                oUserTableH.UserFields.Fields.Item("U_SalesMrgn").Value = _DBDataSourceH.GetValue("U_SalesMrgn", 0).ToString().Trim();
                oUserTableH.UserFields.Fields.Item("U_RDTAXAC").Value = _DBDataSourceH.GetValue("U_RDTAXAC", 0).ToString().Trim();
                oUserTableH.UserFields.Fields.Item("U_INTPDCS").Value = _DBDataSourceH.GetValue("U_INTPDCS", 0).ToString().Trim();
                oUserTableH.UserFields.Fields.Item("U_DUMRAGAC").Value = _DBDataSourceH.GetValue("U_DUMRAGAC", 0).ToString().Trim();
                oUserTableH.UserFields.Fields.Item("U_PPCHRGAC").Value = _DBDataSourceH.GetValue("U_PPCHRGAC", 0).ToString().Trim();
                //oUserTableH.UserFields.Fields.Item("U_RoadTax").Value = _DBDataSourceH.GetValue("U_RoadTax", 0).ToString().Trim();
                int rtnVal = oUserTableH.Update(); // Update Table !
                if (rtnVal != 0)
                {
                    Utilities.Application.Company.GetLastError(out _LastErrorCode, out _LastErrorMsg);
                    return false;
                }
            }
            else
                if (_DBDataSourceH.GetValue("U_VATPer", 0).ToString().Trim() != "")
                {
                    oUserTableH.Code = MaxValueH; // Assign next Code value 
                    oUserTableH.Name = MaxValueH;// Assign next Code value 
                    oUserTableH.UserFields.Fields.Item("U_VATPer").Value = _DBDataSourceH.GetValue("U_VATPer", 0).ToString().Trim();
                    oUserTableH.UserFields.Fields.Item("U_VATINAC").Value = _DBDataSourceH.GetValue("U_VATINAC", 0).ToString().Trim();
                    oUserTableH.UserFields.Fields.Item("U_CustDtyAC").Value = _DBDataSourceH.GetValue("U_CustDtyAC", 0).ToString().Trim();
                    oUserTableH.UserFields.Fields.Item("U_ExDtyAC").Value = _DBDataSourceH.GetValue("U_ExDtyAC", 0).ToString().Trim();
                    oUserTableH.UserFields.Fields.Item("U_AdyAllAC").Value = _DBDataSourceH.GetValue("U_AdyAllAC", 0).ToString().Trim();
                    oUserTableH.UserFields.Fields.Item("U_SalesMrgn").Value = _DBDataSourceH.GetValue("U_SalesMrgn", 0).ToString().Trim();
                    oUserTableH.UserFields.Fields.Item("U_RDTAXAC").Value = _DBDataSourceH.GetValue("U_RDTAXAC", 0).ToString().Trim();
                    oUserTableH.UserFields.Fields.Item("U_INTPDCS").Value = _DBDataSourceH.GetValue("U_INTPDCS", 0).ToString().Trim();
                    oUserTableH.UserFields.Fields.Item("U_DUMRAGAC").Value = _DBDataSourceH.GetValue("U_DUMRAGAC", 0).ToString().Trim();
                    oUserTableH.UserFields.Fields.Item("U_PPCHRGAC").Value = _DBDataSourceH.GetValue("U_PPCHRGAC", 0).ToString().Trim();
                //oUserTableH.UserFields.Fields.Item("U_RoadTax").Value = _DBDataSourceH.GetValue("U_RoadTax", 0).ToString().Trim();
                int rtnVal = oUserTableH.Add();
                    if (rtnVal != 0)
                    {
                        Utilities.Application.Company.GetLastError(out _LastErrorCode, out _LastErrorMsg);
                        return false;
                    }
                }

            
            return true;
        }


        #endregion

        #region Validate Form
        private bool ValidateFrom()
        {
            try
            {
                _Matrix.FlushToDataSource();
               
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                return false;
            }
            return true;
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
                _Con.Alias = "Postable";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "Y";
                _Con.BracketCloseNum = 1;

                _Form.ChooseFromLists.Item("CFL_1").SetConditions(_Cons);
                _Form.ChooseFromLists.Item("CFL_2").SetConditions(_Cons);
                _Form.ChooseFromLists.Item("CFL_3").SetConditions(_Cons);

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
    }
}
