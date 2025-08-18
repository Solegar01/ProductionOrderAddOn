using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;

namespace ProductionOrderAddOn
{
    public class ProdOrderListForm
    {
        private SAPbouiCOM.Form _form;
        private SAPbouiCOM.EditText _fromDate;
        private SAPbouiCOM.EditText _toDate;
        private SAPbouiCOM.EditText _searchText;
        private SAPbouiCOM.Button _btnFilter;
        private SAPbouiCOM.Grid _grid;

        public void Show()
        {
            try
            {
                try
                {
                    var existingForm = Application.SBO_Application.Forms.Item("WorkOrderList");
                    existingForm.Select();
                    return;
                }
                catch
                {
                    
                }

                var formParams = (SAPbouiCOM.FormCreationParams)
                Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                formParams.UniqueID = "WorkOrderList";
                formParams.FormType = "PO_LIST";
                formParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;

                _form = Application.SBO_Application.Forms.AddEx(formParams);
                _form.Title = "Production Order List";
                _form.Width = 600;
                _form.Height = 400;

                // Data Sources FIRST (only once)
                _form.DataSources.UserDataSources.Add("fromDateDS", SAPbouiCOM.BoDataType.dt_DATE);
                _form.DataSources.UserDataSources.Add("toDateDS", SAPbouiCOM.BoDataType.dt_DATE);
                _form.DataSources.DataTables.Add("WOData");
                _form.DataSources.UserDataSources.Add("SearchDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);


                // From Date input
                var fromDateItem = _form.Items.Add("FromDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                fromDateItem.Top = 10; fromDateItem.Left = 100; fromDateItem.Width = 100;
                _fromDate = (SAPbouiCOM.EditText)fromDateItem.Specific;
                _fromDate.DataBind.SetBound(true, "", "fromDateDS");

                // From Date label
                var lblFrom = _form.Items.Add("lblFrom", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                lblFrom.Top = 10; lblFrom.Left = 10; lblFrom.Width = 80;
                ((SAPbouiCOM.StaticText)lblFrom.Specific).Caption = "From Date";
                lblFrom.LinkTo = "FromDate";

                // To Date input
                var toDateItem = _form.Items.Add("ToDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                toDateItem.Top = 10; toDateItem.Left = 300; toDateItem.Width = 100;
                _toDate = (SAPbouiCOM.EditText)toDateItem.Specific;
                _toDate.DataBind.SetBound(true, "", "toDateDS");

                // To Date label
                var lblTo = _form.Items.Add("lblTo", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                lblTo.Top = 10; lblTo.Left = 220; lblTo.Width = 80;
                ((SAPbouiCOM.StaticText)lblTo.Specific).Caption = "To Date";
                lblTo.LinkTo = "ToDate";
                
                // 2. Add EditText control (position x=10, y=10 for example)
                SAPbouiCOM.Item searchItem = _form.Items.Add("SearchTxt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                searchItem.Left = 100;
                searchItem.Top = 30;
                searchItem.Width = 100;
                searchItem.Height = 14;

                _searchText = (SAPbouiCOM.EditText)searchItem.Specific;
                _searchText.DataBind.SetBound(true, "", "SearchDS");

                // Label for the search box
                SAPbouiCOM.Item lblItem = _form.Items.Add("lblSearch", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                lblItem.Left = 10;
                lblItem.Top = 30;
                lblItem.Width = 70;
                lblItem.Height = 14;

                SAPbouiCOM.StaticText lblSearch = (SAPbouiCOM.StaticText)lblItem.Specific;
                lblSearch.Caption = "Search";

                // 5. Link label to edit text
                lblSearch.Item.LinkTo = "SearchTxt";

                // Filter button
                var btnItem = _form.Items.Add("FilterBtn", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                btnItem.Top = 10; btnItem.Left = 420; btnItem.Width = 80;
                _btnFilter = (SAPbouiCOM.Button)btnItem.Specific;
                _btnFilter.Caption = "Filter";

                // Grid
                var gridItem = _form.Items.Add("Grid1", SAPbouiCOM.BoFormItemTypes.it_GRID);
                gridItem.Top = 50; gridItem.Left = 10; gridItem.Width = 560; gridItem.Height = 320;
                _grid = (SAPbouiCOM.Grid)gridItem.Specific;
                _grid.DataTable = _form.DataSources.DataTables.Item("WOData");

                // Subscribe to events only once
                Application.SBO_Application.ItemEvent += OnItemEvent;

                _form.Visible = true;
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void OnItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormUID == _form.UniqueID && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction)
            {
                if (pVal.ItemUID == "FilterBtn")
                {
                    LoadProductionOrders();
                }
            }

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && !pVal.BeforeAction)
            {
                // Check Enter key
                if (pVal.CharPressed == 13)
                {
                    if (pVal.ItemUID == "FromDate" || pVal.ItemUID == "ToDate" || pVal.ItemUID == "SearchTxt")
                    {
                        LoadProductionOrders();
                    }
                }
            }

            //if (pVal.FormUID == _form.UniqueID
            //    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
            //    && pVal.BeforeAction) // only after press
            //{
            //    if (pVal.ItemUID == "Grid1" && pVal.ColUID == "Doc. No" && pVal.Row >= 0)
            //    {
            //        var dt = _grid.DataTable;
            //        string docEntry = dt.GetValue("DocEntry", pVal.Row).ToString();
                    
            //        Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_ProductionOrder, "", docEntry);
            //        BubbleEvent = false; // cancel default link behavior
            //    }
            //}

        }

        private void LoadProductionOrders()
        {
            try
            {
                _form.Freeze(true); // Stop screen updates
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(
                    "Loading Imported Production Orders...",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                );

                string fromDate = _fromDate.Value.Trim();
                string toDate = _toDate.Value.Trim();
                string searchText = _searchText.Value.Trim(); // Your search input

                int currentYear = DateTime.Now.Year;
                string today = DateTime.Now.ToString("yyyyMMdd");
                string select = "SELECT DocEntry, DocNum [Doc. No], DueDate [Order Date], ItemCode [Item Code], PlannedQty [Planned Qty], Status, ISNULL(U_T2_PRODTYPE, '') [Prod. Type] FROM OWOR ";
                string where = "WHERE ISNULL(U_T2_Is_Import,'') = 'Y' ";

                // Add search filter if not empty
                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    string safeSearch = searchText.Replace("'", "''");
                    where += $" AND (CAST(DocNum AS NVARCHAR) LIKE '%{safeSearch}%' OR ItemCode LIKE '%{safeSearch}%') ";
                }

                string query;

                if (string.IsNullOrWhiteSpace(fromDate) && string.IsNullOrWhiteSpace(toDate))
                {
                    query = $@"
        {select}
        {where} AND YEAR(DueDate) = {currentYear}
        ORDER BY DueDate";
                }
                else if (!string.IsNullOrWhiteSpace(fromDate) && string.IsNullOrWhiteSpace(toDate))
                {
                    query = $@"
        {select}
        {where} AND DueDate BETWEEN '{fromDate}' AND '{today}'
        ORDER BY DueDate";
                }
                else if (string.IsNullOrWhiteSpace(fromDate) && !string.IsNullOrWhiteSpace(toDate))
                {
                    string startOfYear = new DateTime(currentYear, 1, 1).ToString("yyyyMMdd");
                    query = $@"
        {select}
        {where} AND DueDate BETWEEN '{startOfYear}' AND '{toDate}'
        ORDER BY DueDate";
                }
                else
                {
                    query = $@"
        {select}
        {where} AND DueDate BETWEEN '{fromDate}' AND '{toDate}'
        ORDER BY DueDate";
                }

                var dt = _form.DataSources.DataTables.Item("WOData");
                dt.ExecuteQuery(query);
                //_grid.DataTable = dt;

                // Enable sort on all columns
                for (int i = 0; i < _grid.Columns.Count; i++)
                {
                    _grid.Columns.Item(i).TitleObject.Sortable = true;
                    _grid.Columns.Item(i).Editable = false;
                }

                FormatGrid();
                _grid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                _grid.AutoResizeColumns();
                _grid.CollapseLevel = 0;

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(
                   "Imported Production Orders loaded successfully",
                   SAPbouiCOM.BoMessageTime.bmt_Short,
                   SAPbouiCOM.BoStatusBarMessageType.smt_Success
               );
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(
                    $"Error: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Medium,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                _form.Freeze(false);
            }
        }

        private void FormatGrid()
        {
            //_grid.Columns.Item("DocEntry").Visible = false;
            // Set link to open existing Production Order
            //var docEntryCol = (SAPbouiCOM.EditTextColumn)_grid.Columns.Item("Doc. No");
            var docEntryCol = (SAPbouiCOM.EditTextColumn)_grid.Columns.Item("DocEntry");
            docEntryCol.LinkedObjectType = "202"; // 202 = Production Order
        }
    }
}
