using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ProductionOrderAddOn.Services;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace ProductionOrderAddOn
{
    class Menu
    {
        private string _oldStatus = "";
        private double _oldQty = 0;
        private SAPbouiCOM.ProgressBar _pb;

        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "ProductionOrderAddOn";
            oCreationPackage.String = "Production Order Add On";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;
            
            Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
            Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception e)
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("ProductionOrderAddOn");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "ProductionOrderAddOn.ImportFile";
                oCreationPackage.String = "Import File Production";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception er)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (BusinessObjectInfo.FormTypeEx == "65211" // Production Order
                    && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    && !BusinessObjectInfo.BeforeAction
                    && BusinessObjectInfo.ActionSuccess)
                {
                    if (!string.IsNullOrWhiteSpace(BusinessObjectInfo.ObjectKey))
                    {
                        int docEntry = ExtractDocEntry(BusinessObjectInfo.ObjectKey); // from AbsoluteEntry
                        CancelSubOrder(docEntry);
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText( ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private int ExtractDocEntry(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml))
                throw new Exception("ObjectKey is empty, cannot extract DocEntry.");

            var doc = new System.Xml.XmlDocument();
            doc.LoadXml(xml);

            // Production Order uses AbsoluteEntry
            var node = doc.SelectSingleNode("//AbsoluteEntry");
            if (node == null || string.IsNullOrWhiteSpace(node.InnerText))
                throw new Exception("AbsoluteEntry node not found in ObjectKey XML.");

            return int.Parse(node.InnerText);
        }



        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "ProductionOrderAddOn.ImportFile")
                {
                    ImportForm activeForm = new ImportForm();
                    activeForm.Show();
                }
                ////Check if it's the Cancel menu (1283 is the system menu ID for Cancel)
                //if (pVal.MenuUID == "1284" && !pVal.BeforeAction)
                //{
                //    // Check if the active form is Production Order
                //    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                //    if (oForm.TypeEx == "65211") // 65211 = Production Order form type
                //    {
                //        CancelSubOrder();

                //    }
                //}
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }
        
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            // Production Order Form
            if (pVal.FormTypeEx == "65211")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction && pVal.ItemUID == "1")
                {
                    UpdateVarHandler(FormUID);
                }
                
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ActionSuccess && pVal.ItemUID == "1")
                {
                    GenerateHandler(FormUID);
                }
                
            }
        }

        private void RefreshFormProdOrder(SAPbouiCOM.Form oForm, int docEntry, string msg)
        {
            try
            {
                string sDocEntry = docEntry.ToString();
                oForm.Close(); // Close current form

                // Timer delay 1 second before reopening
                System.Timers.Timer reopenTimer = new System.Timers.Timer(1000);
                reopenTimer.AutoReset = false;
                reopenTimer.Elapsed += (sender, e) =>
                {
                    try
                    {
                        Application.SBO_Application.OpenForm(
                            SAPbouiCOM.BoFormObjectEnum.fo_ProductionOrder,
                            "",
                            sDocEntry);

                        Application.SBO_Application.StatusBar.SetText(
                            msg,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    catch (Exception exOpen)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Failed to reopen Production Order form: " + exOpen.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Long,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    finally
                    {
                        reopenTimer.Stop();
                        reopenTimer.Dispose();
                    }
                };
                reopenTimer.Start();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error during refresh: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void UpdateVarHandler(string FormUID)
        {
            try
            {
                Company oCompany = Services.CompanyService.GetCompany();
                if(_pb != null) { _pb.Stop();_pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");
                string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                
                int docEntry;
                if (int.TryParse(docEntryStr, out docEntry))
                {
                    var oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS.DoQuery($"SELECT Status,PlannedQty FROM OWOR WHERE DocEntry = {docEntry}");
                    if (!oRS.EoF)
                    {
                        _oldStatus = oRS.Fields.Item("Status").Value.ToString();
                        _oldQty = (double)oRS.Fields.Item("PlannedQty").Value;
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
            }
        }

        private void GenerateHandler(string FormUID)
        {
            int docEntry = 0;
            bool isGenerate = false;
            Company oCompany = null;
            try
            {
                oCompany = Services.CompanyService.GetCompany();
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                oCompany.StartTransaction();
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");

                string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                if (int.TryParse(docEntryStr, out docEntry))
                {
                    var oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS.DoQuery($"SELECT Status, ISNULL(U_T2_PRODTYPE,'') AS ProdType, ISNULL(PlannedQty, 0) AS PlannedQty, ISNULL(U_T2_Is_Import,'N') AS IsImported FROM OWOR WHERE DocEntry = {docEntry}");
                    string newStatus = oRS.Fields.Item("Status").Value.ToString();
                    string prodType = oRS.Fields.Item("ProdType").Value.ToString();
                    string isImported = oRS.Fields.Item("IsImported").Value.ToString();
                    double plannedQty = (double)oRS.Fields.Item("PlannedQty").Value;

                    // Cek perubahan
                    if (_oldStatus != newStatus && newStatus == "R" && prodType == "FG" && isImported == "N")
                    {
                        isGenerate = true;
                        // After update sukses, ambil status baru dan bandingkan
                        int result = Application.SBO_Application.MessageBox(
                            "This action will generate sub production orders. Do you want to continue?",
                            1, // default button = first option
                            "Yes",
                            "No"
                        );

                        if (result == 1) // User clicked "Yes"
                        {
                            // Tambahkan proses kamu di sini
                            if (_pb != null) { _pb.Stop(); _pb = null; }
                            _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Generating sub-orders...", 100, false);

                            // Generate suborder
                            _pb.Value = 0;
                            _pb.Text = "Generating Sub Production Orders...";
                            var listDoc = ProductionOrderSapService.GenerateSubOrder(oCompany, docEntry);
                            foreach (var item in listDoc)
                            {
                                int wipEntry = item.Key;
                                ProductionOrderSapService.LinkWipToFG(oCompany, docEntry, wipEntry);
                            }

                            if (listDoc != null && listDoc.Any())
                            {
                                string remarks = "Sub Production Orders: " + string.Join(" | ", listDoc.Values);
                                ProductionOrderSapService.UpdateRemarks(oCompany, docEntry, remarks);

                            }

                            _pb.Value = 100; // Ensure it reaches the end
                            _pb.Text = "Done.";
                            _pb.Stop();

                            RefreshFormProdOrder(oForm, docEntry, "Sub production orders successfully genareted.");
                        }
                        else
                        {
                            ResetStatus(oCompany, FormUID);
                        }
                    }
                    else if (newStatus == "R" && prodType == "FG" && _oldQty != plannedQty)
                    {
                        isGenerate = false;
                        int result = Application.SBO_Application.MessageBox(
                            "This action will affect the related sub production orders. Do you want to continue?",
                            1, // default button = first option
                            "Yes",
                            "No"
                        );
                        if (result == 1) // User clicked "Yes"
                        {
                            // Tambahkan proses kamu di sini
                            if (_pb != null) { _pb.Stop(); _pb = null; }
                            _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Updating sub-orders...", 100, false);

                            // Generate suborder
                            _pb.Value = 0;
                            _pb.Text = "Updating WIP Production Orders...";
                            if (!ProductionOrderSapService.UpdateSubOrder(oCompany, docEntry))
                                throw new Exception("There is no documents updated.");

                            _pb.Value = 100; // Ensure it reaches the end
                            _pb.Text = "Done.";
                            _pb.Stop();

                            RefreshFormProdOrder(oForm, docEntry, "Sub production orders successfully updated.");
                        }
                        else
                        {
                            ResetQty(oCompany, FormUID);
                        }
                    }
                }
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                }

            }
            catch (Exception ex)
            {
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                {
                    System.Threading.Thread.Sleep(1000); 
                    Application.SBO_Application.StatusBar.SetText(ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                });
                
                if (isGenerate)
                    ResetStatus(oCompany, FormUID);
                else
                    ResetQty(oCompany, FormUID);
            }
            finally
            {
                if (_pb != null) { _pb.Stop();_pb = null; }
            }
        }
        
        private void ResetStatus(Company oCompany, string FormUID)
        {
            int docEntry = 0;
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");

                string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                if (int.TryParse(docEntryStr, out docEntry))
                {
                    ProductionOrderSapService.UpdatePoStatus(oCompany,docEntry, BoProductionOrderStatusEnum.boposPlanned);
                    RefreshFormProdOrder(oForm, docEntry, $"Status reverted to planned.");
                }

            }
            catch (Exception)
            {
                throw;
            }
        }

        private void ResetQty(Company oCompany, string FormUID)
        {
            int docEntry = 0;
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");

                string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                if (int.TryParse(docEntryStr, out docEntry))
                {
                    ProductionOrderSapService.UpdatePoQty(oCompany, docEntry, _oldQty);
                    RefreshFormProdOrder(oForm, docEntry, $"Planned Quantity reverted to previous value: {_oldQty}");
                }

            }
            catch (Exception)
            {
                throw;
            }
        }

        private void CancelSubOrder(int docEntry)
        {
            Company oCompany = null;
            try
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Cancelling sub-orders...", 100, false);

                _pb.Value = 0;
                _pb.Text = "Cancelling Sub Production Orders...";
                oCompany = Services.CompanyService.GetCompany();
                oCompany.StartTransaction();

                var rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery($"SELECT Status FROM OWOR WHERE DocEntry = {docEntry}");

                if (rs.Fields.Item("Status").Value.ToString() == "C")
                {
                    ProductionOrderSapService.CancelSubOrder(oCompany, docEntry);
                    System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                    {
                        System.Threading.Thread.Sleep(1000);
                        Application.SBO_Application.StatusBar.SetText(
                            $"Sub Production Orders were cancelled.",
                            SAPbouiCOM.BoMessageTime.bmt_Long,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    });
                }

                _pb.Value = 100;
                _pb.Text = "Done.";
                _pb.Stop();

                if (oCompany.InTransaction)
                    oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
            }
            catch (Exception)
            {
                if (oCompany.InTransaction)
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
            }
        }

    }
}
