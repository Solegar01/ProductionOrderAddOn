using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ProductionOrderAddOn.Services;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace ProductionOrderAddOn
{
    class Menu
    {
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
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

        // Temp variable di luar method (class-level)
        private string _statusLama = "";

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.ProgressBar progress = null;
            Company oCompany = null;
            // Production Order Form
            if (pVal.FormTypeEx == "65211")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction && pVal.ItemUID == "1")
                {
                    // Simpan status lama
                    try
                    {
                        oCompany = Services.CompanyService.GetCompany();
                        
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                        SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");
                        string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                        int docEntry;
                        if (int.TryParse(docEntryStr, out docEntry))
                        {
                            var oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            oRS.DoQuery($"SELECT Status FROM OWOR WHERE DocEntry = {docEntry}");
                            _statusLama = oRS.Fields.Item("Status").Value.ToString();
                        }
                    }
                    catch (Exception ex)
                    {
                        Application.SBO_Application.StatusBar.SetText("Gagal ambil status lama: " + ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ActionSuccess && pVal.ItemUID == "1")
                {
                    // After update sukses, ambil status baru dan bandingkan

                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    oCompany = Services.CompanyService.GetCompany();
                    oCompany.StartTransaction();
                    SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");

                    int docEntry = 0;
                    try
                    {
                        string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                        if (int.TryParse(docEntryStr, out docEntry))
                        {
                            var oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            oRS.DoQuery($"SELECT Status, ISNULL(U_T2_PRODTYPE,'') AS ProdType  FROM OWOR WHERE DocEntry = {docEntry}");
                            string statusBaru = oRS.Fields.Item("Status").Value.ToString();
                            string prodType = oRS.Fields.Item("ProdType").Value.ToString();

                            // Cek perubahan
                            if (_statusLama != statusBaru && statusBaru == "R" && prodType == "FG")
                            {
                                // Tambahkan proses kamu di sini
                                progress = Application.SBO_Application.StatusBar.CreateProgressBar("Generating sub-orders...", 100, false);

                                // Generate suborder
                                progress.Value = 0;
                                progress.Text = "Generating WIP Production Orders...";
                                var listDoc = ProductionOrderSapService.GenerateSubOrder(oCompany, docEntry);
                                foreach (var item in listDoc)
                                {
                                    int wipEntry = item.Key;
                                    ProductionOrderSapService.LinkWipToFG(oCompany, docEntry, wipEntry);
                                }

                                if (listDoc != null && listDoc.Any())
                                {
                                    string remarks = "WIP Production Orders: " + string.Join(" | ", listDoc.Values);
                                    ProductionOrderSapService.UpdateRemarks(oCompany, docEntry, remarks);
                                    
                                }
                                
                                progress.Value = 100; // Ensure it reaches the end
                                progress.Text = "Done.";
                                progress.Stop();
                                
                                RefreshFormProdOrder(oForm, docEntry);
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
                        if (progress != null) progress.Stop();
                        Application.SBO_Application.StatusBar.SetText("Failed to update: " + ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }

            }
        }

        private void RefreshFormProdOrder(SAPbouiCOM.Form oForm, int docEntry)
        {
            try
            {
                string sDocEntry = docEntry.ToString();

                // Simpan DocEntry ke variabel lokal
                oForm.Close(); // Tutup form dulu

                // Timer delay 1 detik sebelum open form lagi
                System.Timers.Timer reopenTimer = new System.Timers.Timer(1000);
                reopenTimer.AutoReset = false;
                reopenTimer.Elapsed += (sender, e) =>
                {
                    try
                    {
                        Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_ProductionOrder, "", sDocEntry);
                        Application.SBO_Application.StatusBar.SetText("Successfully generated WIP Production Orders.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    catch (Exception exOpen)
                    {
                        Application.SBO_Application.MessageBox("Failed to reopen Production Order form: " + exOpen.Message);
                    }
                    finally
                    {
                        reopenTimer.Dispose();
                    }
                };
                reopenTimer.Start();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Error during refresh: " + ex.Message);
            }
        }


    }
}
