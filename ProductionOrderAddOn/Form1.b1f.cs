using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml;
using ProductionOrderAddOn.Models;
using ProductionOrderAddOn.Services;
using SAPbouiCOM.Framework;

namespace ProductionOrderAddOn
{
    [FormAttribute("ProductionOrderAddOn.ImportFile", "Form1.b1f")]
    class ImportForm : UserFormBase
    {
        private SAPbouiCOM.EditText TxtFrom;
        private SAPbouiCOM.StaticText LblDateFrom;
        private SAPbouiCOM.StaticText LblDateTo;
        private SAPbouiCOM.EditText TxtTo;
        private SAPbouiCOM.Button BtnImport;

        public ImportForm()
        {
        }
        
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>

        public override void OnInitializeComponent()
        {
            this.TxtFrom = ((SAPbouiCOM.EditText)(this.GetItem("TxtFrom").Specific));
            this.LblDateFrom = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.LblDateTo = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.TxtTo = ((SAPbouiCOM.EditText)(this.GetItem("TxtTo").Specific));
            this.BtnImport = ((SAPbouiCOM.Button)(this.GetItem("BtnImport").Specific));
            this.BtnImport.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnImport_ClickBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }


        private void OnCustomInitialize()
        {

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            //try
            //{
            //    SAPbobsCOM.Company oCompany = CompanyService.GetCompany();
            //    if (oCompany.Connected)
            //    {
            //        Application.SBO_Application.MessageBox("Connection Succes");
            //    }
            //}
            //catch (Exception)
            //{
            //    Application.SBO_Application.MessageBox("Connection Fail");
            //}

        }

        private void BtnImport_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            ImportFromExcelProdOrder();
        }

        private void ImportFromExcelProdOrder()
        {
            try
            {
                string fromStr = TxtFrom.Value;
                string toStr = TxtTo.Value;

                DateTime? fromDate = ParseDate_yyyyMMdd(fromStr);

                DateTime? toDate = ParseDate_yyyyMMdd(toStr);
                
                // Open file dialog (STA thread)
                string fileName = null;
                Thread t = new Thread(() =>
                {
                    using (var dummyForm = new System.Windows.Forms.Form { TopMost = true, ShowInTaskbar = false, WindowState = System.Windows.Forms.FormWindowState.Minimized })
                    using (var openFileDialog = new System.Windows.Forms.OpenFileDialog
                    {
                        Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                        Title = "Pilih file Excel"
                    })
                    {
                        dummyForm.Show();
                        dummyForm.Hide();
                        if (openFileDialog.ShowDialog(dummyForm) == System.Windows.Forms.DialogResult.OK)
                        {
                            fileName = openFileDialog.FileName;
                        }
                    }
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();

                if (string.IsNullOrEmpty(fileName)) return;

                // Panggil import service
                var listData = ExcelImportService.ImportProductionOrders(fileName, fromDate, toDate);

                if (!listData.Any())
                    throw new Exception("Data gagal di-import atau tidak ada yang sesuai filter tanggal.");

                int res = ProductionOrderSapService.CreateProductionOrders(listData);

                Application.SBO_Application.StatusBar.SetText("Data berhasil di-import", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText(e.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        DateTime? ParseDate_yyyyMMdd(string raw)
        {
            return DateTime.TryParseExact(
                       raw?.Trim(),             // string sumber
                       "yyyyMMdd",              // format persis
                       CultureInfo.InvariantCulture,
                       DateTimeStyles.None,
                       out DateTime dt)
                   ? (DateTime?)dt
                   : null;
        }
    }
}