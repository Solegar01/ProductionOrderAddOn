using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
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
        private List<ProductionOrderModel> listData;
        const string DT_NAME = "DT_IMPORT";
        SAPbouiCOM.DataTable dt;
        private SAPbouiCOM.Grid GridData;
        private SAPbouiCOM.StaticText LblPath;
        private SAPbouiCOM.EditText TxtPath;
        private SAPbouiCOM.Button BtnBrowse;
        private string fileName;

        public ImportForm()
        {
        }
        
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>

        public override void OnInitializeComponent()
        {
            this.TxtFrom = ((SAPbouiCOM.EditText)(this.GetItem("TxtFrom").Specific));
            this.LblDateFrom = ((SAPbouiCOM.StaticText)(this.GetItem("LblFrom").Specific));
            this.LblDateTo = ((SAPbouiCOM.StaticText)(this.GetItem("LblTo").Specific));
            this.TxtTo = ((SAPbouiCOM.EditText)(this.GetItem("TxtTo").Specific));
            this.BtnImport = ((SAPbouiCOM.Button)(this.GetItem("BtnImport").Specific));
            this.BtnImport.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnImport_ClickBefore);
            this.GridData = ((SAPbouiCOM.Grid)(this.GetItem("GridData").Specific));
            this.LblPath = ((SAPbouiCOM.StaticText)(this.GetItem("LblPath").Specific));
            this.TxtPath = ((SAPbouiCOM.EditText)(this.GetItem("TxtPath").Specific));
            this.BtnBrowse = ((SAPbouiCOM.Button)(this.GetItem("BtnBrowse").Specific));
            this.BtnBrowse.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnBrowse_ClickBefore);
            this.TxtFrom.ValidateAfter += new SAPbouiCOM._IEditTextEvents_ValidateAfterEventHandler(this.DateRangeValidation);
            this.TxtTo.ValidateAfter += new SAPbouiCOM._IEditTextEvents_ValidateAfterEventHandler(this.DateRangeValidation);
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
            try
            {
                this.ImportToSAP();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void BtnBrowse_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this.GetPathFile();
                this.ImportFromExcelProdOrder();
                this.SetDataGrid();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        //private void TxtDateFrom_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        //{
        //    //--------------------------------------------------------  
        //    // 1) PARSE string "yyyyMMdd" dari EditText → DateTime
        //    // ------------------------------------------------------------------
        //    if (!DateTime.TryParseExact(this.TxtFrom.Value, "yyyyMMdd",
        //                                CultureInfo.InvariantCulture,
        //                                DateTimeStyles.None, out DateTime tgl))
        //    {
        //        // format salah → biarkan SAP menampilkan pesan default
        //        return;
        //    }

        //    // ------------------------------------------------------------------  
        //    // 2) VALIDASI: minimal hari ini
        //    // ------------------------------------------------------------------
        //    if (tgl.Date < DateTime.Today)
        //    {
        //        Application.SBO_Application.StatusBar.SetText(
        //            "Invalid date from.",
        //            SAPbouiCOM.BoMessageTime.bmt_Short,
        //            SAPbouiCOM.BoStatusBarMessageType.smt_Error);

        //        // Kembalikan nilai ke hari ini supaya valid
        //        this.TxtFrom.Value = null;
        //        return;
        //    }

        //    DateRangeValidation(pVal);
        //}

        //private void TxtDateTo_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        //{
        //    //--------------------------------------------------------  
        //    // 1) PARSE string "yyyyMMdd" dari EditText → DateTime
        //    // ------------------------------------------------------------------
        //    if (!DateTime.TryParseExact(this.TxtTo.Value, "yyyyMMdd",
        //                                CultureInfo.InvariantCulture,
        //                                DateTimeStyles.None, out DateTime tgl))
        //    {
        //        // format salah → biarkan SAP menampilkan pesan default
        //        return;
        //    }

        //    // ------------------------------------------------------------------  
        //    // 2) VALIDASI: minimal hari ini
        //    // ------------------------------------------------------------------
        //    if (tgl.Date < DateTime.Today)
        //    {
        //        Application.SBO_Application.StatusBar.SetText(
        //            "Invalid date to.",
        //            SAPbouiCOM.BoMessageTime.bmt_Short,
        //            SAPbouiCOM.BoStatusBarMessageType.smt_Error);

        //        // Kembalikan nilai ke hari ini supaya valid
        //        this.TxtTo.Value = null;
        //        return;
        //    }

        //    DateRangeValidation(pVal);
        //}

        private void DateRangeValidation(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DateTime? from = ParseDate_yyyyMMdd(TxtFrom.Value);
            DateTime? to = ParseDate_yyyyMMdd(TxtTo.Value);

            // 5) Jika KEDUA‑DUANYA terisi dan From > To → error
            if (from.HasValue && to.HasValue && from.Value.Date > to.Value.Date)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Date From cannot be greater than Date To",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                // Kembalikan field yang barusan diubah supaya tetap valid
                if (pVal.ItemUID == "TxtFrom")
                    TxtFrom.Value = to.Value.ToString("yyyyMMdd");
                else                                            // TxtTo yang diubah
                    TxtTo.Value = from.Value.ToString("yyyyMMdd");

                return;
            }
        }

        private void GetPathFile()
        {
            try
            {
                Thread t = new Thread(() =>
                {
                    using (var dummyForm = new System.Windows.Forms.Form { TopMost = true, ShowInTaskbar = false, WindowState = System.Windows.Forms.FormWindowState.Minimized })
                    using (var openFileDialog = new System.Windows.Forms.OpenFileDialog
                    {
                        Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                        Title = "Select Excel file"
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
                this.TxtPath.Value = fileName ?? "";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ImportFromExcelProdOrder()
        {
            try
            {
                string fromStr = TxtFrom.Value;
                string toStr = TxtTo.Value;

                DateTime? fromDate = ParseDate_yyyyMMdd(fromStr);

                DateTime? toDate = ParseDate_yyyyMMdd(toStr);
                
                if (string.IsNullOrEmpty(fileName)) throw new Exception("Please select file to import");

                // Panggil import service
                this.listData = ExcelImportService.ImportProductionOrders(fileName, fromDate, toDate);
                

                if (!this.listData.Any())
                    throw new Exception("Data not found");
                
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        private void ImportToSAP()
        {
            try
            {
                if (string.IsNullOrEmpty(this.fileName))
                    throw new Exception("Please select a file to import.");

                int result = Application.SBO_Application.MessageBox(
                    "Are you sure you want to Import to SAP?",
                    1, "Yes", "No", "");

                if (result != 1)
                    return;

                if (listData == null || listData.Count == 0)
                    this.ImportFromExcelProdOrder();

                if (listData == null || listData.Count == 0)
                    throw new Exception("No data found in the selected file.");

                // 🌟 Satu baris rekursif → membuat semua PO hingga WIP selesai
                List<int> allDocEntries = ProductionOrderSapService.CreateProductionOrders(listData);

                if (allDocEntries == null || allDocEntries.Count == 0)
                    throw new Exception("No production orders were created in SAP.");

                Application.SBO_Application.StatusBar.SetText(
                    "Data successfully imported.",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                this.Reset();
            }
            catch (Exception ex)
            {

                throw new Exception("Error during import: " + ex.Message);

                // Optional: log ke file atau tampilkan pesan lebih detail jika diperlukan
            }
        }
        
        private void Reset()
        {
            if (listData != null) this.listData.Clear();
            if (dt != null)
            {
                this.dt.Clear();
                this.GridData.DataTable.Clear();
            }
            this.fileName = String.Empty;
            this.TxtPath.Value = String.Empty;
        }
        
        /// <summary>
        /// Membuat DataTable baru (jika belum ada) ATAU
        /// membersihkan baris lama, lalu mengisi ulang dengan 'data'.
        /// </summary>
        private void BuildOrResetDataTable(SAPbouiCOM.IForm oForm)
        {
            try
            {
                dt = oForm.DataSources.DataTables.Item(DT_NAME);   // ada? ambil
            }
            catch (System.Runtime.InteropServices.COMException)   // belum ada
            {
                dt = oForm.DataSources.DataTables.Add(DT_NAME);    // → buat
            }

            dt.Clear();

            dt.Columns.Add("No.", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("Description", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Quantity);
            dt.Columns.Add("Due Date", SAPbouiCOM.BoFieldsType.ft_Date);

            foreach (var x in listData)
            {
                int row = dt.Rows.Count;   // ambil indeks berikutnya
                dt.Rows.Add();             // tambahkan baris kosong
                dt.SetValue("No.", row, x.ProdNo);
                dt.SetValue("Description", row, x.ProdDesc);
                dt.SetValue("Due Date", row, x.OrderDate);
                dt.SetValue("Qty", row, x.Qty);
            }

            int white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            
            GridData.DataTable = dt;
            GridData.AutoResizeColumns();
            for (int i = 0; i < GridData.Columns.Count; i++)
            {
                var col = GridData.Columns.Item(i);

                // Untuk SAP B1 ≥ 9.2: properti ada di TitleObject
                col.TitleObject.Sortable = true;
                col.Editable = false;
                // Jika Anda memakai versi lama dan TitleObject.Sortable belum ada,
                // gunakan:  col.Sortable = true;
            }
            for (int i = 0; i < GridData.Rows.Count; i++)
            {
                GridData.CommonSetting.SetRowBackColor(i + 1, white);
            }
        }


        private void SetDataGrid()
        {
            var oForm = this.UIAPIRawForm;

            oForm.Freeze(true);          // >>> tahan repaint
            try
            {
                if (listData != null)
                {
                    BuildOrResetDataTable(oForm); // isi dt & bind ke GridData
                    GridData.AutoResizeColumns();

                    SetRowNumber();

                    // pastikan nomor tetap ketika user sort
                    if (!_sortHandlerAdded)
                    {
                        GridData.GridSortAfter += (s, e) => SetRowNumber();
                        _sortHandlerAdded = true;
                    }
                }
            }
            finally
            {
                oForm.Freeze(false);     // >>> lepaskan, UI refresh sekali saja
            }
        }

        bool _sortHandlerAdded = false;


        private void SetRowNumber()
        {
            var grid = this.GridData;
            grid.RowHeaders.TitleObject.Caption = "#";    // judul kolom
            grid.RowHeaders.Width = 30;                   // lebar (pixel) — sesuaikan
            
            int rowCount = grid.DataTable.Rows.Count;
            for (int i = 0; i < rowCount; i++)
            {
                // RowHeaders indeks‑nya sama dengan indeks baris DataTable
                grid.RowHeaders.SetText(i, (i + 1).ToString());
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