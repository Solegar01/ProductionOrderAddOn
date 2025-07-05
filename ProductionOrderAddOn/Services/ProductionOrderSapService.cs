using ProductionOrderAddOn.Models;
using SAPbobsCOM;
using System;
using System.Collections.Generic;

namespace ProductionOrderAddOn.Services
{
    internal static class ProductionOrderSapService
    {
        /// <summary>
        /// Menyimpan list ProductionOrderModel sebagai dokumen Production Order di SAP B1.
        /// </summary>
        /// <returns>Jumlah PO berhasil.</returns>
        public static int CreateProductionOrders(IEnumerable<ProductionOrderModel> models)
        {
            Company oCompany = CompanyService.GetCompany();   // singleton koneksi
            int success = 0;

            // Optional: jalankan dalam 1 transaksi agar atomic
            bool inTran = false;
            try
            {
                if (!oCompany.InTransaction)
                {
                    oCompany.StartTransaction();
                    inTran = true;
                }

                foreach (var m in models)
                {
                    ProductionOrders po = (ProductionOrders)
                        oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);

                    po.ItemNo = m.ProdNo;
                    po.ProductionOrderType = BoProductionOrderTypeEnum.bopotStandard;
                    po.ProductionOrderStatus = BoProductionOrderStatusEnum.boposPlanned;
                    po.PlannedQuantity = m.Qty;
                    po.PostingDate = DateTime.Today;
                    po.StartDate = DateTime.Today;
                    po.DueDate = m.OrderDate;

                    // contoh: isi User Fields jika perlu
                    // po.UserFields.Fields.Item("U_Desc").Value = m.ProdDesc;

                    int rc = po.Add();
                    if (rc != 0)
                    {
                        oCompany.GetLastError(out int errCode, out string errMsg);
                        throw new Exception($"Add() failed ({errCode}) {errMsg}");
                    }

                    int docEntry = int.Parse(oCompany.GetNewObjectKey());

                    UpdatePoStatus(docEntry, BoProductionOrderStatusEnum.boposReleased);

                    success++;
                }

                if (inTran) oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                return success;
            }
            catch
            {
                if (inTran && oCompany.InTransaction)
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }
        }

        public static void UpdatePoStatus(int docEntry,
                                          BoProductionOrderStatusEnum target)
        {
            Company oCompany = CompanyService.GetCompany();

            // Ambil PO
            var po = (ProductionOrders)oCompany.GetBusinessObject(
                         BoObjectTypes.oProductionOrders);

            if (!po.GetByKey(docEntry))
                throw new InvalidOperationException($"PO DocEntry {docEntry} tidak ditemukan.");

            // Jika sudah di status target, abaikan
            if (po.ProductionOrderStatus == target) return;

            po.ProductionOrderStatus = target;

            if (po.Update() != 0)
            {
                oCompany.GetLastError(out int errCode, out string errMsg);
                throw new InvalidOperationException(
                    $"Gagal update status PO ({errCode}) {errMsg}");
            }
        }
    }
}
