using ProductionOrderAddOn.Models;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using Dapper;
using System.Linq;
using System.Runtime.InteropServices;

namespace ProductionOrderAddOn.Services
{
    internal static class ProductionOrderSapService
    {
        private static readonly string _connStr =
        ConfigurationManager.ConnectionStrings["B1Connection"].ConnectionString;
        /// <summary>
        /// Menyimpan list ProductionOrderModel sebagai dokumen Production Order di SAP B1.
        /// </summary>
        /// <returns>Jumlah PO berhasil.</returns>
        //public static List<int> CreateProductionOrders(IEnumerable<ProductionOrderModel> models)
        //{
        //    Company oCompany = CompanyService.GetCompany();   // singleton koneksi
        //    List<int> listDocEntry = new List<int>();

        //    // Optional: jalankan dalam 1 transaksi agar atomic
        //    bool inTran = false;
        //    try
        //    {
        //        if (!oCompany.InTransaction)
        //        {
        //            oCompany.StartTransaction();
        //            inTran = true;
        //        }

        //        foreach (var m in models)
        //        {
        //            ProductionOrders po = (ProductionOrders)
        //                oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);

        //            po.ItemNo = m.ProdNo;
        //            po.ProductionOrderType = BoProductionOrderTypeEnum.bopotStandard;
        //            po.ProductionOrderStatus = BoProductionOrderStatusEnum.boposPlanned;
        //            po.PlannedQuantity = m.Qty;
        //            po.PostingDate = m.OrderDate;
        //            po.StartDate = m.OrderDate;
        //            po.DueDate = m.OrderDate;
        //            po.UserFields.Fields.Item("U_T2_PRODTYPE").Value = m.ProdType.ToString();

        //            // contoh: isi User Fields jika perlu
        //            // po.UserFields.Fields.Item("U_Desc").Value = m.ProdDesc;

        //            int rc = po.Add();
        //            if (rc != 0)
        //            {
        //                oCompany.GetLastError(out int errCode, out string errMsg);
        //                throw new Exception($"Add() failed ({errCode}) {errMsg}");
        //            }

        //            int docEntry = int.Parse(oCompany.GetNewObjectKey());

        //            UpdatePoStatus(docEntry, BoProductionOrderStatusEnum.boposReleased);

        //            listDocEntry.Add(docEntry);
        //        }

        //        if (inTran) oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
        //        return listDocEntry;
        //    }
        //    catch
        //    {
        //        if (inTran && oCompany.InTransaction)
        //            oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
        //        throw;
        //    }
        //}

        public static List<int> CreateProductionOrders(
    IEnumerable<ProductionOrderModel> models,
    HashSet<string> visited = null)
        {
            if (models == null) throw new ArgumentNullException(nameof(models));
            if (visited == null) visited = new HashSet<string>();

            var allDocEntries = new List<int>();

            // Filter model yang belum pernah diproses berdasarkan ProdNo
            var batchModels = models.Where(m => visited.Add(m.ProdNo)).ToList();
            if (batchModels.Count == 0) return allDocEntries;

            Company oCompany = CompanyService.GetCompany();
            bool startedTran = false;

            try
            {
                /* 1️⃣  Mulai transaksi SAP */
                if (!oCompany.InTransaction)
                {
                    oCompany.StartTransaction();
                    startedTran = true;
                }

                /* 2️⃣  Proses pembuatan PO */
                foreach (var m in batchModels)
                {
                    ProductionOrders po = null;
                    try
                    {
                        po = (ProductionOrders)oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);
                        po.ItemNo = m.ProdNo;
                        po.ProductionOrderType = BoProductionOrderTypeEnum.bopotStandard;
                        po.ProductionOrderStatus = BoProductionOrderStatusEnum.boposPlanned;
                        po.PlannedQuantity = m.Qty;
                        po.PostingDate = m.OrderDate.Date;
                        po.StartDate = m.OrderDate.Date;
                        po.DueDate = m.OrderDate.Date;
                        po.UserFields.Fields.Item("U_T2_PRODTYPE").Value = m.ProdType.ToString();
                        if (m.RefProd != null)
                        {
                            po.UserFields.Fields.Item("U_T2_Ref_Production").Value = m.RefProd;
                        }

                        int rc = po.Add();
                        if (rc != 0)
                        {
                            oCompany.GetLastError(out int errCode, out string errMsg);
                            throw new Exception($"Failed to add production order ({errCode}): {errMsg}");
                        }

                        int docEntry = int.Parse(oCompany.GetNewObjectKey());
                        allDocEntries.Add(docEntry);

                        UpdatePoStatus(docEntry, BoProductionOrderStatusEnum.boposReleased);
                    }
                    finally
                    {
                        if (po != null) Marshal.ReleaseComObject(po);
                    }
                }

                /* 3️⃣  Commit transaksi */
                if (startedTran) oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                /* 4️⃣  Ambil WIP dari hasil PO */
                var wipModels = GetProductionOrders(allDocEntries)
                                .ToList();

                /* 5️⃣  REKURSI untuk PO WIP */
                if (wipModels.Count > 0)
                {
                    var subDocEntries = CreateProductionOrders(wipModels, visited);
                    allDocEntries.AddRange(subDocEntries);
                }

                return allDocEntries;
            }
            catch
            {
                if (startedTran && oCompany.InTransaction)
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }
        }


        public static void UpdatePoStatus(int docEntry, BoProductionOrderStatusEnum target)
        {
            Company oCompany = CompanyService.GetCompany();

            // Ambil PO
            var po = (ProductionOrders)oCompany.GetBusinessObject(
                         BoObjectTypes.oProductionOrders);

            if (!po.GetByKey(docEntry))
                throw new InvalidOperationException($"Production Order DocEntry {docEntry} tidak ditemukan.");

            // Jika sudah di status target, abaikan
            if (po.ProductionOrderStatus == target) return;

            po.ProductionOrderStatus = target;

            if (po.Update() != 0)
            {
                oCompany.GetLastError(out int errCode, out string errMsg);
                throw new InvalidOperationException(
                    $"Failed to update status Production Order ({errCode}) {errMsg}");
            }
        }


        public static List<ProductionOrderModel> GetProductionOrders(IEnumerable<int> docEntries)
        {
            if (docEntries == null)
                throw new ArgumentNullException(nameof(docEntries));

            var ids = docEntries.Distinct().ToArray();
            if (ids.Length == 0)
                return new List<ProductionOrderModel>();

            const string sql = @"
                                SELECT
                                    t0.DocEntry      AS RefProd,
                                    t2.Code        AS ProdNo,
                                    t2.ItemName    AS ProdDesc,
                                    t3.PlannedQty  AS Qty,
                                    CAST(t0.PostDate AS DATE) AS OrderDate
                                FROM OWOR  t0
                                INNER JOIN OITT t1 ON t0.ItemCode = t1.Code
                                INNER JOIN ITT1 t2 ON t1.Code     = t2.Father
                                INNER JOIN WOR1 t3 ON t3.DocEntry = t0.DocEntry
                                                  AND t3.ItemCode = t2.Code
                                WHERE t0.DocEntry IN @DocEntryList
                                  AND ISNULL(t2.U_T2_ITEM_GROUP, '') = 'WIP'
                                ORDER BY t0.PostDate DESC, t2.Code;";

            try
            {
                using (var cn = new SqlConnection(_connStr))
                {
                    cn.Open();
                    var result = cn.Query<ProductionOrderModel>(sql, new { DocEntryList = ids }).ToList();

                    // Tandai semua sebagai WIP
                    foreach (var item in result)
                    {
                        item.ProdType = ProductionType.WIP;
                    }

                    return result;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error while retrieving WIP production orders: " + ex.Message, ex);
            }
        }


    }
}
