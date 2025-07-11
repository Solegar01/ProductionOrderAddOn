using ProductionOrderAddOn.Helpers;
using ProductionOrderAddOn.Models;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace ProductionOrderAddOn.Services
{
    internal static class ProductionOrderSapService
    {
        private static readonly Company oCompany = CompanyService.GetCompany();

        public static List<int> CreateProductionOrdersRecursive(string fileName, IEnumerable<ProductionOrderModel> models, HashSet<ProductionKey> visitedKeys = null)
        {
            if (models == null) throw new ArgumentNullException(nameof(models));
            if (visitedKeys == null) visitedKeys = new HashSet<ProductionKey>();

            var allDocEntries = new List<int>();

            // 🔍 Filter hanya model dengan kombinasi ProdNo + OrderDate yang belum diproses
            var batchModels = models
                .Where(m => visitedKeys.Add(new ProductionKey(m.ProdNo, m.OrderDate, m.RefProdEntry)))
                .ToList();

            if (batchModels.Count == 0) return allDocEntries;
            
            bool startedTran = false;

            try
            {
                if (!oCompany.InTransaction)
                {
                    oCompany.StartTransaction();
                    startedTran = true;
                }

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
                        po.Remarks = $"Imported from file {fileName}";

                        if (!string.IsNullOrEmpty(m.RefProdEntry))
                            po.UserFields.Fields.Item("U_T2_Ref_Production").Value = m.RefProdEntry;
                        if (!string.IsNullOrEmpty(m.RefProdEntry))
                            po.UserFields.Fields.Item("U_T2_Ref_Prod_DocNum").Value = m.RefProdNum;
                        if (!string.IsNullOrEmpty(m.RefProdEntry))
                            po.UserFields.Fields.Item("U_T2_Is_Import").Value = "Y";
                        
                        int rc = po.Add();
                        if (rc != 0)
                        {
                            oCompany.GetLastError(out int errCode, out string errMsg);
                            throw new Exception($"Gagal membuat PO {m.ProdNo} ({errCode}): {errMsg}");
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

                if (startedTran) oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                var wipModels = GetProductionOrders(allDocEntries).ToList();

                if (wipModels.Count > 0)
                {
                    var subDocEntries = CreateProductionOrdersRecursive(fileName,wipModels, visitedKeys);
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

            string inClause = string.Join(", ", docEntries);

            string sql = $@"
                        SELECT
                            t0.DocEntry      AS RefProdEntry,
                            t0.DocNum      AS RefProdNum,
                            t2.Code        AS ProdNo,
                            t2.ItemName    AS ProdDesc,
                            t3.PlannedQty  AS Qty,
                            CAST(t0.PostDate AS DATE) AS OrderDate
                        FROM OWOR  t0
                        INNER JOIN OITT t1 ON t0.ItemCode = t1.Code
                        INNER JOIN ITT1 t2 ON t1.Code     = t2.Father
                        INNER JOIN WOR1 t3 ON t3.DocEntry = t0.DocEntry
                                            AND t3.ItemCode = t2.Code
                        WHERE t0.DocEntry IN ({inClause})
                            AND ISNULL(t2.U_T2_ITEM_GROUP, '') = 'WIP'
                        ORDER BY t0.PostDate DESC, t2.Code;";

            try
            {
                var result = new List<ProductionOrderModel>();
                var data = SapQueryHelper.ExecuteQuery(sql, oCompany);

                foreach (var row in data)
                {
                    result.Add(new ProductionOrderModel
                    {
                        RefProdEntry = row["RefProdEntry"].ToString(),
                        RefProdNum = row["RefProdNum"].ToString(),
                        ProdNo = row["ProdNo"].ToString(),
                        ProdDesc = row["ProdDesc"].ToString(),
                        Qty = Convert.ToDouble(row["Qty"]),
                        OrderDate = Convert.ToDateTime(row["OrderDate"]),
                        ProdType = ProductionType.WIP,
                    });
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception("Error while retrieving WIP production orders: " + ex.Message, ex);
            }
        }

        public static bool IsProdOrderExists(ProductionOrderModel model)
        {
            if (model == null)
                throw new ArgumentNullException(nameof(model));

            // Sanitize ProdNo dan format tanggal dengan benar
            string prodNo = model.ProdNo.Replace("'", "''"); // untuk hindari SQL injection
            string dateStr = model.OrderDate.ToString("yyyy-MM-dd");

            string sql = $@"
                SELECT TOP 1 T0.DocEntry
                FROM OWOR T0
                WHERE T0.ItemCode = '{prodNo}'
                AND CAST(T0.PostDate AS DATE) = '{dateStr}'";

            try
            {
                var result = SapQueryHelper.ExecuteQuery(sql, oCompany);
                return result.Count > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error while validating production orders: " + ex.Message, ex);
            }
        }
    }
}
