using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using ProductionOrderAddOn.Models;

namespace ProductionOrderAddOn.Services
{
    internal static class ExcelImportService
    {
        /// <summary>
        /// Import Production Order dari file Excel.
        /// Tiap baris (mulai baris 2) dibuat Production Order baru.
        /// </summary>
        /// <param name="filePath">Path penuh file .xlsx</param>
        /// <returns>Jumlah dokumen berhasil.</returns>
        public static List<ProductionOrderModel> ImportProductionOrders(
            string filePath,
            DateTime? fromDate,
            DateTime? toDate)
        {
            var results = new List<ProductionOrderModel>();

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    IXLWorksheet ws = workbook.Worksheet(1);

                    const int startColDate = 3;   // kolom C
                    const int endCol = 34;  // kolom AH

                    // 1️⃣  Tentukan rentang efektif
                    DateTime rangeStart = (fromDate ?? DateTime.MinValue).Date;
                    DateTime rangeEnd = (toDate ?? DateTime.MaxValue).Date;

                    // 2️⃣  Petakan kolom‑>tanggal di header (baris 3)
                    var colDateMap = new List<(int ColIndex, DateTime OrderDate)>();

                    for (int col = startColDate; col <= endCol; col++)
                    {
                        string headerText = ws.Cell(3, col).GetValue<string>();

                        if (DateTime.TryParse(headerText, out DateTime headerDate))
                        {
                            headerDate = headerDate.Date; // abaikan komponen waktu
                            if (headerDate >= rangeStart && headerDate <= rangeEnd)
                            {
                                colDateMap.Add((col, headerDate));
                            }
                        }
                    }

                    // 3️⃣  Loop baris data (mulai baris ke‑4)
                    foreach (IXLRow row in ws.RowsUsed().Skip(3))
                    {
                        foreach (var (colIndex, orderDate) in colDateMap)
                        {
                            IXLCell cell = row.Cell(colIndex);

                            if (!cell.IsEmpty() && !string.IsNullOrWhiteSpace(cell.GetString()) &&
                                double.TryParse(cell.GetValue<string>(), out double qty))
                            {
                                results.Add(new ProductionOrderModel
                                {
                                    ProdNo = row.Cell(1).GetValue<string>(),
                                    ProdDesc = row.Cell(2).GetValue<string>(),
                                    Qty = qty,
                                    OrderDate = orderDate
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Bungkus ulang supaya caller dapat pesan kontekstual
                throw new Exception($"Failed to proceed Excel file: {ex.Message}", ex);
            }

            return results;
        }

    }

}
