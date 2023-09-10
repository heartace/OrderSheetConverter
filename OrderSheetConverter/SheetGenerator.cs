using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace Studio.DreamRoom.OrderSheetConverter
{
    internal class SheetGenerator
    {
        internal static void Export(SheetData sheetData, string fileName, string sourceFileName)
        {
            using (var p = new ExcelPackage())
            {
                var buyers = sheetData.Buyers;
                var products = sheetData.Products;

                var ws = p.Workbook.Worksheets.Add("Main");

                // headers

                for (int i = 0; i < buyers.Count; i++)
                {
                    var buyer = buyers[i];

                    var baseCol = 3 + i * 2 + 1;

                    var beginCell = $"{baseCol.ToSheetColumn()}1";
                    var endCell = $"{(baseCol + 1).ToSheetColumn()}1";
                    
                    ws.Cells[beginCell].Value = buyer;
                    ws.Cells[$"{beginCell}:{endCell}"].Merge = true;

                    ws.Cells[$"{baseCol.ToSheetColumn()}2"].Value = "数量";
                    ws.Cells[$"{(baseCol + 1).ToSheetColumn()}2"].Value = "价格";
                }

                ws.Cells["A2"].Value = "产品名";
                ws.Cells["B2"].Value = "材质";
                ws.Cells["C2"].Value = "单价";

                var colAfterBuyers = 3 + buyers.Count * 2 + 1;
                ws.Cells[$"{colAfterBuyers.ToSheetColumn()}2"].Value = "总数";
                ws.Cells[$"{(colAfterBuyers + 1).ToSheetColumn()}2"].Value = "循环/卷";
                ws.Cells[$"{(colAfterBuyers + 2).ToSheetColumn()}2"].Value = "卷数";

                // products

                var lastRow = 3;

                foreach (var entry in products)
                {
                    var product = entry.Key;
                    var specsCount = entry.Value.Count;
                    ws.Cells[$"A{lastRow}"].Value = product;
                    ws.Cells[$"A{lastRow}:A{lastRow + specsCount - 1}"].Merge = true;

                    var i = 0;
                    foreach (var specEntry in entry.Value)
                    {
                        var spec = specEntry.Key;
                        ws.Cells[$"B{lastRow + i}"].Value = spec;

                        var orders = specEntry.Value;
                        if (orders.Count > 0)
                        {
                            ws.Cells[$"C{lastRow + i}"].Value = orders[0].UnitPrice;

                            foreach (var order in orders)
                            {
                                var buyerIndex = buyers.IndexOf(order.BuyerName);
                                if (buyerIndex > -1)
                                {
                                    ws.Cells[$"{(3 + buyerIndex * 2 + 1).ToSheetColumn()}{lastRow + i}"].Value = order.Quantity;
                                }
                            }
                        }

                        for (var b = 0; b < buyers.Count; b++)
                        {
                            ws.Cells[$"{(3 + b * 2 + 2).ToSheetColumn()}{lastRow + i}"].Formula = $"C{lastRow + i}*{(3 + b * 2 + 1).ToSheetColumn()}{lastRow + i}";
                        }

                        i++;
                    }

                    lastRow += specsCount;
                }

                for (var i = 0; i < buyers.Count; i++)
                {
                    var col = (5 + i * 2).ToSheetColumn();

                    ws.Cells[$"{col}{lastRow}"].Formula = $"SUM({col}3:{col}{lastRow - 1})";
                }

                for (var i = 3; i < lastRow + 1; i++)
                {
                    var col = (4 + buyers.Count * 2).ToSheetColumn();

                    var lastBuyerCol = (3 + buyers.Count * 2).ToSheetColumn();

                    var mod = i == lastRow ? 1 : 2;

                    ws.Cells[$"{col}{i}"].Formula = $"SUMPRODUCT((MOD(COLUMN(D{i}:{lastBuyerCol}{i}),{mod})=0)*(D{i}:{lastBuyerCol}{i}))";

                    if (i < lastRow)
                    {
                        var cycleCol = (5 + buyers.Count * 2).ToSheetColumn();

                        var targetCol =(6 + buyers.Count * 2).ToSheetColumn();

                        ws.Cells[$"{targetCol}{i}"].Formula = $"=IF({cycleCol}{i},{col}{i}/{cycleCol}{i},\"\")";
                    }
                }

                // apply styles

                var green = Color.FromArgb(236, 241, 224);

                var blue = Color.FromArgb(222, 230, 240);

                // set the background of the first row to green
                using (var range = ws.Cells[$"A1:{(6 + buyers.Count * 2).ToSheetColumn()}1"])
                {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(green);
                }

                // set the background of the second row to blue
                using (var range = ws.Cells[$"A2:{(6 + buyers.Count * 2).ToSheetColumn()}2"])
                {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(blue);
                }

                // set the background of the each buyer's total price column to blue
                var totalSpecsCount = sheetData.ProductSpecsCount;

                for (var i = 0; i < buyers.Count; i++)
                {
                    var col = (5 + i * 2).ToSheetColumn();

                    using (var range = ws.Cells[$"{col}3:{col}{3 + totalSpecsCount - 1}"])
                    {
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(blue);
                    }
                }

                // center align the row & column headers
                ws.Rows[1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Columns[1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // set width for the first two columns
                ws.Columns[1].Width = 25;
                ws.Columns[2].Width = 25;

                // set height for all rows
                for (var i = 1; i < lastRow + 1; i++)
                {
                    ws.Rows[i].Height = 18;
                }

                ws.Workbook.Properties.SetCustomPropertyValue("Sheet Generator", "OrderSheetConverter Desktop");
                ws.Workbook.Properties.SetCustomPropertyValue("Sheet Generator Version", $"{Assembly.GetExecutingAssembly().GetName().Version}");
                ws.Workbook.Properties.SetCustomPropertyValue("Source File", sourceFileName);

                p.SaveAs(fileName);
            }
        }
    }
}
