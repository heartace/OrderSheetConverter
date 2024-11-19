using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Studio.DreamRoom.OrderSheetConverter
{
    internal class SheetParser
    {
        internal static SheetData Parse(DataTable table)
        {
            const string SheetHeaderWeChatName  = "微信名称";
            const string SheetHeaderProductName = "商品名称";
            const string SheetHeaderSpec        = "规格";
            const string SheetHeaderQuantity    = "数量";
            const string SheetHeaderTotalPrice  = "商品总金额";

            var headers = table.Rows[0];
            var headerNames = new List<String>();

            for (int i = 0; i < headers.ItemArray.Length; i++)
            {
                headerNames.Add(headers[i].ToString());
            }

            var colWeChatName = headerNames.IndexOf(SheetHeaderWeChatName);
            var colProductName = headerNames.IndexOf(SheetHeaderProductName);
            var colSpec = headerNames.IndexOf(SheetHeaderSpec);
            var colQuantity = headerNames.IndexOf(SheetHeaderQuantity);
            var colTotalPrice = headerNames.IndexOf(SheetHeaderTotalPrice);

            var buyers = new List<String>();
            for (int i = 1; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                var buyer = row[colWeChatName].ToString();
                if (buyer != null && !buyers.Contains(buyer))
                {
                    buyers.Add(buyer);
                }
            }

            var products = new Dictionary<string, Dictionary<String, List<Order>>>();
            for (int i = 1; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                var rawProductName = row[colProductName].ToString();
                var spec = row[colSpec].ToString();

                if (rawProductName != null && spec != null)
                {
                    var suffix = $"({spec})";
                    if (!rawProductName.EndsWith(suffix))
                    {
                        Debug.WriteLine($"商品名称和规格不匹配 - 商品名称: #{rawProductName}, 规格: {spec}");
                    }

                    var index = rawProductName.IndexOf(suffix);
                    if (index > -1)
                    {
                        var productName = rawProductName.Substring(0, index);
                        if (!products.ContainsKey(productName))
                        {
                            products[productName] = new Dictionary<String, List<Order>>();
                        }

                        if (!products[productName].ContainsKey(spec))
                        {
                            products[productName][spec] = new List<Order>();
                        }

                        var buyer = row[colWeChatName].ToString();
                        var rawQuantity = row[colQuantity].ToString();
                        var rawTotalPrice = row[colTotalPrice].ToString();

                        if (buyer != null && rawQuantity != null && rawTotalPrice != null)
                        {
                            int.TryParse(rawQuantity.Trim(), out int quantity);

                            decimal.TryParse(rawTotalPrice.Trim(), out decimal totalPrice);

                            if (quantity > 0 && totalPrice > 0)
                            {
                                var existingOrder = products[productName][spec].Find(x => x.BuyerName == buyer);

                                if (existingOrder == null)
                                {
                                    var unitPrice = totalPrice / quantity;
                                    products[productName][spec].Add(new Order(buyer, unitPrice, quantity));
                                }
                                else
                                {
                                    Debug.WriteLine($"Existing order found: {existingOrder}");
                                    existingOrder.Quantity += quantity;
                                }
                            }
                        }
                    }
                }
            }

            Debug.WriteLine($"Buyers count: {buyers.Count} | Products count: {products.Count} | Raw orders count: {table.Rows.Count - 1} | Actual orders count: {products.Values.AsQueryable().Sum(x => x.Values.AsQueryable().Sum(y => y.Count))}"); 

            return new SheetData(table.Rows.Count - 1, buyers, products);
        }
    }

    internal struct SheetData
    {
        internal int OriginalOrdersCount { get; init; }

        internal List<String> Buyers { get; init; }

        internal Dictionary<string, Dictionary<String, List<Order>>> Products { get; init; }

        internal SheetData(int originalOrdersCount, List<String> buyers, Dictionary<string, Dictionary<String, List<Order>>> products)
        {
            OriginalOrdersCount = originalOrdersCount;
            Buyers = buyers;
            Products = products;
        }

        internal int ProductSpecsCount
        {
            get
            {
                return Products.Sum(p => p.Value.Count);
            }
        }

        internal int OrdersCount
        {
            get
            {
                return Products.Values.AsQueryable().Sum(x => x.Values.AsQueryable().Sum(y => y.Count));
            }
        }
    }

    internal class Order
    {
        internal String BuyerName { get; init; }

        internal decimal UnitPrice { get; init; }

        internal int Quantity { get; set; }

        public String DisplayName
        {
            get
            {
                return $"{BuyerName} ({Quantity})";
            }
        }

        internal Order(String buyerName, decimal unitPrice, int quantity)
        {
            BuyerName = buyerName;
            UnitPrice = unitPrice;
            Quantity = quantity;
        }

        override public string ToString()
        {
            return $"<Order> {{ BuyerName: {BuyerName}, UnitPrice: {UnitPrice}, Quantity: {Quantity} }}";
        }
    }
}
