using MiniOrm.DbContext;
using MiniOrm.Extension;
using MiniOrm.Models;
using System;
using System.Collections.Generic;

namespace MiniOrm
{
    class Program
    {
        static void Main(string[] args)
        {
            Context context = new ExcelContext();

            IEnumerable<VendorModel> vendors = context.GetList<VendorModel>();
            IEnumerable<OrderModel> orders = context.GetList<OrderModel>();

            foreach (var item in vendors)
                Console.WriteLine($"VendorName: {item.VendorName}, Longitude: {item.Longitude}, Latitude: {item.Latitude}");

            foreach (var item in orders)
                Console.WriteLine($"OrderNo: {item.OrderNo}, Longitude: {item.Longitude}, Latitude: {item.Latitude}");

        }
    }
}
