using MiniOrm.Attributes;

namespace MiniOrm.Models
{
    [SpreadSheet(Name = "Siparişler")]
    public class OrderModel
    {
        [ExcelColumn(Name = "Sipariş Numarası")]
        public int OrderNo { get; set; }

        [ExcelColumn(Name = "Latitude")]
        public double Latitude { get; set; }

        [ExcelColumn(Name = "Longitude")]
        public double Longitude { get; set; }
    }
}
