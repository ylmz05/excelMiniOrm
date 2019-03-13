using MiniOrm.Attributes;

namespace MiniOrm.Models
{
    [SpreadSheet(Name = "Bayi Koordinatları")]
    public class VendorModel
    {
        [ExcelColumn(Name = "Bayi Adı")]
        public string VendorName { get; set; }

        [ExcelColumn(Name = "Latitude")]
        public double Latitude { get; set; }

        [ExcelColumn(Name = "Longitude")]
        public double Longitude { get; set; }
    }
}
