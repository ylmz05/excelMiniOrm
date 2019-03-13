namespace MiniOrm.DbContext
{
    class ExcelContext : Context
    {
        private const string PATH_SIPARIS_BAYI = "siparis ve bayi koordinatları.xlsx";

        public ExcelContext(string path = PATH_SIPARIS_BAYI) : base(path)
        {
        }
    }
}
