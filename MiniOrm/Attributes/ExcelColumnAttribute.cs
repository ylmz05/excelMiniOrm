using System;

namespace MiniOrm.Attributes
{
    public class ExcelColumnAttribute : Attribute
    {
        public string Name { get; set; }
    }
}
