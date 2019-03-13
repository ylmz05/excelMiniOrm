using MiniOrm.DbContext;
using MiniOrm.Helpers;
using System.Collections.Generic;

namespace MiniOrm.Extension
{
    public static class Extension
    {
        public static IEnumerable<T> GetList<T>(this Context context)
        {
            return ExcelHelper.GetDataFromExcellSheet<T>(context.PATH);
        }
    }
}
