using DocumentFormat.OpenXml.Packaging;

namespace MiniOrm.DbContext
{
    public abstract class Context
    {
        public readonly SpreadsheetDocument DataContext;

        public readonly string PATH;

        public Context(string path)
        {
            PATH = path;

            DataContext = Load();
        }

        public virtual SpreadsheetDocument Load()
        {
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(PATH, false))
            {
                return spreadSheetDocument;
            }
        }
    }
}
