using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MiniOrm.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace MiniOrm.Helpers
{
    public class ExcelHelper
    {
        public static IEnumerable<T> GetDataFromExcellSheet<T>(string fileName)
        {
            IList<T> modelListInstance = (IList<T>)Activator.CreateInstance(typeof(List<T>));
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                string spreadSheetName = typeof(T).GetCustomAttribute<SpreadSheetAttribute>().Name;

                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.Where(x => x.Name == spreadSheetName).FirstOrDefault().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                string[] columns = new string[rows.ElementAt(0).Count()];

                for (int i = 0; i < columns.Length; i++)
                    columns[i] = GetCellValue(spreadSheetDocument, rows.ElementAt(0).ElementAt(i) as Cell);

                for (int rowIndex = 1; rowIndex < rows.Count(); rowIndex++)
                {
                    T modelInstance = (T)Activator.CreateInstance(typeof(T));
                    PropertyInfo[] properties = modelInstance.GetType().GetProperties();

                    for (int i = 0; i < rows.ElementAt(rowIndex).Descendants<Cell>().Count(); i++)
                    {
                        PropertyInfo propertyInfo = properties.Where(x => x.GetCustomAttribute<ExcelColumnAttribute>().Name == columns[i]).FirstOrDefault();
                        propertyInfo.SetValue(modelInstance, Convert.ChangeType(GetCellValue(spreadSheetDocument, rows.ElementAt(rowIndex).Descendants<Cell>().ElementAt(i)), propertyInfo.PropertyType));
                    }

                    modelListInstance.Add(modelInstance);
                }
            }

            return modelListInstance.AsEnumerable();
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }
    }
}
