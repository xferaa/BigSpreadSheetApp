using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace BigSpreadSheetApp
{
    public class BigSpreadSheetParser
    {
        private readonly string filePath;

        public BigSpreadSheetParser(string filePath)
        {
            this.filePath = filePath;
        }

        public List<List<object>> ParseSpreadSheet()
        {
            List<List<object>> result = new List<List<object>>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                var sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ToList();

                while (reader.Read())
                {
                    if(reader.ElementType == typeof(Row))
                    {
                        reader.ReadFirstChild();

                        List<object> row = new List<object>();

                        do
                        {
                            if (reader.ElementType == typeof(Cell))
                            {
                                Cell c = (Cell)reader.LoadCurrentElement();

                                if (c == null || c.DataType == null || !c.DataType.HasValue)
                                    continue;

                                object value;

                                switch(c.DataType.Value)
                                {
                                    case CellValues.Boolean:
                                        value = bool.Parse(c.CellValue.InnerText);
                                        break;
                                    case CellValues.Date:
                                        value = DateTime.Parse(c.CellValue.InnerText);
                                        break;
                                    case CellValues.Number:
                                        value = double.Parse(c.CellValue.InnerText);
                                        break;
                                    case CellValues.InlineString:
                                    case CellValues.String:
                                        value = c.CellValue.InnerText;
                                        break;
                                    case CellValues.SharedString:
                                        value = sharedStringTable.ElementAt(int.Parse(c.CellValue.InnerText)).InnerText;
                                        break;
                                    default:
                                        continue;
                                }

                                if (value != null)
                                    row.Add(value);
                            }

                        } while (reader.ReadNextSibling());

                        if (row.Any())
                            result.Add(row);
                    }
                }
            }

            return result;
        }
    }
}
