using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelConvertToolsXiongSetup
{
    public class ExcelImport
    {
        public static void Import(string path)
        {
            DataTable dataTable = new DataTable();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                string id = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().First<Sheet>().Id.Value;
                Row[] array = ((WorksheetPart)document.WorkbookPart.GetPartById(id)).Worksheet.GetFirstChild<SheetData>().Descendants<Row>().ToArray<Row>();
                foreach (Cell cell in (OpenXmlElement)((IEnumerable<Row>)array).ElementAt<Row>(0))
                    dataTable.Columns.Add(ExcelImport.GetCellValue(document, cell));
                for (int index1 = 1; index1 < ((IEnumerable<Row>)array).Count<Row>(); ++index1)
                {
                    DataRow row = dataTable.NewRow();
                    for (int index2 = 0; index2 < array[index1].Descendants<Cell>().Count<Cell>(); ++index2)
                        row[index2] = (object)ExcelImport.GetCellValue(document, array[index1].Descendants<Cell>().ElementAt<Cell>(index2));
                    dataTable.Rows.Add(row);
                }
            }
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart sharedStringTablePart = document.WorkbookPart.SharedStringTablePart;
            string s = cell?.CellValue?.InnerXml ?? "";
            if (cell.DataType != null && (cell.DataType.Value == CellValues.SharedString || cell.DataType.Value == CellValues.String || cell.DataType.Value == CellValues.Number))
                return sharedStringTablePart.SharedStringTable.ChildElements[int.Parse(s)].InnerText;
            return s;
        }
    }


    public class ExcelOpenXml
    {
        public static string GetNewExcelFileName(string suffix = ".xlsx")
        {
            return DateTime.Now.ToString("yyMMdd-HHmmss") + suffix;
        }

        public static void Create(string filename, DataSet ds)
        {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            Workbook workbook = new Workbook();
            Sheets sheets = new Sheets();
            for (int index1 = 0; index1 < ds.Tables.Count; ++index1)
            {
                DataTable table = ds.Tables[index1];
                string tableName = table.TableName;
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                Worksheet worksheet = new Worksheet();
                SheetData sheetData = new SheetData();
                Sheet sheet = new Sheet()
                {
                    Id = (StringValue)spreadsheetDocument.WorkbookPart.GetIdOfPart((OpenXmlPart)worksheetPart),
                    SheetId = UInt32Value.FromUInt32((uint)(index1 + 1)),
                    Name = (StringValue)tableName
                };
                sheets.Append((OpenXmlElement)sheet);
                uint num1 = 1;
                Row row1 = new Row();
                int num2 = (int)num1;
                uint num3 = (uint)(num2 + 1);
                row1.RowIndex = UInt32Value.FromUInt32((uint)num2);
                Row row2 = row1;
                sheetData.Append((OpenXmlElement)row2);
                for (int index2 = 0; index2 < table.Columns.Count; ++index2)
                {
                    Cell cell = new Cell();
                    cell.CellValue = new CellValue(table.Columns[index2].ColumnName);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    row2.Append((OpenXmlElement)cell);
                }
                for (int index2 = 0; index2 < table.Rows.Count; ++index2)
                {
                    Row row3 = new Row()
                    {
                        RowIndex = UInt32Value.FromUInt32(num3++)
                    };
                    sheetData.Append((OpenXmlElement)row3);
                    for (int index3 = 0; index3 < table.Columns.Count; ++index3)
                    {
                        Cell cell = new Cell();
                        object obj = table.Rows[index2][index3];
                        cell.CellValue = new CellValue(obj.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        row3.Append((OpenXmlElement)cell);
                    }
                }
                worksheet.Append((OpenXmlElement)sheetData);
                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();
            }
            workbook.Append((OpenXmlElement)sheets);
            workbookPart.Workbook = workbook;
            workbookPart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        public static DataTable GetSheet(string filename, string sheetName)
        {
            DataTable dt = new DataTable();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filename, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault<Sheet>();
                if (sheet == null)
                    throw new ArgumentException("未能找到" + sheetName + " sheet 页");
                SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault<SharedStringTablePart>();
                SharedStringTable sharedStringTable = (SharedStringTable)null;
                if (sharedStringTablePart != null)
                    sharedStringTable = sharedStringTablePart.SharedStringTable;
                Func<Row, int> func = (Func<Row, int>)(r =>
                {
                    foreach (Cell element in r.Elements<Cell>())
                        dt.Columns.Add(ExcelOpenXml.GetCellVal(element, sharedStringTable));
                    return dt.Columns.Count;
                });
                Action<Row> action = (Action<Row>)(r =>
                {
                    DataRow row = dt.NewRow();
                    int num = 0;
                    int count = dt.Columns.Count;
                    foreach (Cell element in r.Elements<Cell>())
                    {
                        if (num < count)
                            row[num++] = (object)ExcelOpenXml.GetCellVal(element, sharedStringTable);
                        else
                            break;
                    }
                    dt.Rows.Add(row);
                });
                foreach (Row element in (workbookPart.GetPartById((string)sheet.Id) as WorksheetPart).Worksheet.Elements<SheetData>().First<SheetData>().Elements<Row>())
                {
                    if ((uint)element.RowIndex == 1U)
                    {
                        int num = func(element);
                    }
                    else
                        action(element);
                }
            }
            return dt;
        }

        private static string GetCellVal(Cell cell, SharedStringTable sharedStringTable)
        {
            string s = cell.InnerText;
            if (cell.DataType != null)
            {
                if (cell.DataType.Value == CellValues.SharedString)
                {
                    if (sharedStringTable != null)
                        s = sharedStringTable.ElementAt<OpenXmlElement>(int.Parse(s)).InnerText;
                }
                else
                    s = cell?.InnerText ?? "";
            }
            return s;
        }
    }
}
