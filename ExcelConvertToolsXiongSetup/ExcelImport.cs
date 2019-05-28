using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelConvertToolsXiongSetup
{
    //    使用：

    //    object[] objTitle = new object[] { "SIM", "ICCID" };
    //string descFile = "d:\\test.xlsx";
    //using (var oxExt = new ExcelUtil.ExcelExport(descFile, objTitle.Length, "清单"))
    //{
    //oxExt.Write(objTitle);
    //}


    public class ExcelOpenXml
    {
        public static string GetNewExcelFileName(string suffix = ".xlsx")
        {
            return DateTime.Now.ToString("yyMMdd-HHmmss") + suffix;
        }

        public static void Create(string filename, DataSet ds)
        {
            Regex regex = new Regex(@"^(-?\d+)(\.\d+)?$");

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            Workbook workbook = new Workbook();
            Sheets sheets = new Sheets();

            #region 创建多个 sheet 页

            //创建多个sheet
            for (int s = 0; s < ds.Tables.Count; s++)
            {
                DataTable dt = ds.Tables[s];
                var tname = dt.TableName;

                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                Worksheet worksheet = new Worksheet();
                SheetData sheetData = new SheetData();

                //创建 sheet 页
                Sheet sheet = new Sheet()
                {
                    //页面关联的 WorksheetPart
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = UInt32Value.FromUInt32((uint)s + 1),
                    Name = tname
                };
                sheets.Append(sheet);

                #region 创建sheet 行
                Row row;
                uint rowIndex = 1;
                //添加表头
                row = new Row()
                {
                    RowIndex = UInt32Value.FromUInt32(rowIndex++)
                };
                sheetData.Append(row);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    Cell newCell = new Cell();
                    newCell.CellValue = new CellValue(dt.Columns[i].ColumnName);
                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                    row.Append(newCell);
                }
                //添加内容
                object val = null;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    row = new Row()
                    {
                        RowIndex = UInt32Value.FromUInt32(rowIndex++)
                    };
                    sheetData.Append(row);

                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        Cell newCell = new Cell();
                        val = dt.Rows[i][j];
                        var str = (val?.ToString() ?? "").Trim();
                        if (str.Length == 0 || regex.IsMatch(str))
                        {
                            newCell.CellValue = new CellValue(str);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else
                        {
                            newCell.CellValue = new CellValue(val.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        row.Append(newCell);
                    }

                }
                #endregion

                worksheet.Append(sheetData);
                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();
            }
            #endregion

            workbook.Append(sheets);
            workbookpart.Workbook = workbook;

            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        public static DataTable GetSheet(string filename, string sheetName, int jumpRowIndex = 1)
        {
            DataTable dt = new DataTable();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filename, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                //通过sheet名查找 sheet页
                Sheet sheet = wbPart
                    .Workbook
                    .Descendants<Sheet>()
                    .FirstOrDefault(s => s.Name == sheetName);

                if (sheet == null)
                {
                    throw new ArgumentException("未能找到" + sheetName + " sheet 页");
                }

                //获取Excel中共享表
                SharedStringTablePart sharedStringTablePart = wbPart
                    .GetPartsOfType<SharedStringTablePart>()
                    .FirstOrDefault();
                SharedStringTable sharedStringTable = null;
                if (sharedStringTablePart != null)
                    sharedStringTable = sharedStringTablePart.SharedStringTable;
                #region 构建datatable

                //添加talbe列,返回列数
                Func<Row, int> addTabColumn = (r) =>
                {
                    //遍历单元格
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        dt.Columns.Add(GetCellVal(c, sharedStringTable));
                    }
                    return dt.Columns.Count;
                };
                //添加行
                Action<Row> addTabRow = (r) =>
                {
                    DataRow dr = dt.NewRow();
                    int colIndex = 0;
                    int colCount = dt.Columns.Count;
                    //遍历单元格
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        if (colIndex >= colCount)
                            break;
                        dr[colIndex++] = GetCellVal(c, sharedStringTable);
                    }
                    dt.Rows.Add(dr);
                };
                #endregion


                //通过 sheet.id 查找 WorksheetPart 
                WorksheetPart worksheetPart
                    = wbPart.GetPartById(sheet.Id) as WorksheetPart;
                //查找 sheetdata
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                //遍历行
                foreach (Row r in sheetData.Elements<Row>())
                {
                    //构建table列
                    if (r.RowIndex == jumpRowIndex)
                    {
                        addTabColumn(r);
                        continue;
                    }
                    //构建table行
                    addTabRow(r);
                }

            }
            return dt;
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="sharedStringTable"></param>
        /// <returns></returns>
        static string GetCellVal(Cell cell, SharedStringTable sharedStringTable)
        {
            var val = cell.InnerText;

            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    //从共享表中获取值
                    case CellValues.SharedString:
                        if (sharedStringTable != null)
                            val = sharedStringTable
                                .ElementAt(int.Parse(val))
                                .InnerText;
                        break;
                    default:
                        val = string.Empty;
                        break;
                }

            }
            return val;
        }

    }


    public class ExcelImport
    {
        public static void Import(string path)
        {
            DataTable dt = new DataTable();

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                //string relationshipId = sheets.First().Id.Value = sheets.First(x => x.Name == "TestSheet").Id.Value;
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                Row[] rows = sheetData.Descendants<Row>().ToArray();

                // 设置表头DataTable
                foreach (Cell cell in rows.ElementAt(0))
                {
                    dt.Columns.Add((string)GetCellValue(spreadSheetDocument, cell));
                }

                // 添加内容
                for (int rowIndex = 1; rowIndex < rows.Count(); rowIndex++)
                {
                    DataRow tempRow = dt.NewRow();

                    for (int i = 0; i < rows[rowIndex].Descendants<Cell>().Count(); i++)
                    {
                        tempRow[i] = GetCellValue(spreadSheetDocument, rows[rowIndex].Descendants<Cell>().ElementAt(i));
                    }
                    dt.Rows.Add(tempRow);
                }
            }
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell?.CellValue?.InnerXml ?? "";

            if (cell.DataType != null && (cell.DataType.Value == CellValues.SharedString || cell.DataType.Value == CellValues.String || cell.DataType.Value == CellValues.Number))
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else //浮点数和日期对应的cell.DataType都为NULL
            {
                // DateTime.FromOADate((double.Parse(value)); 如果确定是日期就可以直接用过该方法转换为日期对象，可是无法确定DataType==NULL的时候这个CELL 数据到底是浮点型还是日期.(日期被自动转换为浮点
                return value;
            }
        }
    }
}
