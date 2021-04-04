using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LectorExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"C:\temp\Prueba3Hojas.xlsx";
            var result = new DataTable();
            MoveExcelSheetsToCSV(fileName, false);


        }
        static private void MoveExcelSheetsToCSV(string fname, bool firstRowIsHeader)
        {         
            var Headers = new List<string>();
            var selectedList = new List<string>();
            selectedList.Add("HojaNumero1");
            selectedList.Add("TerceraHoja");
            DataTable dt = new DataTable();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fname, false))
            {
                //Read the first Sheets 
                IEnumerable<Sheet> sheets = doc.WorkbookPart.Workbook.Sheets.Descendants<Sheet>(); //.Where(sh => sh.Name.ToString().StartsWith("H"));
                var selectedSheets = new List<Sheet>();
                foreach (var item in sheets)
                {
                    if (IsSelected(item, selectedList))
                        selectedSheets.Add(item);
                }
                //Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                foreach (var sheet in selectedSheets)
                {
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                    int counter = 0;
                    string docPath = @"c:\Temp\";
                    string fileName = sheet.Name + ".csv";
                    Console.WriteLine(sheet.Name);
                    using (StreamWriter outputFile = new StreamWriter(Path.Combine(docPath, fileName)))
                    {
                        foreach (Row row in rows)
                        {
                            StringBuilder line = new StringBuilder();
                            counter = counter + 1;
                            //Read the first row as header
                            if (counter == 1)
                            {
                                var j = 1;
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    var colunmName = firstRowIsHeader ? GetCellValue(doc, cell) : "Field" + j++;
                                    Console.WriteLine(colunmName);
                                    Headers.Add(colunmName);
                                    //dt.Columns.Add(colunmName);
                                }
                            }
                            else
                            {

                                //dt.Rows.Add();
                                //int i = 0;
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    //dt.Rows[dt.Rows.Count - 1][i] = GetCellValue(doc, cell);
                                    line.Append(GetCellValue(doc, cell));
                                    line.Append(";");
                                    //i++;
                                }
                                if (!string.IsNullOrEmpty(line.ToString()))
                                    line.Remove(line.Length - 1, 1);
                            }
                            Console.WriteLine(line.ToString());
                            outputFile.WriteLine(line);
                        }
                    }
                }
                //return dt;
            }
            
        }

        public static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }
            return value;
        }

        public static bool IsSelected(Sheet sheet, List<String> listToSelect)
        {   

           foreach(var element in listToSelect)
           {
                if (element.Equals(sheet.Name.ToString(), StringComparison.OrdinalIgnoreCase))
                    return true;
           }
            return false;
        }

        private DataTable ReadExcelSheet(string fname, bool firstRowIsHeader)
        {
            List<string> Headers = new List<string>();
            DataTable dt = new DataTable();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fname, false))
            {
                //Read the first Sheets 
                Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                int counter = 0;
                foreach (Row row in rows)
                {
                    counter = counter + 1;
                    //Read the first row as header
                    if (counter == 1)
                    {
                        var j = 1;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            var colunmName = firstRowIsHeader ? GetCellValue(doc, cell) : "Field" + j++;
                            Console.WriteLine(colunmName);
                            Headers.Add(colunmName);
                            dt.Columns.Add(colunmName);
                        }
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = GetCellValue(doc, cell);
                            i++;
                        }
                    }
                }

            }
            return dt;
        }

        private void CreateExcelFile(DataTable table, string destination)
        {
           // hfFileName.Value = destination;
           // lblFileName.Text = string.Empty;
            using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                //foreach (System.Data.DataTable table in ds.Tables)
                //{

                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;
                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                sheets.Append(sheet);

                DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }


                sheetData.AppendChild(headerRow);

                foreach (System.Data.DataRow dsrow in table.Rows)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    foreach (String col in columns)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                //}
            }
         //   btnDownloadExcel.Visible = true;
         //   lblFileName.Text = "Servicing file created successfully";
        }
    }
}
