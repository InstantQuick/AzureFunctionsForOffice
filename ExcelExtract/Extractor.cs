using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AzureFunctionsForOffice.ExcelExtract
{
    public class Extractor
    {
        public static List<Dictionary<string, string>> Extract(byte[] workbook, PostBody request)
        {
            var extractedData = new List<Dictionary<string, string>>();
            var columnNames = request.ColumnNames;
            int startRow = 0;

            using (var documentStream = new MemoryStream(workbook))
            {
                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(documentStream, false))
                {
                    IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>()
                        .Elements<Sheet>();
                    string relationshipId = sheets.First().Id.Value;
                    WorksheetPart worksheetPart =
                        (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                    Worksheet workSheet = worksheetPart.Worksheet;
                    SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                    var rows = sheetData.Descendants<Row>().ToList();

                    if (request.FirstRowIsColumnNames)
                    {
                        startRow = 1;
                        columnNames = GetColumnNames(spreadSheetDocument, rows.FirstOrDefault());
                    }

                    foreach (var row in rows.Skip(startRow))
                    {
                        extractedData.Add(GetRowData(spreadSheetDocument, row, columnNames));
                    }
                }
            }

            return extractedData;
        }

        private static Dictionary<string, string> GetRowData(SpreadsheetDocument document, Row row, List<string> columnNames)
        {
            var rowData = new Dictionary<string, string>();
            foreach (var cell in row.Descendants<Cell>())
            {
                var columnIndex = GetColumnIndexFromName(GetColumnName(cell.CellReference));
                if (columnIndex != null && columnIndex < columnNames.Count)
                {
                    var columnName = columnNames[(int)columnIndex];
                    rowData[columnName] = GetCellValue(document, cell);
                }
            }
            //Ensure the columns
            foreach (var columnName in columnNames)
            {
                if (!rowData.ContainsKey(columnName))
                {
                    rowData[columnName] = "";
                }
            }
            return rowData;
        }

        private static List<string> GetColumnNames(SpreadsheetDocument document, Row row)
        {
            var columnNames = new List<string>();
            foreach (var cell in row.Descendants<Cell>())
            {
                var value = GetCellValue(document, cell);
                if (string.IsNullOrEmpty(value)) break;
                if (columnNames.Contains(value))
                {
                    throw new FormatException($"Duplicate column name {value} is not allowed.");
                }
                columnNames.Add(value);
            }

            return columnNames;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue?.InnerXml;

            if (value != null && cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value ?? string.Empty;
            }
        }

        private static readonly List<char> Letters = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ' ' };

        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Column Name (ie. B)</returns>
        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);

            return match.Value;
        }

        /// <summary>
        /// Given just the column name (no row index), it will return the zero based column index.
        /// </summary>
        /// <param name="columnName">Column Name (ie. A or AB)</param>
        /// <returns>Zero based index if the conversion was successful; otherwise null</returns>
        public static int? GetColumnIndexFromName(string columnName)
        {
            int? columnIndex = null;
            var colLetters = columnName.ToCharArray().ToList();

            for (int i = 0; i < colLetters.Count; i++)
            {
                var c = colLetters[i];
                int? indexValue = Letters.IndexOf(c);

                if (indexValue != -1)
                {
                    if (i == 0 && colLetters.Count > 1)
                    {
                        columnIndex = columnIndex == null ? (indexValue + 1) * 26 : columnIndex + ((indexValue + 1) * 26);
                    }
                    else
                    {
                        columnIndex = columnIndex == null ? indexValue : columnIndex + indexValue;
                    }
                }
            }
            return columnIndex;
        }
    }
}
