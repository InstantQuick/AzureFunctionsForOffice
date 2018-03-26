using System;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace AzureFunctionsForOffice.ExcelMerge
{
    public class Merger
    {
        public static Byte[] Merge(byte[] workbook, PostBody request)
        {
            using (var templateStream = new MemoryStream(workbook))
            {
                var wb = new XLWorkbook(templateStream);
                FillNamedCells(request, wb);
                FillTables(request, wb);

                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms);
                    return ms.ToArray();
                }
            }
        }

        private static void FillNamedCells(PostBody request, XLWorkbook wb)
        {
            foreach (var cellName in request.NamedCellValues.Keys)
            {
                var cell = wb.Cell(cellName);
                if (cell != null)
                {
                    var stringValue = request.NamedCellValues[cellName];

                    //For merging, this function is very eager to transform to
                    //decimal values. The formatting should be done by the sheet.
                    //You may require different behaviour
                    if (!Decimal.TryParse(stringValue,
                        NumberStyles.AllowCurrencySymbol |
                        NumberStyles.AllowThousands |
                        NumberStyles.AllowDecimalPoint,
                        new CultureInfo("en-US")
                        , out var possibleDecimalValue))
                    {
                        cell.Value = stringValue;
                    }
                    else
                    {
                        cell.Value = possibleDecimalValue;
                    }
                }
            }
        }

        private static void FillTables(PostBody request, XLWorkbook wb)
        {
            foreach (var requestWorksheet in request.Worksheets)
            {
                var worksheet = wb.Worksheets.FirstOrDefault(ws => ws.Name == requestWorksheet.Key);
                if (worksheet != null)
                {
                    foreach (var valueTable in requestWorksheet.Value.Tables)
                    {
                        var excelTable = worksheet.Tables.FirstOrDefault(t => t.Name == valueTable.Key);
                        if (excelTable != null)
                        {
                            //Excel is 1 based
                            //The table has a header, so row 2 is the first line of data
                            var currentExcelRow = 2;
                            for (int lineIndex = 0; lineIndex < valueTable.Value.Count; lineIndex++)
                            {
                                if (currentExcelRow <= valueTable.Value.Count)
                                {
                                    excelTable.Row(currentExcelRow).InsertRowsBelow(1);
                                    var range = excelTable.Row(currentExcelRow).RangeAddress;
                                    var rowData = worksheet.Range(range);
                                    excelTable.Row(currentExcelRow + 1).Cell(1).Value = rowData;
                                }

                                var lineItem = valueTable.Value[lineIndex];
                                var currentColumn = 1;
                                foreach (var tableCellValue in lineItem)
                                {
                                    var stringValue = tableCellValue;

                                    if (!Decimal.TryParse(stringValue,
                                        NumberStyles.AllowCurrencySymbol |
                                        NumberStyles.AllowThousands |
                                        NumberStyles.AllowDecimalPoint,
                                        new CultureInfo("en-US")
                                        , out var possibleDecimalValue))
                                    {
                                        excelTable.Cell(currentExcelRow, currentColumn).Value = stringValue;
                                    }
                                    else
                                    {
                                        excelTable.Cell(currentExcelRow, currentColumn).Value = possibleDecimalValue;
                                    }

                                    currentColumn++;
                                }

                                //Don't insert a blank line at the end!
                                if (currentExcelRow <= valueTable.Value.Count)
                                {
                                    currentExcelRow++;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
