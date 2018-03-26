using System.Collections.Generic;

namespace AzureFunctionsForOffice.ExcelMerge
{
    //Deserialized from JSON that looks like this...
    //Use either Workbook OR WorkbookUrl
    /*
    {
        "Workbook": "BASE64STRING",
        "WorkbookUrl": "https://someweb/ExcelFile.xlsx",
        "FileName": "example.xlsx",
        "NamedCellValues": {
        "Name": "Doug Ware",
        "Quest": "Demo the Azure Function!"
        },
        "Worksheets": {
            "Detail": {
                "Tables": {
                    "DetailTable": [
                        [ "Apples", "1", "1.50" ],
                        [ "Plums", "4", "0.66" ],
                        [ "Pears", "2", "2.00" ]
                    ]
                }
            }
        }
    }
    */
    public class Worksheet
    {
        public Dictionary<string, List<List<string>>> Tables { get; set; }
    }

    public class PostBody
    {
        public string Workbook { get; set; } = string.Empty;
        public string WorkbookUrl { get; set; } = string.Empty;
        public string FileName { get; set; } = "workbook.xlsx";
        public Dictionary<string, string> NamedCellValues { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, Worksheet> Worksheets { get; set; } = new Dictionary<string, Worksheet>();
    }
}
