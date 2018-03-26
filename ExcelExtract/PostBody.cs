using System.Collections.Generic;

namespace AzureFunctionsForOffice.ExcelExtract
{
    //Deserialized from JSON that looks like this...
    //Use either Workbook OR WorkbookUrl
    /*
    {
        "Workbook": "BASE64STRING",
        "WorkbookUrl": "https://someweb/ExcelFile.xlsx",
        "FieldData": {
            "Field": "Value"
        },
        "RepeatingSectionData": {
            "SectionControlTag": [
                {
                    "Field": "Value"
                }
            ]
        }
    }
     */
    public class PostBody
    {
        public string Workbook { get; set; } = string.Empty;
        public string WorkbookUrl { get; set; } = string.Empty;
        public bool FirstRowIsColumnNames { get; set; } = false;
        public List<string> ColumnNames { get; set; } = new List<string>();
    }
}
