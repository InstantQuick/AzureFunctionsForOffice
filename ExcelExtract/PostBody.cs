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
    /// <summary>
    /// Deserialized from JSON that looks like this...
    /// Use either Workbook OR WorkbookUrl
    /// /*
    /// {
    ///     "Workbook": "BASE64STRING",
    ///     "WorkbookUrl": "https://someweb/ExcelFile.xlsx",
    ///     "FieldData": {
    ///         "Field": "Value"
    ///     },
    ///     "RepeatingSectionData": {
    ///         "SectionControlTag": [
    ///             {
    ///                 "Field": "Value"
    ///             }
    ///         ]
    ///     }
    /// }
    /// </summary>
    /// <code>
    /// {
    ///    "Workbook": "BASE64STRING",
    ///    "WorkbookUrl": "https://someweb/ExcelFile.xlsx",
    ///    "FieldData": {
    ///        "Field": "Value"
    ///    },
    ///    "RepeatingSectionData": {
    ///        "SectionControlTag": [
    ///            {
    ///                "Field": "Value"
    ///            }
    ///        ]
    ///    }
    /// } 
    /// </code>
    public class PostBody
    {
        /// <summary>
        /// Base64string of the Excel file with data to extract
        /// </summary>
        public string Workbook { get; set; } = string.Empty;

        /// <summary>
        /// Url of the Excel file with data to extract
        /// </summary>
        public string WorkbookUrl { get; set; } = string.Empty;

        /// <summary>
        /// If true, the first row is column names in the Excel file
        /// </summary>
        public bool FirstRowIsColumnNames { get; set; } = false;

        /// <summary>
        /// The column names to use when the first row is not column names
        /// </summary>
        public List<string> ColumnNames { get; set; } = new List<string>();
    }
}
