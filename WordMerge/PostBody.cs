using System.Collections.Generic;

namespace AzureFunctionsForOffice.WordMerge
{
    //Deserialized from JSON that looks like this...
    //Use either DocumentTemplate OR DocumentTemplateUrl
    /*
    {
        "DocumentTemplate": "BASE64STRING",
        "DocumentTemplateUrl": "https://someweb/somedoc.docx",
        "FileName": "Sample.docx",
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
        public string DocumentTemplate { get; set; } = string.Empty;
        public string DocumentTemplateUrl { get; set; } = string.Empty;
        public string FileName { get; set; } = "document.docx";

        //Normal non-repeating fields
        //as name-value pairs
        public Dictionary<string, string> FieldData { get; set; }

        //The keys of MergeData are the tag names for the repeating sections
        //The values are lists of name value collections for the section.
        //Each item in the list is a row, the key is the field name and the value is the field value
        public Dictionary<string, List<Dictionary<string, string>>> RepeatingSectionData { get; set; }
    }
}
