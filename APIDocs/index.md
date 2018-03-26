# Azure Functions for Office
Azure Functions For Office is a set of utility [Azure Functions](https://azure.microsoft.com/en-us/services/functions/) for working with Office documents and eventually Microsoft Graph. 
This initial version consists of three functions for working with Excel and Word documents:
* [Excel Extract](articles/ExcelExtract.html) - reads rows and columns from an Excel file and returns easily consumed JSON
* [Excel Merge](articles/ExcelMerge.html) - Merges data in JSON format with an Excel document given as a URL or as a Base64 encoded string
* [Word Merge](articles/WordMerge.html) - Merges data in JSON format with an Word document given as a URL or as a Base64 encoded string

## Navigating the Documentation
These documents consist of [articles](articles/intro.html) that explain what the functions do and [API documentation for .NET developers](api/index.html) linked to the source code in [GitHub](https://github.com/InstantQuick/AzureFunctionsForOffice).
