# Azure Functions for Office

Azure Functions For Office is a set of utility [Azure Functions](https://azure.microsoft.com/en-us/services/functions/) for working with Office documents and eventually Microsoft Graph. 

This initial version consists of three functions for working with Excel and Word documents:

1. [Excel Extract](ExcelExtract) - reads rows and columns from an Excel file and returns easily consumed JSON
2. [Excel Merge](ExcelMerge) - Merges data in JSON format with an Excel document given as a URL or as a Base64 encoded string
3. [Word Merge](WordMerge) - Merges data in JSON format with an Word document given as a URL or as a Base64 encoded string

The function app contains a fourth utility timer function, [Keep Warm](KeepWarm), which executes every few minutes to keep the functions loaded in consumption plans as to avoid cold starts.

