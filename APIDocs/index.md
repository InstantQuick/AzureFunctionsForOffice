# Azure Functions for Office
AzureFunctionsForOffice is a multi-tenant, multi-addin back-end for SharePoint add-ins built on [Azure Functions](https://azure.microsoft.com/en-us/services/functions/). 
The goal of this project is to provide the minimal set of functions necessary to support the common scenarios shared by most SharePoint provider hosted add-ins cheaply and reliably.

Features include:
* Centralized Identity and ACS token management 
* Installation and provisioning of add-in components to SharePoint
* Remote event dispatching to add-in specific back-end services via message queues including
  * App installation
  * App launch
  * SharePoint Remote Events

## Navigating the Documentation
These documents consist of [articles](articles/intro.html) that explain what the functions do, how to set up the hosting environment, and how to use the functions in your add-ins and [API documentation for .NET developers](api/index.html) linked to the source code in [GitHub](https://github.com/InstantQuick/AzureFunctionsForOffice).

## A Note on Terminology
These documents use the term **client** to refer to a given SharePoint add-in. A client is identified using its **client ID** which is the GUID that identifies the add-in's ACS client ID in the [SharePoint add-in's AppManifest.xml](https://msdn.microsoft.com/en-us/library/office/fp179918.aspx#AppManifest).

## Functions
There are five functions in this function app.
