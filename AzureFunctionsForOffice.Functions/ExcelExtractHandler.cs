using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Script.Serialization;
using AzureFunctionsForOffice.ExcelExtract;


namespace AzureFunctionsForOffice.Functions
{
    public class ExcelExtractHandlerFunctionArgs : AzureFunctionArgs { }

    public class ExcelExtractHandler : AzureFunctionsForOfficeBase
    {
        private readonly HttpRequestMessage _request;
        private readonly HttpResponseMessage _response;

        /// <summary>
        /// Initializes the handler for a given HttpRequestMessage received from the function trigger
        /// </summary>
        /// <param name="request">The current request</param>
        public ExcelExtractHandler(HttpRequestMessage request)
        {
            _request = request;
            _response = request.CreateResponse();
        }
        public async System.Threading.Tasks.Task<HttpResponseMessage> ExecuteAsync(ExcelExtractHandlerFunctionArgs args)
        {
            // Get request body
            string requestJson = await _request.Content.ReadAsStringAsync();
            var js = new JavaScriptSerializer();

            try
            {
                var request = js.Deserialize<PostBody>(requestJson);
                if (request == null)
                {
                    throw new InvalidOperationException("No merge data");
                }
                if (request.Workbook == string.Empty && request.WorkbookUrl == string.Empty)
                {
                    throw new InvalidOperationException("Document template is missing");
                }

                byte[] workbook;
                if (request.Workbook != string.Empty)
                {
                    workbook = Convert.FromBase64String(request.Workbook);
                }
                else
                {
                    workbook = await (new WebClient()).DownloadDataTaskAsync(request.WorkbookUrl);
                }

                Log("Got workbook.");

                var extractedData = Extractor.Extract(workbook, request);
                var result = (new JavaScriptSerializer()).Serialize(extractedData);
                var columnCount = extractedData.Count > 0 ? extractedData[0].Keys.Count : 0;
                Log($"Got data from {extractedData.Count} rows and {columnCount} columns.");

                _response.Content = new StringContent(result);
                _response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                _response.StatusCode = HttpStatusCode.OK;
                return _response;
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                _response.StatusCode = HttpStatusCode.BadRequest;
                _response.Content = new StringContent(GetErrorPage());
                _response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
                return _response;
            }
        }
    }
}
