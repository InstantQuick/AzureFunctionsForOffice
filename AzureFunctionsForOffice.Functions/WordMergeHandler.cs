using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Script.Serialization;
using AzureFunctionsForOffice.WordMerge;

namespace AzureFunctionsForOffice.Functions
{
    public class WordMergeHandlerFunctionArgs : AzureFunctionArgs { }

    public class WordMergeHandler : AzureFunctionsForOfficeBase
    {
        private readonly HttpRequestMessage _request;
        private readonly HttpResponseMessage _response;

        /// <summary>
        /// Initializes the handler for a given HttpRequestMessage received from the function trigger
        /// </summary>
        /// <param name="request">The current request</param>
        public WordMergeHandler(HttpRequestMessage request)
        {
            _request = request;
            _response = request.CreateResponse();
        }
        public async System.Threading.Tasks.Task<HttpResponseMessage> ExecuteAsync(WordMergeHandlerFunctionArgs args)
        {
            // Get request body
            string mergeJson = await _request.Content.ReadAsStringAsync();
            var js = new JavaScriptSerializer();

            try
            {
                var request = js.Deserialize<PostBody>(mergeJson);
                if (request == null)
                {
                    throw new InvalidOperationException("No merge data");
                }
                if (request.DocumentTemplateUrl == string.Empty && request.DocumentTemplate == string.Empty)
                {
                    throw new InvalidOperationException("Document template is missing");
                }

                byte[] template;
                if(request.DocumentTemplate != string.Empty)
                {
                    template = Convert.FromBase64String(request.DocumentTemplate);
                }
                else
                {
                    template = await (new WebClient()).DownloadDataTaskAsync(request.DocumentTemplateUrl);
                }
                
                var result = Merger.MergeDocumentWithContent(template, request.FieldData, request.RepeatingSectionData);

                _response.Content = new ByteArrayContent(result);
                _response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = request.FileName
                };
                _response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                _response.Content.Headers.ContentType.MediaType = MimeMapping.GetMimeMapping(request.FileName);
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
