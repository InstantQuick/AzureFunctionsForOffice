using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using AzureFunctionsForOffice.Functions;
using System.Configuration;
using AzureFunctionsForOffice.WordMerge;

namespace AzureFunctionsForOffice.Host
{
    public static class WordMerge
    {
        /// <summary>
        /// This function provides basic facilities to merge data in JSON format with a Microsoft Word document.
        /// To use this function POST a raw JSON  <see cref="PostBody"/> 
        /// </summary>
        /// <param name="req"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        [FunctionName("WordMerge")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "WordMerge")]HttpRequestMessage req, TraceWriter log)
        {
            Log(log, $"C# HTTP trigger function processed a request! RequestUri={req.RequestUri}");
            var func = new WordMergeHandler(req);
            func.FunctionNotify += (sender, args) => Log(log, args.Message);

            var functionArgs = new WordMergeHandlerFunctionArgs
            {
                StorageAccount = ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
                StorageAccountKey = ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]
            };

            return await Task.Run(() => func.ExecuteAsync(functionArgs));
        }

        private static void Log(TraceWriter log, string message)
        {
            log.Info(message);
        }
    }
}
