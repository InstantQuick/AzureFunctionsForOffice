using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using AzureFunctionsForOffice.Functions;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;

namespace AzureFunctionsForOffice.Host
{
    /// <summary>
    /// 
    /// </summary>
    public static class ExcelExtract
    {
        [FunctionName("ExcelExtract")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "ExcelExtract")]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            Log(log, $"C# HTTP trigger function processed a request! RequestUri={req.RequestUri}");
            var func = new ExcelExtractHandler(req);
            func.FunctionNotify += (sender, args) => Log(log, args.Message);

            var functionArgs = new ExcelExtractHandlerFunctionArgs
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
