using System.IO;
using System.Reflection;
using System.Text;

namespace AzureFunctionsForOffice.Functions
{
    public class AzureFunctionsForOfficeBase : FunctionBase
    {
        protected string GetErrorPage()
        {
            try
            {
                return
                    GetFile("AzureFunctionsForOffice.Functions.Resources.Error.html", Assembly.GetExecutingAssembly());
            }
            catch
            {
                return "";
            }
        }

        private string GetFile(string key, Assembly assembly)
        {
            if (assembly == null) return null;
            var stream = assembly.GetManifestResourceStream(key);
            if (stream == null) return null;
            using (var streamReader = new StreamReader(stream))
            {
                return Encoding.UTF8.GetString(ReadFully(streamReader.BaseStream));
            }
        }

        private byte[] ReadFully(Stream input)
        {
            var buffer = new byte[16 * 1024];
            using (var ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }
    }
}
