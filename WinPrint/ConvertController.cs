using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;
using NLog;

namespace WinPrint
{
    public class ConvertController : ApiController
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly string _srcPath;
        private readonly string _destPath;
        private readonly string[] _acceptableExts = {".rtf", ".doc", ".docx"};

        public ConvertService PrintSvc { get; }

        public ConvertController()
        {
            _srcPath = ConfigurationManager.AppSettings["srcDocsPath"];
            _destPath = ConfigurationManager.AppSettings["destDocsPath"];

            PrintSvc = new ConvertService(_srcPath, _destPath);
        }

        [Route("print")]
        public IEnumerable<string> Get()
        {
            return new[] { "aa", "bb" };
        }

        [Route("convert-document")]
        public async Task<HttpResponseMessage> Post()
        {
            if (!Request.Content.IsMimeMultipartContent())
            {
                return Request.CreateErrorResponse(HttpStatusCode.UnsupportedMediaType, "Unsupported media type.");
            }
            
            var provider = new MultipartMemoryStreamProvider();
            await Request.Content.ReadAsMultipartAsync(provider);
            
            HttpContent content = provider.Contents.First();
            string uploadedFileName = content.Headers.ContentDisposition.FileName.Trim('\"');
            byte[] uploadedFile = await provider.Contents.First().ReadAsByteArrayAsync();

            if (!_acceptableExts.Any(ext => uploadedFileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Can only accept RTF, DOC, DOCX files.");
            }

            try
            {
                CleanupOldFiles();
                string fileName = DateTime.UtcNow.ToString("yyyy-MM-dd_HH_mm_ss_") + uploadedFileName;
                
                File.WriteAllBytes(Path.Combine(_srcPath, fileName), uploadedFile);
                string filenameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                
                PrintSvc.ProcessThroughWord(fileName);

                var filePathToSend = Path.Combine(_destPath, $"{filenameWithoutExt}.pdf");

                if (!File.Exists(filePathToSend))
                {
                    _logger.Info($"Could not find generated file ({filePathToSend}).");
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Could not find generated file.");
                }

                var result = CreateFileResponse(filePathToSend, uploadedFileName);

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex.Message);
            }

            return Request.CreateErrorResponse(HttpStatusCode.ServiceUnavailable, "Failed to convert document.");
        }

        private HttpResponseMessage CreateFileResponse(string filePathToSend, string uploadedFileName)
        {
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            var stream = new FileStream(filePathToSend, FileMode.Open, FileAccess.Read);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = uploadedFileName;
            return result;
        }

        private void CleanupOldFiles()
        {
            var filePrefix = DateTime.UtcNow.AddDays(-2).ToString("yyyy - MM - dd_*");

            RemoveOldFiles(filePrefix, _srcPath);
            RemoveOldFiles(filePrefix, _destPath);
        }

        private void RemoveOldFiles(string filePrefix, string path)
        {
            var oldFiles = Directory.GetFiles(path, filePrefix);
            foreach (var oldFile in oldFiles)
            {
                File.Delete(oldFile);
            }
        }
    }
}
