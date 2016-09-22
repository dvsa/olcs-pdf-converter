﻿using System;
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

namespace OLCSConverter
{
    public class ConvertController : ApiController
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly string _srcPath;
        private readonly string _destPath;
        private readonly string[] _acceptableExts = {".rtf", ".doc", ".docx"};
        private readonly bool _canShowWord;

        public ConvertService PrintSvc { get; }

        public ConvertController()
        {
            _srcPath = ConfigurationManager.AppSettings["srcDocsPath"];
            _destPath = ConfigurationManager.AppSettings["destDocsPath"];
            _canShowWord = bool.Parse(ConfigurationManager.AppSettings["canShowWord"]);

            PrintSvc = new ConvertService(_srcPath, _destPath, _canShowWord);
        }

        [Route("test")]
        public IEnumerable<string> Get()
        {
            return new[] { "Reached", "OK" };
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
                    _logger.Error($"Could not find generated file = ({filePathToSend}).");
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
            RemoveOldFiles(_srcPath);
            RemoveOldFiles(_destPath);
        }

        private void RemoveOldFiles(string path)
        {
            var twoDaysAgo = DateTime.UtcNow.AddDays(-2);

            var files = Directory.GetFiles(path, "*.*", SearchOption.TopDirectoryOnly);
            foreach (var file in files)
            {
                if (File.GetCreationTimeUtc(file) < twoDaysAgo)
                {
                    File.Delete(file);
                }
            }
        }
    }
}
