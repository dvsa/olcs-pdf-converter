using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using NLog;

namespace OLCSConverter
{
    public class ConvertService
    {
        private Application _wordInstance;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private object _missing = Missing.Value;
        private readonly object _objTrue = true;
        private readonly object _objFalse = false;
        private readonly string _srcPath;
        private readonly string _destPath;
        private readonly bool _canShowWord;
        public Application WordInstance
        {
            get
            {
                if (_wordInstance == null)
                {
                    _wordInstance = new Application();
                    _wordInstance.Visible = _canShowWord;
                }

                return _wordInstance;
            }
        }

        public ConvertService(string srcPath, string destPath, bool canShowWord)
        {
            _srcPath = srcPath;
            _destPath = destPath;
            _canShowWord = canShowWord;
        }

        public void ProcessThroughWord(string fileName)
        {
            try
            {
                _logger.Info($"Converting fileName = {fileName}");

                object objFilePath = Path.Combine(_srcPath, fileName);
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);

                WordInstance.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                var doc = WordInstance.Documents.Open(
                    FileName: ref objFilePath,
                    ConfirmConversions: _missing,
                    ReadOnly: _objFalse,
                    AddToRecentFiles: _missing,
                    PasswordDocument: _missing,
                    PasswordTemplate: _missing,
                    Revert: _missing,
                    WritePasswordDocument: _missing,
                    WritePasswordTemplate: _missing, 
                    Format: _missing, 
                    Encoding: _missing, 
                    Visible: _objTrue);

                //WordInstance.Documents.Open(
                //    FileName: ref objFilePath,
                //    ConfirmConversions: _objFalse,
                //    ReadOnly: _objTrue,
                //    AddToRecentFiles: _objFalse,
                //    PasswordDocument: ref _missing,
                //    PasswordTemplate: ref _missing,
                //    Revert: ref _missing,
                //    WritePasswordDocument: ref _missing,
                //    WritePasswordTemplate: ref _missing,
                //    Format: ref _missing,
                //    Encoding: ref _missing,
                //    Visible: _objFalse, 
                //    OpenAndRepair: _objTrue, 
                //    DocumentDirection: ref _missing, 
                //    NoEncodingDialog: _objTrue, 
                //    XMLTransform: ref _missing);

                doc.Activate();
                doc.SaveAs(Path.Combine(_destPath, $"{fileNameWithoutExt}.pdf"), WdSaveFormat.wdFormatPDF);
                doc.Close(false, ref _missing, ref _missing);
                doc = null;

                _logger.Info($"Successfully converted fileName = {fileName}");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Error converted fileName = {fileName}");
                _logger.Error(ex, ex.Message);
                _logger.Error(ex, ex.StackTrace);

                throw;
            }
            finally
            {
                WordInstance.Quit(_objFalse);
                _wordInstance = null;
            }
        }
    }
}
