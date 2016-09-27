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
                _logger.Info($"Converting - {fileName}");

                object objFilePath = Path.Combine(_srcPath, fileName);

                var fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                _logger.Debug($"Filename without extension - {fileNameWithoutExt}");
                _logger.Debug($"Setting Word alerts to none - {fileName}");
                WordInstance.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                _logger.Debug($"After set Word alerts to none - {fileName}");
                
                var doc = WordInstance.Documents.Open(
                    FileName: objFilePath,
                    ConfirmConversions: _objFalse,
                    ReadOnly: _objTrue,
                    AddToRecentFiles: _objFalse,
                    PasswordDocument: _missing,
                    PasswordTemplate: _missing,
                    Revert: _missing,
                    WritePasswordDocument: _missing,
                    WritePasswordTemplate: _missing, 
                    Format: _missing, 
                    Encoding: _missing, 
                    Visible: _objFalse, 
                    OpenAndRepair: _objTrue,
                    DocumentDirection: _missing, 
                    NoEncodingDialog: _objTrue, 
                    XMLTransform: _missing);

                _logger.Debug($"Opened document - {fileName}");
                doc.Activate();
                _logger.Debug($"Activated document - {fileName}");
                doc.SaveAs(Path.Combine(_destPath, $"{fileNameWithoutExt}.pdf"), WdSaveFormat.wdFormatPDF);
                _logger.Debug($"Saved PDF - {fileName}");
                doc.Close(false, _missing, _missing);
                _logger.Debug($"Closed document - {fileName}");
                doc = null;

                _logger.Info($"Successfully converted - {fileName}");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Error converting - {fileName}");
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
