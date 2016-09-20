﻿using System;
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
        public Application WordInstance
        {
            get
            {
                if (_wordInstance == null)
                {
                    _wordInstance = new Application();
                    _wordInstance.Visible = false;
                }

                return _wordInstance;
            }
        }

        public ConvertService(string srcPath, string destPath)
        {
            _srcPath = srcPath;
            _destPath = destPath;
        }
        
        public void ProcessThroughWord(string fileName)
        {
            try
            {
                _logger.Info($"fileName = {fileName}");

                object objFilePath = Path.Combine(_srcPath, fileName);
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                
                WordInstance.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                
                var doc = WordInstance.Documents.Open(ref objFilePath, _missing, _objFalse, _missing, _missing, 
                                                        _missing, _missing, _missing, _missing, _missing, _missing, _objTrue);
                
                doc.Activate();
                doc.SaveAs(Path.Combine(_destPath, $"{fileNameWithoutExt}.pdf"), WdSaveFormat.wdFormatPDF);
                doc.Close(false, ref _missing, ref _missing);
                doc = null;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error in print service");
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