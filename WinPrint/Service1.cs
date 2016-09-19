using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace WinPrint
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            CreateDocument();
        }

        protected override void OnStop()
        {
        }


        //Create a missing variable for missing value
        private object _missing = System.Reflection.Missing.Value;
        private Application _wordInstance;
        public Application WordInstance
        {
            get
            {
                if (_wordInstance == null)
                {
                    _wordInstance = new Microsoft.Office.Interop.Word.Application();
                    _wordInstance.Visible = false;
                }

                return _wordInstance;
            }
        }

        public void CreateDocument()
        {
            try
            {
                //Create an instance for word app
                //Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

                //Set animation status for word application
                //winword.ShowAnimation = false;

                //Set status for word application is to be visible or not.
                //winword.Visible = false;



                var filePaths = Directory.GetFiles(@"C:\Users\Sachin.Patel\Documents\olcs_files\RTF's\", "*.rtf");

                foreach (var filePath in filePaths)
                {
                    string fileName = Path.GetFileNameWithoutExtension(filePath);

                    var doc = WordInstance.Documents.Open(filePath, _missing, false);

                    //Create a new document
                    //Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);



                    //Save the document
                    //doc.SaveFormat = WdSaveFormat.wdFormatPDF
                    doc.SaveAs($"C:\\temp\\{fileName}.pdf", WdSaveFormat.wdFormatPDF);

                    // object filename = @"c:\temp1.docx";
                    //document.SaveAs2(ref filename);
                    
                    doc.Close(false, ref _missing, ref _missing);
                    doc = null;
                }
                

                //object rtfFilePath = @"C:\Users\sachin.patel\Documents\DummyRTF.rtf";
                //var doc = WordInstance.Documents.Open(rtfFilePath, _missing, true);

                ////Create a new document
                ////Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);



                ////Save the document
                ////doc.SaveFormat = WdSaveFormat.wdFormatPDF
                //doc.SaveAs(@"C:\temp\DummyRTF.pdf", WdSaveFormat.wdFormatPDF);

                //// object filename = @"c:\temp1.docx";
                ////document.SaveAs2(ref filename);
                //doc.Close(ref _missing, ref _missing, ref _missing);
                //doc = null;
                
            }
            catch (Exception ex)
            {

                //Command Failed

            }
            finally
            {
                WordInstance.Quit(ref _missing, ref _missing, ref _missing);
                _wordInstance = null;
            }
        }
    }
}
