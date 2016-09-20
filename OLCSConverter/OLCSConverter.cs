using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using Microsoft.Owin.Hosting;
using NLog;

namespace OLCSConverter
{
    public partial class OLCSConverter : ServiceBase
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private IDisposable _server;
        public OLCSConverter()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            StartOptions opts = new StartOptions();
            opts.Urls.Add("http://localhost:8080");
            opts.Urls.Add("http://+:8080");
            _server = WebApp.Start<Startup>(opts);
        }
        
        protected override void OnStop()
        {
            if (_server != null)
            {
                _server.Dispose();
            }

            CloseWordInstances();

            base.OnStop();
        }

        private void CloseWordInstances()
        {
            var processes = Process.GetProcessesByName("WinWord");

            foreach (var process in processes)
            {
                process.Kill();
            }
        }
    }
}
