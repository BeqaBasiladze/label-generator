using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrackingDocGenerator.Services
{
    public class PrinterService
    {
        public void Print(string docxPath)
        {
            if (string.IsNullOrEmpty(docxPath) || !System.IO.File.Exists(docxPath))
                return;

            ProcessStartInfo info = new ProcessStartInfo
            {
                Verb = "Print",
                FileName = docxPath,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            try
            {
                using (Process process = new Process())
                {
                    process.StartInfo = info;
                    process.Start();

                    process.WaitForExit(10000);
                }
            }
            catch(Exception ex)
            {
                throw new ArgumentException("Print error!", nameof(docxPath), ex);
            }
        }
    }
}
