using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;

namespace DociFlow.Lib
{
    public class Wrapper
    {
        private readonly string _dociFlowPath;

        public Wrapper(string dociFlowExePath)
        {
            if (!File.Exists(dociFlowExePath)) throw new FileNotFoundException("DociFlow.exe not exist on given path", dociFlowExePath);
            _dociFlowPath = dociFlowExePath;
        }

        public bool Run(string destination, bool landscape, string htmlFile = null, string docFile = null)
        {
            string format = string.Empty;
            if (!string.IsNullOrEmpty(htmlFile))
            {
                format = $"/htmlfile \"{htmlFile}\"";
            }
            else if (!string.IsNullOrEmpty(docFile))
            {
                format = $"/docfile \"{docFile}\"";
            }
            ProcessStartInfo startInfo = new ProcessStartInfo()
            {
                FileName = _dociFlowPath,
                Arguments = string.Format("/destination \"{0}\"{1}{2}", destination, ((landscape) ? " /landscape " : " "), format),
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = false
            };

            var process = new Process()
            {
                StartInfo = startInfo
            };

            try
            {
                process.Start();
                DateTime timeout = DateTime.Now.AddMinutes(3);

                while (true)
                {
                    process.Refresh();

                    if (process == null || process.HasExited)
                    {
                        break;
                    }
                    if (!process.Responding)
                    {
                        process.Kill();
                        throw new Exception("soffice process stops responding!");
                    }
                    if (timeout <= DateTime.Now)
                    {
                        throw new Exception("Timeout");
                    }

                    Thread.Sleep(200);
                }

                if (process != null && !process.HasExited)
                    process.Close();

                return (process.ExitCode == Consts.ExitCodes.OK) ? true : false;
            }
            catch (Exception)
            {
                try
                {
                    process?.Close();
                }
                catch (InvalidOperationException)
                {
                }
                throw;
            }
        }
    }
}
