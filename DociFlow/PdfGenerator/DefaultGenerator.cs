using MaeveFramework.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DociFlow.PdfGenerator
{
    public class DefaultGenerator : IPdfGenerator
    {
        private Browser _browser;

        public DefaultGenerator()
        {
            
        }

        public void Dispose()
        {
            _browser?.Dispose();
        }

        public void DocxToPdf(string wordFilePath, string pdfDestinationPath)
        {
            if (String.IsNullOrWhiteSpace(Properties.Settings.Default.LibreOfficePath) || !Directory.Exists(Properties.Settings.Default.LibreOfficePath))
                throw new Exception($"Missing LibreOffice instance. Check path in field {nameof(Properties.Settings.Default.LibreOfficePath)} on your settings file.");

            string tempPdfPath = Path.Combine(Path.GetDirectoryName(pdfDestinationPath), Path.GetFileNameWithoutExtension(wordFilePath) + ".pdf");
            Program.Logger.Debug($"Converting DOC to PDF: {wordFilePath} to {tempPdfPath}");

            FileInfo sofficeFi = new FileInfo(Path.Combine(Properties.Settings.Default.LibreOfficePath, "App", "libreoffice", "program", "soffice.exe"));
            if (!sofficeFi.Exists) throw new FileNotFoundException("LibreOffice executable not founded!", sofficeFi.FullName);

            MaeveFramework.Helpers.Retry.Do(() =>
            {
                string cachePath = new Uri(Path.Combine(Program.CacheDir.FullName, "libreoffice")).AbsoluteUri;

                if (File.Exists(tempPdfPath))
                    File.Delete(tempPdfPath);
                if (File.Exists(pdfDestinationPath))
                    File.Delete(pdfDestinationPath);

                var sofficeProcess = new Process()
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = sofficeFi.FullName,
                        Arguments = $"\"-env:UserInstallation={cachePath}\" --headless --norestore --convert-to pdf:writer_pdf_Export --outdir \"{Path.GetDirectoryName(tempPdfPath)}\" \"{wordFilePath}\"",
                    },
                };
                sofficeProcess.Start();

                DateTime timeout = DateTime.Now.AddMinutes(5);
                while (true)
                {
                    if (File.Exists(tempPdfPath))
                    {
                        // Expected, it's OK
                        break;
                    }
                    if (timeout < DateTime.Now)
                    {
                        throw new Exception("Failed to create PDF in given time");
                    }

                    sofficeProcess.Refresh();

                    if (sofficeProcess == null || sofficeProcess.HasExited)
                    {
                        // Exptected, it's OK
                        break;
                    }
                    if(!sofficeProcess.Responding)
                    {
                        sofficeProcess.Kill();
                        throw new Exception("soffice stops responding");
                    }

                    Thread.Sleep(100);
                }
            }, TimeSpan.FromSeconds(10), 2);

            try
            {
                Retry.Do(() => File.Delete(wordFilePath), 1.Seconds(), 3);
                Retry.Do(() => File.Move(tempPdfPath, pdfDestinationPath), 5.Seconds(), 3);

                if (!File.Exists(pdfDestinationPath))
                    throw new FileNotFoundException("Failed to finalize PDF generation", pdfDestinationPath);

                //_browser = new Browser(Properties.Settings.Default.CefRenderingWait);
                //_browser.TakeScreenshot(new Uri(pdfDestinationPath).AbsoluteUri, Path.Combine(Path.GetDirectoryName(pdfDestinationPath), Path.GetFileNameWithoutExtension(pdfDestinationPath) + ".png"));
            }
            catch(Exception)
            {
                throw;
            }
            finally
            {
                _browser?.Dispose();
            }
        }

        public void HtmlToPdf(string htmlFilePath, string pdfDestinationPath, bool landscape = false)
        {
            _browser = new Browser(Properties.Settings.Default.CefRenderingWait);
            _browser.DownloadPdf(new Uri(htmlFilePath).AbsoluteUri, pdfDestinationPath, landscape, Path.Combine(Path.GetDirectoryName(pdfDestinationPath), Path.GetFileNameWithoutExtension(pdfDestinationPath) + ".png"));
            _browser.Dispose();
        }
    }
}
