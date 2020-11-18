using CefSharp;
using CefSharp.OffScreen;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DociFlow
{
    /// <summary>
    /// Initlialize CEF component
    /// </summary>
    public class Browser : IDisposable
    {
        /// <summary>
        /// Default timeout to handle browser requests
        /// </summary>
        public const int TIMEOUT_SECONDS = 30;
        public readonly int RenderingWaitMs;

        private BrowserSettings _browserSettings;

        /// <summary>
        /// Initlialize btowser
        /// </summary>
        public Browser(int renderingWait)
        {
            RenderingWaitMs = renderingWait;

            Cef.EnableWaitForBrowsersToClose();

            _browserSettings = new BrowserSettings()
            {
                Javascript = CefState.Disabled,
                DefaultEncoding = "UTF-8"
            };
            var settings = new CefSettings()
            {
                WindowlessRenderingEnabled = true,
                CachePath = Path.GetFullPath(Path.Combine(Program.CacheDir.FullName, "cef", "cache"))
            };

            if (!Cef.Initialize(settings))
                throw new Exception("Failed to initialize");

            Program.Logger.Debug($"CEF {Cef.CefVersion} initialized");

            Program.ApplicationClosing += Program_ApplicationClosing;
        }

        /// <summary>
        /// Dispose CEF when app is closing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Program_ApplicationClosing(object sender, EventArgs e)
        {
            this.Dispose();
        }

        /// <summary>
        /// Create PDF document from given source
        /// </summary>
        /// <param name="source"></param>
        /// <param name="destinationPath"></param>
        /// <param name="landscape">PDF landcape orientation</param>
        /// <param name="screenshot">Path to save screenshot</param>
        public void DownloadPdf(string source, string destinationPath, bool landscape = false, string screenshot = null)
        {
            Program.Logger.Debug($"Downloading PDF: {source}");

            using (var browser = new ChromiumWebBrowser(source, _browserSettings))
            {
                bool handled = false;
                Exception error = null;
                DateTime timeout = DateTime.Now.AddSeconds(TIMEOUT_SECONDS);

                while (!browser.IsBrowserInitialized)
                {
                    Thread.Sleep(2000);
                    if (timeout < DateTime.Now) throw new Exception("Failed to initialize CEF!");
                }
                timeout = DateTime.Now.AddSeconds(TIMEOUT_SECONDS);

                browser.LoadError += (sender, frameArgs) =>
                {
                    error = new Exception($"Browser load error {frameArgs.ErrorCode} {frameArgs.ErrorText}");
                    handled = true;
                };
                browser.FrameLoadEnd += (sender, frameArgs) =>
                {
                    try
                    {
                        // Check to ensure it is the main frame which has finished loading
                        // (rather than an iframe within the main frame).
                        if (frameArgs.Frame.IsMain)
                        {
                            new Task(() =>
                            {
                                // Wait a little bit, because Chrome won't have rendered the new page yet.
                                // There's no event that tells us when a page has been fully rendered.
                                Thread.Sleep(RenderingWaitMs);

                                // Wait for tpdf print
                                var task = browser.PrintToPdfAsync(destinationPath, new PdfPrintSettings() { Landscape = landscape });
                                task.Wait();

                                if (!task.Result) throw new Exception("Failed to create PDF");

                                if (!String.IsNullOrEmpty(screenshot))
                                {
                                    using (var bitmap = browser.ScreenshotOrNull(PopupBlending.Main))
                                    {
                                        bitmap.Save(screenshot, ImageFormat.Png);
                                    }
                                }

                                handled = true;
                            }).Start();
                        }
                    }
                    catch (Exception ex)
                    {
                        error = ex;
                        handled = true;
                    }
                };

                while (!handled)
                {
                    Thread.Sleep(300);
                    if (timeout < DateTime.Now)
                        throw new TimeoutException($"Taking screenshot from {source} was taking to long", error);
                }
                if (error != null) throw error;
            }
        }

        /// <summary>
        /// Take screenshot of given source 
        /// </summary>
        /// <param name="source"></param>
        /// <param name="destinationPath"></param>
        public void TakeScreenshot(string source, string destinationPath, int width = 595, int height = 842, bool disableScrolbars = true)
        {
            Program.Logger.Debug($"Taking screenshot: {source}");

            using (var browser = new ChromiumWebBrowser(source, _browserSettings))
            {
                bool handled = false;
                Exception error = null;
                DateTime timeout = DateTime.Now.AddSeconds(TIMEOUT_SECONDS);

                while (!browser.IsBrowserInitialized)
                {
                    Thread.Sleep(2000);
                    if (timeout < DateTime.Now) throw new Exception("Failed to initialize CEF!");
                }
                timeout = DateTime.Now.AddSeconds(TIMEOUT_SECONDS);

                browser.Size = new System.Drawing.Size(width, height);
                browser.LoadError += (sender, frameArgs) =>
                {
                    error = new Exception($"Browser load error {frameArgs.ErrorCode} {frameArgs.ErrorText}");
                    handled = true;
                };
                browser.FrameLoadEnd += (sender, frameArgs) =>
                {
                    try
                    {
                        // Check to ensure it is the main frame which has finished loading
                        // (rather than an iframe within the main frame).
                        if (frameArgs.Frame.IsMain)
                        {
                            // Wait a little bit, because Chrome won't have rendered the new page yet.
                            // There's no event that tells us when a page has been fully rendered.
                            Thread.Sleep(RenderingWaitMs);

                            if (disableScrolbars)
                            {
                                frameArgs.Frame.Browser.MainFrame.ExecuteJavaScriptAsync("document.body.style.overflow = 'hidden'");
                                Thread.Sleep(80); // Wait for script execution
                            }

                            new Task(() =>
                            {
                                // Wait for the screenshot to be taken.
                                var task = browser.ScreenshotAsync();
                                task.Wait();

                                // Save the Bitmap to the path.
                                // The image type is auto-detected via the ".png" extension.
                                task.Result.Save(destinationPath);

                                // We no longer need the Bitmap.
                                // Dispose it to avoid keeping the memory alive.  Especially important in 32-bit applications.
                                task.Result.Dispose();

                                handled = true;
                            }).Start();
                        }
                    }
                    catch(Exception ex)
                    {
                        error = ex;
                        handled = true;
                    }
                };
                browser.Load(source);

                while (!handled)
                {
                    Thread.Sleep(300);
                    if (timeout < DateTime.Now)
                        throw new TimeoutException($"Taking screenshot from {source} was taking to long", error);
                }
                if (error != null) throw error;
            }
        }

        public void Dispose()
        {
            Program.ApplicationClosing -= Program_ApplicationClosing;

            // Clean up Chromium objects.  You need to call this in your application otherwise
            // you will get a crash when closing.
            Cef.Shutdown();
        }
    }
}
