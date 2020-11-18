using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using DociFlow.PdfGenerator;
using MaeveFramework.Helpers;
using Microsoft.Extensions.Configuration;

namespace DociFlow
{
    public class Program
    {
        public static string BaseDir
        {
            get
            {
                return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).Replace("file:\\", string.Empty);
            }
        }
        public static string AppVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public static Dictionary<string, string> Arguments { get; private set; }
        public static readonly MaeveFramework.Logger.Abstractions.ILogger Logger;
        public static event EventHandler ApplicationClosing;
        public static DirectoryInfo CacheDir;

        public static string LibreOfficePath { get; private set; }
        public static int CefRenderingWait { get; private set; }

        static Program()
        {
            MaeveFramework.Logger.LoggingManager.UseConsole();
            Logger = MaeveFramework.Logger.LoggingManager.GetLogger();

            CacheDir = new DirectoryInfo(Path.Combine(BaseDir, "cache"));
        }

        public static void Main(string[] args)
        {
            try
            {
                var builder = new ConfigurationBuilder()
                    .AddJsonFile($"appsettings.json", false, false);
                var configuration = builder.Build();
                LibreOfficePath = configuration.GetSection(nameof(LibreOfficePath)).Value;
                CefRenderingWait = int.Parse(configuration.GetSection(nameof(CefRenderingWait)).Value);

                if (!args.Any() || args[0] == "/help")
                {
                    Console.WriteLine($"DociFlow {AppVersion}");
                    Console.WriteLine("/landsacpe - Create PDF in landspace orientation");
                    Console.WriteLine("/docfile {path} - path to Microsoft Word file");
                    Console.WriteLine("/htmlfile {path} - path to PDF file");
                    Console.WriteLine("/destination {path} - path where to save PDF");

                    Environment.ExitCode = -1;
                    return;
                }

                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                Arguments = args.ArgsToDictionary("/");
                Logger.Info($"DociFlow {AppVersion} basedir: {BaseDir} arguments: {string.Join(" ", args)}");

                if (!Arguments.ContainsKey("destination") || String.IsNullOrEmpty(Arguments["destination"]))
                    throw new ArgumentNullException("/destination", "/destination argument is not valid");

                if (!CacheDir.Exists)
                {
                    CacheDir.Create();
                    CacheDir.Refresh();
                }

                using (IPdfGenerator pdfGenerator = new PdfGenerator.DefaultGenerator())
                {
                    if (Arguments.ContainsKey("docfile"))
                    {
                        if (!File.Exists(Arguments["docfile"])) throw new FileNotFoundException("Source file not exist", Arguments["docfile"]);
                        if (Arguments.ContainsKey("landscape")) Logger.Warn("Ignoring /lanscape argument");

                        pdfGenerator.DocxToPdf(Arguments["docfile"], Arguments["destination"]);
                    }
                    else if (Arguments.ContainsKey("htmlfile"))
                    {
                        if (!File.Exists(Arguments["htmlfile"])) throw new FileNotFoundException("Source file not exist", Arguments["htmlfile"]);
                        pdfGenerator.HtmlToPdf(Arguments["htmlfile"], Arguments["destination"], (Arguments.ContainsKey("landscape") ? true : false));
                    }
                    else
                    {
                        throw new ArgumentNullException("/htmlfile or /docfile", "Invalid source file (missing valid parameter)");
                    }
                }

                stopwatch.Stop();

                Logger.Info($"Completed in {stopwatch.ElapsedMilliseconds}ms");
                Environment.ExitCode = 0;
            }
            catch (OutOfMemoryException ex)
            {
                Logger.Error(ex, "Out of memory excepion - Closeing app!");
                Environment.ExitCode = -6;
            }
            catch (ArgumentNullException ex)
            {
                Logger.Error($"Missing argument: {ex.ParamName}");
                Environment.ExitCode = -1;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Unhandled exception - Closeing app!");
                Environment.ExitCode = -3;
            }
            finally
            {
                OnApplicationCloseing();
            }

            Environment.Exit(Environment.ExitCode);
        }

        private static void OnApplicationCloseing()
        {
            ApplicationClosing?.Invoke(null, EventArgs.Empty);
        }
    }
}
