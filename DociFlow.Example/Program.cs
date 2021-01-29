using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DociFlow.Example
{
    class Program
    {
        const string JSON = "{\"now\":\"28 - 01 - 2021\",\"studentname\":\"Kowalski Jakub\",\"coursfrom\":\"24 - 11 - 2019\",\"coursto\":\"24 - 03 - 2021\",\"level\":\"  A1\",\"documentnumber\":\"2021 / 8\",\"language\":\"ANGIELSKI\",\"languageGenetive\":\"ANGIELSKIEGO\"}";

        static void Main(string[] args)
        {
            ProcessFile("exampledoc.docx");
            Thread.Sleep(1000);
            ProcessFile("exampledoc2.docx");

            Console.ReadLine();
        }

        private static void ProcessFile(string file)
        {
            string tempFile = Path.GetFileName(Path.GetTempFileName());
            //File.Copy(file, tempFile);

            using (var doci = new DociFlow.Lib.Word.SeekAndReplace())
            {

                doci.Open(tempFile);
                doci.FindAndReplace(System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(JSON));
            }

            string newFile = (Path.GetFileNameWithoutExtension(tempFile) + ".docx");
            File.Move(tempFile, newFile);

            Process.Start(newFile);
        }
    }
}
