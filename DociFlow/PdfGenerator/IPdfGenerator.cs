using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DociFlow.PdfGenerator
{
    public interface IPdfGenerator : IDisposable
    {
        void DocxToPdf(string wordFilePath, string pdfDestinationPath);
        void HtmlToPdf(string htmlFilePath, string pdfDestinationPath, bool landscape = false);
    }
}
