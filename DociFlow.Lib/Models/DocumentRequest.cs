using System;
using System.Collections.Generic;
using System.Text;

namespace DociFlow.Lib.Models
{
    public class DocumentRequest
    {
        public Guid Guid { get; set; }
        public DateTime Created { get; set; }
        public DateTime LastChange { get; set; }
        public int TemplateId { get; set; }
        public string Status { get; set; }
        public string PdfPath { get; set; }
        public Guid? StudentId { get; set; }
        public Guid CourseId { get; set; }
        public string Variables { get; set; }
        public string Error { get; set; }
    }
}
