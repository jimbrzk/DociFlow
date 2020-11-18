using System;
using System.Collections.Generic;
using System.Text;

namespace DociFlow.Lib.Models
{
    public class Template
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime Created { get; set; }
        public DateTime LastChange { get; set; }
        public string Type { get; set; }
        public bool Landscape { get; set; }
        public string TemplatePath { get; set; }
    }
}
