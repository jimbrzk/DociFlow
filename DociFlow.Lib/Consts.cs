using System;
using System.Collections.Generic;
using System.Text;

namespace DociFlow.Lib
{
    public class Consts
    {
        public class ExitCodes
        {
            public const int OK = 0;
            public const int MISSING_ARGUMENTS = -1;
            public const int OUT_OF_MEMORY = -6;
            public const int UNHANDLED_EXCEPTION = -3;
        }

        public class TemplateTypes
        {
            public const string HTML = "HTML";
            public const string WORD = "DOC";
        }

        public class DocumentRequestStatus
        {
            public const string ERROR = "ERROR";
            public const string READY = "READY";
            public const string PROCESSING = "PROCESSING";
            public const string WAITING = "WAITING";
        }
    }
}
