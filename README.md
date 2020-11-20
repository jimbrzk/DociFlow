# How to replace tag inside Word files

```
var vars = Dictionary<string, string>() 
{
    { "testtag", "Test tag value" }
}
using (var doci = new DociFlow.Lib.Word.SeekAndReplace())
{
    // test.docx contains text with tag "{{testtag}}" it will be replaced with "Test tag value" in less than one seconds!
    doci.Open("test.docx");
    doci.FindAndReplace(vars, "{{", "}}");
}
```

# DociFlow
## Documents converter DOCX -> PDF and HTML -> PDF

Befaore convert lf Microsoft Word files to PDF you must install LibreOffice on your machine.
DociFlow starts LibreOffice process in background to convert Word documents to PDF.
You can provide LibreOfficePath in appsettings.json

If you wan't to convert HTML to PDF, DociFlow dose not need anything, it will start CEF browser in background and make required conversion.

## Run in console
simply start DociFlow.exe with arguments presented below

### Command line arguments
/landsacpe - Create PDF in landspace orientation (required for HTML conversion)
/docfile {path} - path to Microsoft Word file (chose this or /htmlfile)
/htmlfile {path} - path to PDF file (chose this or /docfile)
/destination {path} - path where to save PDF

## Use wrapper
You can use Wrapper inside DociFlow.Lib project

```
bool success = new DociFlow.Lib.Wrapper("{{PATH_TO_DOCLIFLOW.EXE}}").Run("{{PDF_DESTINATION}}", {{LANDSCAPE_TRUE_OR_FALSE}}, "{{HTML_FILE_TO_CONVERT}}, "{{DOCFILE_TO_CONVERT}}"");
```

### Exit codes
OK = 0;
MISSING_ARGUMENTS = -1;
OUT_OF_MEMORY = -6;
UNHANDLED_EXCEPTION = -3;

### Logging
Currently logging is based only on console output.