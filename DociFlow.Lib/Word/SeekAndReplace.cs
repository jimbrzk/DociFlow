using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace DociFlow.Lib.Word
{
    /// <summary>
    /// Search text inside Word file and replace it.
    /// </summary>
    public class SeekAndReplace : IDisposable
    {
        private ZipArchive _zip;
        private IEnumerable<ZipArchiveEntry> _zipWordEnties;

        /// <summary>
        /// Open Word file for processing
        /// </summary>
        /// <param name="path">Path to .docx file</param>
        /// <exception cref="InvalidDataException">If Word archive dose not contains word directory with .xml files this exception will be throwen indicating that this file is probably not MS Word one.</exception>
        public void Open(string path)
        {
            _zip = ZipFile.Open(path, ZipArchiveMode.Update);
            _zipWordEnties = _zip.Entries.Where(x => x.FullName.StartsWith("word/") && x.FullName.EndsWith(".xml"));
            if (!(_zipWordEnties?.Any() ?? false)) throw new InvalidDataException("This is probaly not a MS Office Word file.");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="variables">Dictionary with variables, Key is </param>
        /// <param name="openTag">String starting a variable</param>
        /// <param name="closeTag">String ending a variable</param>
        /// <exception cref="NullReferenceException">If Open method was not called or current instance is Disposed</exception>
        public void FindAndReplace(Dictionary<string, string> variables, string openTag = "{{", string closeTag = "}}")
        {
            if (_zip == null || _zipWordEnties == null)
                throw new NullReferenceException("Document is not opened properly!");

            foreach (ZipArchiveEntry archiveEntry in _zipWordEnties)
            {
                string sourceString = null;
                using (Stream sourceStream = archiveEntry.Open())
                {
                    byte[] bytes = new byte[sourceStream.Length];
                    sourceStream.Read(bytes, 0, bytes.Length);
                    sourceString = Encoding.UTF8.GetString(bytes);
                }
                if (String.IsNullOrWhiteSpace(sourceString)) continue;

                MatchCollection matches = Regex.Matches(sourceString, $@"{openTag}.+?<w:t>(.+?)<\/w:t>.+?{closeTag}", RegexOptions.Multiline, TimeSpan.FromMinutes(1));
                foreach (Match match in matches)
                {
                    try
                    {
                        if (match.Success && match.Groups.Count > 1 && variables.ContainsKey(match.Groups[1].Value))
                        {
                            sourceString = sourceString.Replace(match.Value, 
                                match.Value.Replace(match.Groups[1].Value, variables[match.Groups[1].Value])
                                .TrimStart(openTag.ToCharArray())
                                .TrimEnd(closeTag.ToCharArray()));
                        }
                    }
                    catch (Exception) { }
                }

                using (Stream destinationStream = archiveEntry.Open())
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(sourceString);
                    destinationStream.SetLength(bytes.Length);
                    destinationStream.Write(bytes, 0, bytes.Length);
                }

                sourceString = null;
            }
        }

        /// <summary>
        /// Dispose Word file
        /// </summary>
        public void Dispose()
        {
            _zipWordEnties = null;
            _zip?.Dispose();
        }
    }
}
