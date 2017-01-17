using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Collections;

namespace COWTranslation
{
    public class DocMerger
    {
        public delegate void IntegerProgressHandler(int i);
        public delegate void StringProgressHandler(int i, string txt);
        public event IntegerProgressHandler IntegerProgress;
        public event StringProgressHandler StringProgress;

        private string _DestinationFile;

        public string DestinationFile
        {
            get { return _DestinationFile; }
            set { _DestinationFile = value; }
        }
        private string[] _WordFiles;

        public string[] WordFiles
        {
            get { return _WordFiles; }
            set { _WordFiles = value; }
        }

        public void Merge()
        {
            if (!File.Exists(_DestinationFile))
            { return; }
            ArrayList ar = new ArrayList();
            float All = (float)(_WordFiles.Length + 1);
            float i = 0f;
            float prog = 0f;
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(_DestinationFile, true))
            {
                var mainPart = myDoc.MainDocumentPart;
                foreach (string Ifile in _WordFiles)
                {
                    string altChunkId = "AltChunkId" + DateTime.Now.Ticks.ToString().Substring(0, 2);
                    if (!ar.Contains(altChunkId))
                    { ar.Add(altChunkId); }
                    else
                    {
                        while (ar.Contains(altChunkId))
                        {
                            Random rnd = new Random();
                            altChunkId = "AltChunkId" + rnd.Next(1000, 100000).ToString();
                        }
                        ar.Add(altChunkId);
                    }
                    var chunk = mainPart.AddAlternativeFormatImportPart(
                        DocumentFormat.OpenXml.Packaging.AlternativeFormatImportPartType.WordprocessingML, altChunkId);
                    using (FileStream fileStream = File.Open(Ifile, FileMode.Open))
                    {
                        chunk.FeedData(fileStream);
                    }
                    var altChunk = new DocumentFormat.OpenXml.Wordprocessing.AltChunk();
                    altChunk.Id = altChunkId;
                    mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Last());
                    i++;
                    prog = (i / All) * 100f;
                    if (StringProgress != null)
                        StringProgress((int)prog, Ifile);
                    if (IntegerProgress != null)
                        IntegerProgress((int)prog);
                }
                mainPart.Document.Save();
            }
            if (StringProgress != null)
                StringProgress((int)100, "");
            if (IntegerProgress != null)
                IntegerProgress((int)100);
        }
    }
}
