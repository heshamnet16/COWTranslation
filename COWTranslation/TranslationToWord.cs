using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Collections;
namespace COWTranslation
{
    public class TranslationToWord
    {
        #region Properties
        private string _SourceFileName;

        public string SourceFileName
        {
            get { return _SourceFileName; }
            set { _SourceFileName = value; }
        }
        private string _DestinationFileName;

        public string DestinationFileName
        {
            get { return _DestinationFileName; }
            set { _DestinationFileName = value; }
        }
        private bool _AddAsNewPage;

        public bool AddAsNewPage
        {
            get { return _AddAsNewPage; }
            set { _AddAsNewPage = value; }
        }
        private bool _InsertImage;

        public bool InsertImage
        {
            get { return _InsertImage; }
            set { _InsertImage = value; }
        }

        private Dictionary<string, string> _BookMarksToTexts;

        public Dictionary<string, string> BookMarksToTexts
        {
            get { return _BookMarksToTexts; }
            set { _BookMarksToTexts = value; }
        }
        private Dictionary<string, byte[]> _BookMarksToImage;

        public Dictionary<string, byte[]> BookMarksToImage
        {
            get { return _BookMarksToImage; }
            set { _BookMarksToImage = value; }
        }
        #endregion
        private Dictionary<int, Dictionary<string, byte[]>> _PagePictures;
        private Dictionary<int, Dictionary<string, string>> _PageStrings;

        public void AddPage(int key, Dictionary<string, byte[]> pictures, Dictionary<string, string> strings)
        {
            if (_PageStrings == null)
                _PageStrings = new Dictionary<int, Dictionary<string, string>>();
            if (_PagePictures == null)
                _PagePictures = new Dictionary<int, Dictionary<string, byte[]>>();
            _PagePictures.Add(key, pictures);
            _PageStrings.Add(key, strings);
            if (_BookMarksToImage.Count == 0)
                _BookMarksToImage = pictures;
            if (_BookMarksToTexts.Count == 0)
                _BookMarksToTexts = strings; 
        }
        public void AddPage(int key, Dictionary<string, byte[]> pictures)
        {
            if (_PagePictures == null)
                _PagePictures = new Dictionary<int, Dictionary<string, byte[]>>();
            _PagePictures.Add(key, pictures);
            if (_BookMarksToImage.Count == 0)
                _BookMarksToImage = pictures;
        }
        public void AddPage(int key, Dictionary<string, string> strings)
        {
            if (_PageStrings == null)
                _PageStrings = new Dictionary<int, Dictionary<string, string>>();
            _PageStrings.Add(key, strings);
            if (_BookMarksToTexts.Count == 0)
                _BookMarksToTexts = strings;
        }
        public void AddTextBookMarks(string BookMarkName, string text)
        {
            this._BookMarksToTexts.Add(BookMarkName.ToLower(), text);
        }
        public void AddImageBookMarks(string text , System.Drawing.Image img)
        {
            MemoryStream mem = new MemoryStream();
            img.Save(mem,System.Drawing.Imaging.ImageFormat.Jpeg );
            this.BookMarksToImage.Add(text.ToLower(), mem.ToArray());
        }
# region Constructors
        public TranslationToWord(string SourcFile,string destFile)
        {
            _AddAsNewPage = false;
            _BookMarksToImage = new Dictionary<string, byte[]>();
            _BookMarksToTexts = new Dictionary<string, string>();
            _DestinationFileName = destFile;
            _SourceFileName = SourcFile;            
        }
        public TranslationToWord(string SourcFile, string destFile,bool AsNewPage)
        {
            _AddAsNewPage = AsNewPage ;
            _BookMarksToImage = new Dictionary<string, byte[]>();
            _BookMarksToTexts = new Dictionary<string, string>();
            _DestinationFileName = destFile;
            _SourceFileName = SourcFile;
        }
#endregion

        public void DoIt()
        {
            if (!File.Exists(this._SourceFileName))
            {
                throw new Exception("الملف المصدر غير موجود!");
            }
            if (_AddAsNewPage == false && File.Exists(_DestinationFileName))
            {
                File.Delete(_DestinationFileName);
            }
            try 
            {
                if (File.Exists(_DestinationFileName))
                {
                    using (WordprocessingDocument SourceApp = WordprocessingDocument.Open(_SourceFileName, false))
                    {
                        using (WordprocessingDocument DestApp = WordprocessingDocument.Open(_DestinationFileName, true))
                        {
                            Document SourceDoc = SourceApp.MainDocumentPart.Document;
                            Document DestDoc = DestApp.MainDocumentPart.Document;
                            SectionProperties sectProp = DestDoc.Descendants<SectionProperties>().First<SectionProperties>();
                            DestDoc.Descendants<SectionProperties>().First<SectionProperties>().Remove();
                            BookmarkStart[] booksS = SourceDoc.Descendants<BookmarkStart>().ToArray<BookmarkStart>();
                            foreach (int key in _PageStrings.Keys)
                            {
                                Body Mem = new Body(SourceDoc.Body.OuterXml);
                                Mem.Descendants<SectionProperties>().First<SectionProperties>().Remove();
                                foreach (BookmarkStart boS in booksS)
                                {
                                    if (_PagePictures[key].ContainsKey(boS.Name.ToString().ToLower()))
                                    {
                                        ImagePart imagePart = DestDoc.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);

                                        using (MemoryStream stream = new MemoryStream(_PagePictures[key][boS.Name.ToString().ToLower()]))
                                        {
                                            imagePart.FeedData(stream);
                                        }
                                        UpdatePictureData(ref Mem, boS.Name.ToString(), DestDoc.MainDocumentPart.GetIdOfPart(imagePart));
                                    }//end Image if
                                    else if (_PageStrings[key].ContainsKey(boS.Name.ToString().ToLower()))
                                    {
                                        UpadteTextAfterBookmarke(ref Mem, boS.Name.ToString(), _PageStrings[key][boS.Name.ToString().ToLower()].ToString());
                                    }// enf Text If
                                    else
                                    {

                                    }
                                }

                                Mem.RemoveAllChildren<BookmarkStart>();
                                Mem.RemoveAllChildren<BookmarkEnd>();
                                Break brak = new Break();
                                brak.Type = BreakValues.Page;
                                Break brak2 = new Break(); brak2.Type = BreakValues.Page;
                                ParagraphProperties Ppro = SourceDoc.Descendants<ParagraphProperties>().LastOrDefault<ParagraphProperties>();
                                Paragraph PageBre;
                                if (Ppro == null ) 
                                {                                
                                    PageBre = new Paragraph(new Run(new OpenXmlElement[] { brak }));
                                }
                                else
                                {
                                    ParagraphProperties tt = new ParagraphProperties(Ppro.OuterXml);
                                    PageBre = new Paragraph(new Run(new OpenXmlElement[] {tt, brak }));                                                                
                                }
                                Paragraph PageBre2 = new Paragraph(new Run(new OpenXmlElement[] { new LastRenderedPageBreak(), brak2 }));
                                //DestDoc.Body.Append(PageBre);
                                DestDoc.Body.Append(PageBre2);
                                foreach (OpenXmlElement child in Mem.ChildElements)
                                {
                                    if (child is Paragraph || child is Table )
                                    {
                                        if (child.Descendants<BookmarkStart>() != null && child.Descendants<BookmarkStart>().FirstOrDefault<BookmarkStart>() != null)
                                        {
                                            foreach (BookmarkStart bookMs in child.Descendants<BookmarkStart>())
                                            {
                                                bookMs.Remove();
                                                if (child.Descendants<BookmarkEnd>() != null && child.Descendants<BookmarkEnd>().FirstOrDefault<BookmarkEnd>() != null)
                                                {
                                                    foreach (BookmarkEnd bookEn in child.Descendants<BookmarkEnd>())
                                                    {
                                                        bookEn.Remove();
                                                    }
                                                }
                                            }
                                        }
                                        if (child is Paragraph)
                                            DestDoc.Body.Append(new Paragraph(child.OuterXml));
                                        else
                                            DestDoc.Body.Append(new Table(child.OuterXml));
                                    }

                                }
                            }    //End Of for Pages                             
                            DestDoc.Body.Append(sectProp);
                            //DestDoc.Save();
                            GC.Collect();
                        }
                    }
                }
                else
                {
                    File.Copy(_SourceFileName, _DestinationFileName);
                    using (WordprocessingDocument DestApp = WordprocessingDocument.Open(_DestinationFileName, true))
                    {
                        Document DestDoc = DestApp.MainDocumentPart.Document;
                        SectionProperties sectProp = DestDoc.Descendants<SectionProperties>().First<SectionProperties>();
                        DestDoc.Descendants<SectionProperties>().First<SectionProperties>().Remove();
                        BookmarkStart[] booksS = DestDoc.Descendants<BookmarkStart>().ToArray<BookmarkStart>();
                        Body Mem = new Body(DestDoc.Body.OuterXml);
                        foreach (BookmarkStart boS in booksS)
                        {
                            if (_BookMarksToImage.ContainsKey(boS.Name.ToString().ToLower()))
                            {
                                ImagePart imagePart = DestApp.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);

                                using (MemoryStream stream = new MemoryStream(_BookMarksToImage[boS.Name.ToString().ToLower()]))
                                {
                                    imagePart.FeedData(stream);
                                }
                                UpdatePictureData(ref Mem, boS.Name.ToString(), DestApp.MainDocumentPart.GetIdOfPart(imagePart));
                            }//end Image if
                            else if (_BookMarksToTexts.ContainsKey(boS.Name.ToString().ToLower()))
                            {
                                UpadteTextAfterBookmarke(ref Mem, boS.Name.ToString(), _BookMarksToTexts[boS.Name.ToString().ToLower()].ToString());
                            }// enf Text If
                            else
                            {

                            }
                        }
                        Mem.RemoveAllChildren<BookmarkStart>();
                        Mem.RemoveAllChildren<BookmarkEnd>();
                        DestDoc.Body.RemoveAllChildren();
                        foreach (OpenXmlElement child in Mem.ChildElements)
                        {
                            if ( child is Paragraph || child is Table )
                            {
                                if (child.Descendants<BookmarkStart>() != null && child.Descendants<BookmarkStart>().FirstOrDefault<BookmarkStart>() != null)
                                {
                                    foreach (BookmarkStart bookMs in child.Descendants<BookmarkStart>())
                                    {
                                        bookMs.Remove();
                                        if (child.Descendants<BookmarkEnd>() != null && child.Descendants<BookmarkEnd>().FirstOrDefault<BookmarkEnd>() != null)
                                        {
                                            foreach (BookmarkEnd bookEn in child.Descendants<BookmarkEnd>())
                                            {
                                                bookEn.Remove();
                                            }
                                        }
                                    }
                                }
                                if (child is Paragraph )
                                    DestDoc.Body.Append(new Paragraph(child.OuterXml));
                                else
                                    DestDoc.Body.Append(new Table (child.OuterXml));
                            }
                        }
                        DestDoc.Body.Append(sectProp);
                        DestDoc.Save();
                    }//End Of using Destination File
                    if (_PageStrings.Count > 1 || _PagePictures.Count > 1)
                    {
                        int key1=-1;
                        int key2=-1;
                        foreach (int key in _PagePictures.Keys)
                        {                            
                            foreach(string Inkey in _PagePictures[key].Keys)
                            {
                                if (_BookMarksToImage.ContainsKey(Inkey))
                                {
                                    key1 = key;
                                    break;
                                }
                            }
                            if (key1 > -1) break;
                        }
                        foreach (int key in _PageStrings.Keys)
                        {
                            foreach (string Inkey in _PageStrings[key].Keys)
                            {
                                if (_BookMarksToTexts.ContainsKey(Inkey))
                                {
                                    key2 = key;
                                    break;
                                }
                            }
                            if (key2 > -1) break;
                        }
                        if(key1>-1)
                            _PagePictures.Remove(key1);
                        if (key2 > -1)
                            _PageStrings.Remove(key2);
                        DoIt();
                    }
                }
            }
            catch (Exception ex) { throw ex; }
            GC.Collect();            
        }

        private static bool UpdatePictureData(ref Body doc, string BookMark, string RelatianID)
        {
            try
            {
                BookmarkStart[] bookS = doc.Descendants<BookmarkStart>().ToArray<BookmarkStart>();
                if (bookS != null)
                {
                    bool Founded = false;
                    foreach (BookmarkStart boS in bookS)
                    {
                        if (boS.Name.ToString().ToLower() == BookMark.ToLower())
                        {
                            Founded = true;
                            BookmarkEnd boE = null;
                            foreach (BookmarkEnd booe in doc.Descendants<BookmarkEnd>())
                            {
                                if (booe.Id.InnerText == boS.Id.InnerText)
                                {
                                    boE = booe;
                                    break;
                                }
                            }
                            Drawing dra = boS.NextSibling<Drawing>();
                            if (dra == null)

                            {
                                try
                                {
                                    if (boS.NextSibling() != null && boS.NextSibling().HasChildren)
                                        dra = boS.NextSibling().Descendants<Drawing>().FirstOrDefault<Drawing>();
                                    else
                                        if (boS.NextSibling().NextSibling() != null && boS.NextSibling().NextSibling().HasChildren)
                                            dra = boS.NextSibling().NextSibling().Descendants<Drawing>().FirstOrDefault<Drawing>();
                                }
                                catch { }
                            }
                            if (dra != null)
                            {
                                DocumentFormat.OpenXml.Drawing.Blip bli = dra.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First<DocumentFormat.OpenXml.Drawing.Blip>();
                                if (bli != null)
                                {
                                    bli.Embed = RelatianID ;
                                    boS.Remove();
                                    boE.Remove();
                                    return true;
                                }
                                else
                                    return false;
                            }
                            else
                            {
                                Drawing Pdra = boS.PreviousSibling<Drawing>();
                                if (Pdra == null)
                                {
                                    try
                                    {
                                        if (boS.PreviousSibling() != null && boS.PreviousSibling().HasChildren)
                                            Pdra = boS.PreviousSibling().Descendants<Drawing>().FirstOrDefault<Drawing>();
                                        else
                                            if (boS.PreviousSibling().PreviousSibling() != null && boS.PreviousSibling().PreviousSibling().HasChildren)
                                                Pdra = boS.PreviousSibling().PreviousSibling().Descendants<Drawing>().FirstOrDefault<Drawing>();
                                    }
                                    catch { }
                                }

                                if (Pdra != null)
                                {
                                    DocumentFormat.OpenXml.Drawing.Blip bli = Pdra.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First<DocumentFormat.OpenXml.Drawing.Blip>();
                                    if (bli != null)
                                    {
                                        bli.Embed = RelatianID;
                                        boS.Remove();
                                        boE.Remove();
                                        return true;
                                    }
                                    else
                                        return false;
                                }
                                else
                                {
                                    for (OpenXmlElement ele = boS.NextSibling(); ele != null; ele = ele.NextSibling())
                                    {
                                        if (ele != null)
                                        {
                                            if (ele is Drawing)
                                            {
                                                DocumentFormat.OpenXml.Drawing.Blip bli = ele.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First<DocumentFormat.OpenXml.Drawing.Blip>();
                                                if (bli != null)
                                                {
                                                    bli.Embed = RelatianID;
                                                    boS.Remove();
                                                    boE.Remove();
                                                    return true;
                                                }
                                                else
                                                    return false;
                                            }
                                            else
                                            {
                                                if (ele.Descendants<Drawing>() != null && ele.Descendants<Drawing>().FirstOrDefault<Drawing>() != null)
                                                {
                                                    DocumentFormat.OpenXml.Drawing.Blip bli = ele.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault<DocumentFormat.OpenXml.Drawing.Blip>();
                                                    if (bli != null)
                                                    {
                                                        bli.Embed = RelatianID;
                                                        boS.Remove();
                                                        boE.Remove();
                                                        return true;
                                                    }
                                                    else
                                                        return false;   
                                                }
                                            }

                                        }
                                    }
                                    return false;
                                }
                            }
                        }
                    }
                    if (!Founded)
                    {
                        throw new Exception("الإشارة المرجعية \"" + BookMark + "\" غير موجودة");
                        //return false;
                    }
                }
                else
                { return false; }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return false;
        }
        public static bool UpdatePictureData(string BookMark, byte[] img, string FileName)
        {
            try
            {
                if (!System.IO.File.Exists(FileName))
                {
                    return false;
                }
                using (WordprocessingDocument WordFile = WordprocessingDocument.Open(FileName, true))
                {
                    MainDocumentPart main = WordFile.MainDocumentPart;
                    Document doc = main.Document;
                    BookmarkStart[] bookS = doc.Descendants<BookmarkStart>().ToArray<BookmarkStart>();
                    if (bookS != null)
                    {
                        bool Founded = false;
                        foreach (BookmarkStart boS in bookS)
                        {
                            if (boS.Name.ToString().ToLower() == BookMark .ToLower())
                            {
                                Founded = true;
                                BookmarkEnd boE = null;
                                foreach (BookmarkEnd booe in doc.Descendants<BookmarkEnd>())
                                {
                                    if (booe.Id.InnerText  == boS.Id.InnerText )
                                    {
                                        boE = booe;
                                        break;
                                    }
                                }
                                Drawing dra = boS.NextSibling<Drawing>();
                                if (dra == null)
                                {
                                    try
                                    {
                                        if (boS.NextSibling() != null && boS.NextSibling().HasChildren)
                                            dra = boS.NextSibling().Descendants<Drawing>().FirstOrDefault<Drawing>();
                                        else
                                            if (boS.NextSibling().NextSibling() != null && boS.NextSibling().NextSibling().HasChildren)
                                                dra = boS.NextSibling().NextSibling().Descendants<Drawing>().FirstOrDefault<Drawing>();
                                    }
                                    catch { }
                                }
                                if (dra != null)
                                {
                                    DocumentFormat.OpenXml.Drawing.Blip bli = dra.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First<DocumentFormat.OpenXml.Drawing.Blip>();
                                    if (bli != null)
                                    {
                                        var imagpart = main.GetPartById(bli.Embed);
                                        imagpart.FeedData(new MemoryStream(img));
                                        boS.Remove();
                                        boE.Remove();
                                        doc.Save();
                                        return true;
                                    }
                                    else
                                        return false;
                                }
                                else
                                {
                                    Drawing Pdra = boS.PreviousSibling<Drawing>();
                                    if (Pdra == null)
                                    {
                                        try
                                        {
                                            if (boS.PreviousSibling() != null && boS.PreviousSibling().HasChildren)
                                                Pdra = boS.PreviousSibling().Descendants<Drawing>().FirstOrDefault<Drawing>();
                                            else
                                                if (boS.PreviousSibling().PreviousSibling() != null && boS.PreviousSibling().PreviousSibling().HasChildren)
                                                    Pdra = boS.PreviousSibling().PreviousSibling().Descendants<Drawing>().FirstOrDefault<Drawing>();
                                        }
                                        catch { }
                                    }

                                    if (Pdra != null)
                                    {
                                        DocumentFormat.OpenXml.Drawing.Blip bli = Pdra.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First<DocumentFormat.OpenXml.Drawing.Blip>();
                                        if (bli != null)
                                        {
                                            var imagpart = main.GetPartById(bli.Embed);
                                            imagpart.FeedData(new MemoryStream(img));
                                            boS.Remove();
                                            boE.Remove();
                                            doc.Save();
                                            return true;
                                        }
                                        else
                                            return false;
                                    }
                                    else
                                    {
                                        for (OpenXmlElement ele = boS.NextSibling(); ele != null; ele = ele.NextSibling())
                                        {
                                            if (ele != null)
                                            {
                                                if (ele is Drawing)
                                                {
                                                    DocumentFormat.OpenXml.Drawing.Blip bli = ele.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First<DocumentFormat.OpenXml.Drawing.Blip>();
                                                    if (bli != null)
                                                    {
                                                        var imagpart = main.GetPartById(bli.Embed);
                                                        imagpart.FeedData(new MemoryStream(img));
                                                        boS.Remove();
                                                        boE.Remove();
                                                        return true;
                                                    }
                                                    else
                                                        return false;
                                                }
                                                else
                                                {
                                                    if (ele.Descendants<Drawing>() != null && ele.Descendants<Drawing>().FirstOrDefault<Drawing>() != null)
                                                    {
                                                        DocumentFormat.OpenXml.Drawing.Blip bli = ele.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault<DocumentFormat.OpenXml.Drawing.Blip>();
                                                        if (bli != null)
                                                        {
                                                            var imagpart = main.GetPartById(bli.Embed);
                                                            imagpart.FeedData(new MemoryStream(img));
                                                            boS.Remove();
                                                            boE.Remove();
                                                            return true;
                                                        }
                                                        else
                                                            return false;
                                                    }
                                                }

                                            }
                                        }
                                        return false;
                                    }
                                }
                            }
                        }
                        if (!Founded)
                        {
                            throw new Exception("الإشارة المرجعية \"" + BookMark + "\" غير موجودة");
                            //return false;
                        }
                    }
                    else
                    { return false; }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return false;
       }

        private static bool UpadteTextAfterBookmarke(ref Body doc, string bookMark, string txt)
        {
            try
            {
                BookmarkStart[] bookS = doc.Descendants<BookmarkStart>().ToArray<BookmarkStart>();
                if (bookS != null)
                {
                    bool Founded = false;
                    foreach (BookmarkStart boS in bookS)
                    {
                        if (boS.Name.ToString().ToLower() == bookMark.ToLower())
                        {
                            Founded = true;
                            BookmarkEnd boE = boS.NextSibling<BookmarkEnd>();
                            if (boE != null)
                            {
                                Run PreRun = boS.PreviousSibling<Run>();
                                Run NextRun = boS.NextSibling<Run>();
                                Run NewRun = new Run();
                                OpenXmlElement ParentPara = boE.Parent;
                                if (PreRun != null)
                                {
                                    string RunPro = PreRun.Descendants<RunProperties>().FirstOrDefault<RunProperties>().OuterXml;

                                    if (RunPro != null)
                                        NewRun.AppendChild<RunProperties>(new RunProperties(RunPro));
                                    NewRun.AppendChild<Text>(new Text(txt));
                                    ParentPara.InsertAfter<Run>(NewRun, PreRun);
                                    ParentPara.RemoveChild<BookmarkStart>(boS);
                                    ParentPara.RemoveChild<BookmarkEnd>(boE);
                                    return true;
                                }
                                else
                                {
                                    if (NextRun != null)
                                    {
                                        string RunPro = NextRun.Descendants<RunProperties>().FirstOrDefault<RunProperties>().OuterXml;
                                        if (RunPro != null)
                                            NewRun.AppendChild<RunProperties>(new RunProperties(RunPro));
                                        NewRun.AppendChild<Text>(new Text(txt));
                                        ParentPara.InsertBefore<Run>(NewRun, NextRun);
                                        ParentPara.RemoveChild<BookmarkStart>(boS);
                                        ParentPara.RemoveChild<BookmarkEnd>(boE);
                                        return true;
                                    }
                                    else
                                    {
                                        try
                                        {
                                            string PparaPro = ParentPara.Descendants<ParagraphProperties>().FirstOrDefault<ParagraphProperties>().ParagraphMarkRunProperties.OuterXml;
                                            if (PparaPro != null)
                                            {
                                                NewRun.AppendChild<RunProperties>(new RunProperties(PparaPro));
                                            }
                                        }
                                        catch { }
                                        NewRun.AppendChild<Text>(new Text(txt));
                                        ParentPara.Append(NewRun);
                                        ParentPara.RemoveChild<BookmarkStart>(boS);
                                        ParentPara.RemoveChild<BookmarkEnd>(boE);
                                        return true;
                                    }
                                }
                            }
                            else
                                return false;
                        }
                    }
                    if (!Founded)
                    {
                        throw new Exception("الإشارة المرجعية \"" + bookMark + "\" غير موجودة");
                        //return false;
                    }
                }
                else
                { return false; }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return false;
        }
        public static bool UpadteTextAfterBookmarke(string bookMark, string txt, string Filename)
        {
            try
            {
                if (!System.IO.File.Exists(Filename))
                {
                    return false;
                }
                using (WordprocessingDocument WordFile = WordprocessingDocument.Open(Filename, true))
                {
                    MainDocumentPart main = WordFile.MainDocumentPart;
                    Document doc = main.Document;
                    BookmarkStart[] bookS = doc.Descendants<BookmarkStart>().ToArray<BookmarkStart>();
                    if (bookS != null)
                    {
                        bool Founded = false;
                        foreach (BookmarkStart boS in bookS)
                        {
                            if (boS.Name.ToString().ToLower()  == bookMark.ToLower() )
                            {
                                Founded = true;
                                BookmarkEnd boE = boS.NextSibling<BookmarkEnd>();
                                if (boE != null)
                                {
                                    Run PreRun = boS.PreviousSibling<Run>();
                                    Run NextRun = boS.NextSibling<Run>();
                                    Run NewRun = new Run();
                                    OpenXmlElement ParentPara = boE.Parent;
                                    if (PreRun != null)
                                    {
                                        try
                                        {
                                            string RunPro = PreRun.Descendants<RunProperties>().FirstOrDefault<RunProperties>().OuterXml;

                                            if (RunPro != null)
                                                NewRun.AppendChild<RunProperties>(new RunProperties(RunPro));
                                        }
                                        catch
                                        { }
                                        
                                        NewRun.AppendChild<Text>(new Text(txt));
                                        ParentPara.InsertAfter<Run>(NewRun, PreRun);
                                        ParentPara.RemoveChild<BookmarkStart>(boS);
                                        ParentPara.RemoveChild<BookmarkEnd>(boE);
                                        doc.Save();
                                        return true;
                                    }
                                    else
                                    {
                                        if (NextRun != null)
                                        {
                                            try
                                            {
                                                string RunPro = NextRun.Descendants<RunProperties>().FirstOrDefault<RunProperties>().OuterXml;
                                                if (RunPro != null)
                                                    NewRun.AppendChild<RunProperties>(new RunProperties(RunPro));
                                            }
                                            catch { }
                                            NewRun.AppendChild<Text>(new Text(txt));
                                            ParentPara.InsertBefore<Run>(NewRun, NextRun);
                                            ParentPara.RemoveChild<BookmarkStart>(boS);
                                            ParentPara.RemoveChild<BookmarkEnd>(boE);
                                            doc.Save();
                                            return true;
                                        }
                                        else
                                        {
                                            try
                                            {
                                                string PparaPro = ParentPara.Descendants<ParagraphProperties>().FirstOrDefault<ParagraphProperties>().ParagraphMarkRunProperties.OuterXml;
                                                if (PparaPro != null)
                                                {
                                                    NewRun.AppendChild<RunProperties>(new RunProperties(PparaPro));
                                                }
                                            }
                                            catch { }
                                            NewRun.AppendChild<Text>(new Text(txt));
                                            ParentPara.Append(NewRun);
                                            ParentPara.RemoveChild<BookmarkStart>(boS);
                                            ParentPara.RemoveChild<BookmarkEnd>(boE);
                                            doc.Save();
                                            return true;
                                        }
                                    }
                                }
                                else
                                    return false;
                            }
                        }
                        if (!Founded)
                        {
                            throw new Exception("الإشارة المرجعية \"" + bookMark + "\" غير موجودة");
                            //return false;
                        }
                    }
                    else
                    { return false; }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return false;
        }
    }    
}
