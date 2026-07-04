using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;    //AltChunk
using Spire.Doc;
using Spire.Doc.Interface;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestSpireDoc
{
    internal class Transform
    {
        Spire.Doc.Document doc = new Spire.Doc.Document();
        public string inputHtmlPath = string.Empty;
        public string outputDocxPath = string.Empty;
        public Transform() 
        {
            this.inputHtmlPath = @"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm";
            this.outputDocxPath = @"C:\Users\Raise\Desktop\TestSpire\GSM900_Report.docx";
        }
        ~Transform()
        {
            if(this.doc!=null)
            {
                doc.Close();
            }
        }
        public bool Check(string inputHtmlPath, string outputHtmlPath)
        {
            return System.IO.File.Exists(this.inputHtmlPath)& System.IO.File.Exists(this.outputDocxPath);
            //return false;
        }
        public bool Translate()
        {
            bool exists = System.IO.File.Exists(this.inputHtmlPath);
            //Console.WriteLine(exists);
            if(exists)
            {
                // XHTMLValidationType.None 可容忍普通 HTML 的小瑕疵
                doc.LoadFromFile(this.inputHtmlPath,
                                 FileFormat.Html, Spire.Doc.Documents.XHTMLValidationType.None);
                //Console.WriteLine(doc.ToString());
                doc.SaveToFile(this.outputDocxPath, FileFormat.Docx);
                //doc.SaveToFile(@"C:\Reports\SAR_Report.docx", FileFormat.Docx2016);
            }

            return exists;
        }

        public static void TestTranslate(string inputHtmlFilePath,string outputDocxFilePath)
        {
            Spire.Doc.Document doc = new Spire.Doc.Document();
            bool exists = System.IO.File.Exists(@"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm");
            //Console.WriteLine(exists);
            // XHTMLValidationType.None 可容忍普通 HTML 的小瑕疵
            doc.LoadFromFile(@"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm",
                             FileFormat.Html, Spire.Doc.Documents.XHTMLValidationType.None);
            //Console.WriteLine(doc.ToString());
            doc.SaveToFile(@"C:\Users\Raise\Desktop\TestSpire\GSM900_Report.docx", FileFormat.Docx);
            //doc.SaveToFile(@"C:\Reports\SAR_Report.docx", FileFormat.Docx2016);
            doc.Close();
        }
    }

    internal class MergeDocx
    {
        public static void MergeDocsByAltChunk(string[] docxFiles, string outputPath)
        {
            //以第一个文件为基础
            File.Copy(docxFiles[0], outputPath, true);

            using (WordprocessingDocument mainDoc =
                WordprocessingDocument.Open(outputPath, true))
            {
                MainDocumentPart mainPart = mainDoc.MainDocumentPart;
                DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                foreach (var file in docxFiles.Skip(1))
                {
                    //var rel = mainPart.AddExternalRelationship(
                    //    AlternativeFormatImportPart.RelationshipType,
                    //    new Uri(file, UriKind.Absolute)
                    //);
                    var chunk = mainPart.AddAlternativeFormatImportPart(
                        AlternativeFormatImportPartType.WordprocessingML);

                    var rel = mainPart.AddExternalRelationship(
                        chunk.RelationshipType,
                        new Uri(file, UriKind.Absolute)
                    );
                    //var rel = mainPart.AddExternalRelationship(
                    //                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
                    //                new Uri(file, UriKind.Absolute)
                    //            );

                    mainPart.AddAlternativeFormatImportPart(
                        AlternativeFormatImportPartType.WordprocessingML,
                        rel.Id
                    );

                    body.Append(new AltChunk { Id = rel.Id });
                }

                mainPart.Document.Save();
            }
        }

        public static void MergeDocsEmbedded(string[] docxFiles, string outputPath)
        {
            File.Copy(docxFiles[0], outputPath, true);

            using (WordprocessingDocument mainDoc =
                WordprocessingDocument.Open(outputPath, true))
            {
                MainDocumentPart mainPart = mainDoc.MainDocumentPart;
                DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                foreach (var file in docxFiles.Skip(1))
                {
                    // 把待合并 docx 作为部件嵌入主文档
                    AlternativeFormatImportPart chunk =
                        mainPart.AddAlternativeFormatImportPart(
                            AlternativeFormatImportPartType.WordprocessingML);

                    using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                    {
                        chunk.FeedData(fs);
                    }

                    //页尾加分割符
                    body.Append(new Paragraph(
                            new Run(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page })
                        ));

                    // ✅ 终极兜底：通过 MainDocumentPart 查 ID
                    string id = mainPart.GetIdOfPart(chunk);
                    body.Append(new AltChunk { Id = id });
                }

                mainPart.Document.Save();
            }
        }

        //页尾加分割符
        //public static void splitPage()
        //{
        //    body.Append(new Paragraph(
        //            new Run(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page })
        //        ));
        //    body.Append(new AltChunk { Id = chunk.Id });
        //}
    }
}
