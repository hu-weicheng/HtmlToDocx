using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;    //AltChunk

using Spire.Doc;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace TestSpireDoc
{
    internal class TransformCore
    {
        Spire.Doc.Document doc = new Spire.Doc.Document();
        public string inputHtmlPath = string.Empty;
        public string outputDocxPath = string.Empty;

        public TransformCore()
        {
            this.inputHtmlPath = @"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm";
            this.outputDocxPath = @"C:\Users\Raise\Desktop\TestSpire\GSM900_Report.docx";
        }
        ~TransformCore()
        {
            if (this.doc != null)
            {
                doc.Close();
            }
        }
        internal bool Translate_V1()
        {
            bool exists = System.IO.File.Exists(this.inputHtmlPath);
            //Console.WriteLine(exists);
            if (exists)
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

        public bool Translate()
        {
            try
            {
                // XHTMLValidationType.None 可容忍普通 HTML 的小瑕疵
                doc.LoadFromFile(this.inputHtmlPath,
                                 FileFormat.Html, Spire.Doc.Documents.XHTMLValidationType.None);
                //Console.WriteLine(doc.ToString());
                doc.SaveToFile(this.outputDocxPath, FileFormat.Docx);
                //doc.SaveToFile(@"C:\Reports\SAR_Report.docx", FileFormat.Docx2016);

                return true;
            }
            catch
            { return false; }
        }
        public static void TestTranslate(string inputHtmlFilePath, string outputDocxFilePath)
        {
            Spire.Doc.Document doc = new Spire.Doc.Document();
            bool exists = System.IO.File.Exists(inputHtmlFilePath);
            //Console.WriteLine(exists);
            // XHTMLValidationType.None 可容忍普通 HTML 的小瑕疵
            doc.LoadFromFile(inputHtmlFilePath,
                             FileFormat.Html, Spire.Doc.Documents.XHTMLValidationType.None);
            //Console.WriteLine(doc.ToString());
            doc.SaveToFile(outputDocxFilePath, FileFormat.Docx);
            //doc.SaveToFile(@"C:\Reports\SAR_Report.docx", FileFormat.Docx2016);
            doc.Close();
        }
    }
    internal class Transform:TransformCore
    {
        
        public string inputCaseDir=string.Empty;
        
        public string fileName = string.Empty;
        public string outputPath = string.Empty;
        
        public Transform() 
        {
            this.inputHtmlPath = @"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm";
            this.outputPath = @"C:\Users\Raise\Desktop\TestSpire";
            this.outputDocxPath = @"C:\Users\Raise\Desktop\TestSpire\GSM900_Report.docx";
            //this.fileName = Path.GetFileName(this.inputHtmlPath);
            this.fileName= Path.GetFileNameWithoutExtension(inputHtmlPath);
            this.inputCaseDir = @"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India";
        }
        //析构会先子类后基类，自动调用。


        public static string RemoveTrailingDashNumber(string filePath)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            // 匹配 " - 1" "-2" ... "-999" 等，仅在末尾
            return Regex.Replace(fileName, @"-\d+$", "");
        }
        public static bool Check(string inputHtmlPath, string outputHtmlPath)
        {
            bool exist = System.IO.File.Exists(inputHtmlPath);

            // 取出目录部分
            string outputDir = Path.GetDirectoryName(outputHtmlPath);

            // 如果目录不存在，就创建
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            
            return exist;
        }

        public static bool CheckDir(string inputDir, string outputPath)
        {
            bool exist = Directory.Exists(inputDir);

            // 取出目录部分
            string outputDir = Path.GetDirectoryName(outputPath);

            // 如果目录不存在，就创建
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            return exist;
        }


    }


    //输入root
    //Case_2601U39467E/
    //├─ GSM 900 Head Left Cheek Mid-2/
    //│   ├─ GSM 900 Head Left Cheek Mid-2.htm ✅
    //├─ GSM 1800 Head Right Cheek/
    //│   ├─ GSM 1800 Head Right Cheek - 1.htm ✅
    //├─ Logs/
    //│   ├─ debug.log
    /***********************************************/
    //2601U39467E-SA(OD1806)India/
    //├─ GSM 900 Head Left Cheek Mid-2/
    //│   ├─ GSM 900 Head Left Cheek Mid-2.htm ✅
    //├─ GSM 1800 Head Right Cheek/
    //│   ├─ GSM 1800 Head Right Cheek - 1.htm ✅
    //├─ GSM 900 Body Back Mid-2/
    //│   ├─ GSM 900 Body Back Mid-2.htm ✅
    //├─ GSM 1800 Body Back Mid-2/
    //│   ├─ GSM 1800 Body Back Mid-2.htm ✅
    //输出
    //Case_2601U39467E\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm
    //Case_2601U39467E\GSM 1800 Head Right Cheek\GSM 1800 Head Right Cheek - 1.htm
    internal class TreeHtmls
    {
        public List<string> htmPaths = new List<string>();
        public string rootDir = string.Empty;

        public int Count { get { return htmPaths.Count; }  }

        public void Scan(string rootDir)
        {
            htmPaths.Clear();

            if (!Directory.Exists(rootDir))
                return;

            foreach (var file in Directory.EnumerateFiles(
                rootDir, "*.*", SearchOption.AllDirectories))
            {
                string ext = Path.GetExtension(file).ToLower();
                if (ext == ".htm" || ext == ".html")
                {
                    htmPaths.Add(file);
                }
            }
        }
        public static List<string> GetAllHtmFiles(string rootDir)
        {
            List<string> htmPaths = new List<string>();

            if (!Directory.Exists(rootDir))
                return htmPaths;

            // SearchOption.AllDirectories = 包含所有子文件夹
            foreach (string file in Directory.GetFiles(
                rootDir,
                "*.*",
                SearchOption.AllDirectories))
            {
                string ext = Path.GetExtension(file).ToLower();
                if (ext == ".htm" || ext == ".html")
                {
                    htmPaths.Add(file);
                }
            }

            return htmPaths;
        }
    }
    internal class FileTree
    {
        public List<string> filePaths = new List<string>();

        public int Count { get { return filePaths.Count; } }

        public static List<string> GetAllHtmFiles(string rootDir)
        {
            List<string> htmPaths = new List<string>();

            if (!Directory.Exists(rootDir))
                return htmPaths;

            // SearchOption.AllDirectories = 包含所有子文件夹
            foreach (string file in Directory.GetFiles(
                rootDir,
                "*.*",
                SearchOption.AllDirectories))
            {
                string ext = Path.GetExtension(file).ToLower();
                if (ext == ".htm" || ext == ".html")
                {
                    htmPaths.Add(file);
                }
            }

            return htmPaths;
        }
        public static List<string> GetAllDocFiles(string rootDir)
        {
            List<string> docPaths = new List<string>();

            if (!Directory.Exists(rootDir))
                return docPaths;

            // SearchOption.AllDirectories = 包含所有子文件夹
            foreach (string file in Directory.GetFiles(
                rootDir,
                "*.*",
                SearchOption.AllDirectories))
            {
                string ext = Path.GetExtension(file).ToLower();
                if (ext == ".doc" || ext == ".docx")
                {
                    docPaths.Add(file);
                }
            }

            return docPaths;
        }
    }
    internal class MergeDocx
    {
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

    internal class MergeDocxV2
    {
        public WordprocessingDocument mainDoc = null;
        public string outputFilePath = string.Empty;   //合并后最终文件名称
        //文件内容处理的主要工具
        public MainDocumentPart mainPart = null;
        public DocumentFormat.OpenXml.Wordprocessing.Body body = null;

        public void Initinal(string docxFilePath,string outputFilePath)
        {
            // 确保输出目录存在
            string outputDir = Path.GetDirectoryName(outputFilePath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            this.outputFilePath = outputFilePath;
            if (File.Exists(this.outputFilePath)==false)
            {
                File.Delete(this.outputFilePath);
            }
            File.Copy(docxFilePath, this.outputFilePath, true);
            
            this.mainDoc = WordprocessingDocument.Open(this.outputFilePath, true);
            //
            this.mainPart = this.mainDoc.MainDocumentPart;
            this.body = this.mainPart.Document.Body;
        }
        public void ProcessContent(string docxFile)
        {
            // 把待合并 docx 作为部件嵌入主文档
            AlternativeFormatImportPart chunk =
                this.mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML);

            using (var fs = new FileStream(docxFile, FileMode.Open, FileAccess.Read))
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
        public void Save()
        {
            this.mainPart.Document.Save();
            // ✅ 关键：保存整个 WordprocessingDocument，而不是仅 Document
            this.mainDoc.Save();
            //你把第一个 DOCX 原地打开并修改
            //后续 Save()实际是写回原文件
            //AltChunk 指向的 Parts 被污染
            //Word 打开后极易损坏，或只剩第一份内容
        }
        ~MergeDocxV2()
        {
            //if (this.mainDoc != null)
            //{
            //    this.mainDoc.Dispose();
            //}
            this.mainDoc?.Dispose();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="docxFiles"></param>
        /// <param name="outputPath">文件名路径</param>
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
                        mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML);

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
    }

    internal class MergeDocxE2:MergeDocxV2
    {
        public string inputCaseDir = string.Empty;
        public string outputDir = string.Empty;
    }
    internal class MergeDocxE:MergeDocx
    {
        public string inputCaseDir = string.Empty;
        public string outputDir = string.Empty;
    }
}
