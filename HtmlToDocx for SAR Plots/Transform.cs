using Spire.Doc;

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
        Document doc = new Document();
        public string inputHtmlPath = string.Empty;
        public string fileName = string.Empty;
        public string outputPath = string.Empty;
        public string outputDocxPath = string.Empty;
        public Transform() 
        {
            this.inputHtmlPath = @"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm";
            this.outputPath = @"C:\Users\Raise\Desktop\TestSpire";
            this.outputDocxPath = @"C:\Users\Raise\Desktop\TestSpire\GSM900_Report.docx";
            //this.fileName = Path.GetFileName(this.inputHtmlPath);
            this.fileName= Path.GetFileNameWithoutExtension(inputHtmlPath);
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
            bool exist = System.IO.File.Exists(this.inputHtmlPath);

            // 取出目录部分
            string outputDir = Path.GetDirectoryName(outputHtmlPath);

            // 如果目录不存在，就创建
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            
            return exist;
        }

        //public bool Translate()
        //{
        //    bool exists = System.IO.File.Exists(this.inputHtmlPath);
        //    //Console.WriteLine(exists);
        //    if(exists)
        //    {
        //        // XHTMLValidationType.None 可容忍普通 HTML 的小瑕疵
        //        doc.LoadFromFile(this.inputHtmlPath,
        //                         FileFormat.Html, Spire.Doc.Documents.XHTMLValidationType.None);
        //        //Console.WriteLine(doc.ToString());
        //        doc.SaveToFile(this.outputDocxPath, FileFormat.Docx);
        //        //doc.SaveToFile(@"C:\Reports\SAR_Report.docx", FileFormat.Docx2016);
        //    }

        //    return exists;
        //}

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
        public static void TestTranslate(string inputHtmlFilePath,string outputDocxFilePath)
        {
            Document doc = new Document();
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
}
