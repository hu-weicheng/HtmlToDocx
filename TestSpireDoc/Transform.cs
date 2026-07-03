using Spire.Doc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestSpireDoc
{
    internal class Transform
    {
        Document doc = new Document();
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
