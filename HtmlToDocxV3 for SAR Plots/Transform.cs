using Spire.Doc;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TestSpireDoc
{
    internal class Transform
    {
        Document doc = new Document();
        public string inputCaseDir=string.Empty;
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
            this.inputCaseDir = @"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India";
        }
        ~Transform()
        {
            if(this.doc!=null)
            {
                doc.Close();
            }
        }


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
        public static void TestTranslate(string inputHtmlFilePath,string outputDocxFilePath)
        {
            Document doc = new Document();
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
}
