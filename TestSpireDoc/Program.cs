using Spire;
using Spire.Doc;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using TestSpireDoc;
//using Spire.License;
//using Spire.Doc.Documents;

//Word 输出无 Evaluation Warning
//免费用于商业项目（有页数限制）
//Word：≤ 500 段落 / 25 表格 / 文件 ≤ 512KB

namespace TestSpireDoc
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Document doc = new Document();
            //bool exists = System.IO.File.Exists(@"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm");
            //Console.WriteLine(exists);
            //// XHTMLValidationType.None 可容忍普通 HTML 的小瑕疵
            //doc.LoadFromFile(@"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm",
            //                 FileFormat.Html,Spire.Doc.Documents.XHTMLValidationType.None);
            //Console.WriteLine(doc.ToString());
            //doc.SaveToFile(@"C:\Users\Raise\Desktop\TestSpire\GSM900_Report.docx", FileFormat.Docx);
            ////doc.SaveToFile(@"C:\Reports\SAR_Report.docx", FileFormat.Docx2016);
            //doc.Close();

            //string input = @"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm";
            ////string sample = "GSM 900 Head Left Cheek Mid-2.htm";
            //string o = Path.GetFileNameWithoutExtension(input);
            //Console.WriteLine(o);
            //Console.WriteLine(RemoveTrailingDashNumber(o));
            //Console.WriteLine(RemoveTrailingDashNumber(input

            //string root = @"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India";
            //var htmDirs = GetAllHtmFiles(root);
            //foreach (var Dir in htmDirs)
            //{
            //    Console.WriteLine(Dir);
            //}    

            string docroot = @"C:\Users\Raise\Desktop\TestSpire\2601U29087E-SA(LF1013)India";
            //搜集docx文件
            List<string> doclist= new List<string>();
            foreach (var file in Directory.EnumerateFiles(
                docroot, "*.*", SearchOption.AllDirectories))
            {
                string ext = Path.GetExtension(file).ToLower();
                if (ext == ".doc" || ext == ".docx")
                {
                    doclist.Add(file);
                }
            }
            foreach (var file in doclist)
            {
                Console.WriteLine(file);
            }

            string outputPath = docroot+ @" Merge SAR Plots.docx";
            MergeDocx.MergeDocsEmbedded(doclist.ToArray(), outputPath);
        }

        public static List<string> GetAllHtmFiles(string rootDir)
        {
            List<string> htmPaths = new List<string>();

            if (!Directory.Exists(rootDir))
                return htmPaths;

            // SearchOption.AllDirectories = 包含所有子文件夹
            foreach (string file in Directory.GetFiles(
                rootDir,
                " *.*",
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
        public static string RemoveTrailingDashNumber(string filePath)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            // 匹配 "-1" "-2" ... "-999" 等，仅在末尾
            return Regex.Replace(fileName, @"-\d+$", "");
        }
        static void Test_nugetsdk()
        {
            Document doc = new Document();
            bool exists = System.IO.File.Exists(@"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm");
            Console.WriteLine(exists);
            // XHTMLValidationType.None 可容忍普通 HTML 的小瑕疵
            doc.LoadFromFile(@"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm",
                             FileFormat.Html, Spire.Doc.Documents.XHTMLValidationType.None);
            Console.WriteLine(doc.ToString());
            doc.SaveToFile(@"C:\Users\Raise\Desktop\TestSpire\GSM900_Report.docx", FileFormat.Docx);
            //doc.SaveToFile(@"C:\Reports\SAR_Report.docx", FileFormat.Docx2016);
            doc.Close();
        }
    }
}
