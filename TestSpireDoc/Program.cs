using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire;
using TestSpireDoc;
using Spire.Doc;
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
            Document doc = new Document();
            bool exists = System.IO.File.Exists(@"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm");
            Console.WriteLine(exists);
            // XHTMLValidationType.None 可容忍普通 HTML 的小瑕疵
            doc.LoadFromFile(@"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm",
                             FileFormat.Html,Spire.Doc.Documents.XHTMLValidationType.None);
            Console.WriteLine(doc.ToString());
            doc.SaveToFile(@"C:\Users\Raise\Desktop\TestSpire\GSM900_Report.docx", FileFormat.Docx);
            //doc.SaveToFile(@"C:\Reports\SAR_Report.docx", FileFormat.Docx2016);
            doc.Close();
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
