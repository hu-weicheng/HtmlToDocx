using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;    //AltChunk

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeDocx
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("参数数量太少。");
                return;
            }
            //string docroot = @"C:\Users\Raise\Desktop\TestSpire\2601U29087E-SA(LF1013)India";
            string docroot = args[0];
            //搜集docx文件
            List<string> doclist = new List<string>();
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

            string outputPath = docroot + @" Merge SAR Plots.docx";
            MergeDocsEmbedded(doclist.ToArray(), outputPath);
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
    }
}
