using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using TestSpireDoc;

namespace MergeDocxFiles
{
    public partial class Form1 : Form
    {
        MergeDocxE md = new MergeDocxE();
        FileTree ft = new FileTree();
        public Form1()
        {
            InitializeComponent();

            
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        //输入Docx案件目录
        private void button4_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "请选择输出目录";
                dlg.SelectedPath = textBox2.Text;

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    this.md.inputCaseDir = textBox3.Text = dlg.SelectedPath;
                }
            }
        }

        //输出SAR Plots.docx目录
        private void button3_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "请选择输出目录";
                dlg.SelectedPath = textBox2.Text;

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    this.md.outputDir = textBox2.Text = dlg.SelectedPath;
                }
            }
        }

        //合并button
        private void button6_Click(object sender, EventArgs e)
        {
            textBox4.Text += DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF ") +
    "合并中，请稍等。" + "\r\n";

            if (Transform.CheckDir(this.md.inputCaseDir,this.md.outputDir))
            {
                //搜索所有docx文件
                textBox4.Text += DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF ") +
"搜索所有docx文件，请稍等。" + "\r\n";
                this.ft.filePaths.Clear();
                this.ft.filePaths=FileTree.GetAllDocFiles(this.md.inputCaseDir);

                //校准进度条processBar1
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = this.ft.Count;
                progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
                textBox4.Text += DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF ") +
$"搜索到docx文件{this.ft.Count}个，请稍等。" + "\r\n";


                textBox4.Text += DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF ") +
"开始合并文件" + "\r\n";
                string fileName = Transform.RemoveTrailingDashNumber(this.ft.filePaths[0]);
                string outfilepath = Path.Combine(this.md.outputDir, fileName + " SAR Plots.docx");

                File.Copy(this.ft.filePaths[0], outfilepath, true);

                using (WordprocessingDocument mainDoc =
                    WordprocessingDocument.Open(outfilepath, true))
                {
                    MainDocumentPart mainPart = mainDoc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    for (int i=1;i< this.ft.Count;i++)
                    {
                        var file = this.ft.filePaths[1];
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

                        //更新进度条
                        progressBar1.Value = i + 1;
                        Application.DoEvents();
                    }

                    mainPart.Document.Save();
                }
            }
            else
            {
                textBox4.Text += DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF ") +
    "输入文件或文件夹不存在。" + "\r\n";
            }
        }
    }
}
