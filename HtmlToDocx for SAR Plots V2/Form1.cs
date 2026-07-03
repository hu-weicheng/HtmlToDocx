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

using TestSpireDoc;//自定义命名空间

namespace HtmlToDocx_for_SAR_Plots
{
    public partial class HtmlToDocx : Form
    {
        Transform tr = new Transform();
        TreeHtmls htmls = new TreeHtmls();
        public HtmlToDocx()
        {
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            //this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void HtmlToDocx_Load(object sender, EventArgs e)
        {
            //button1背景改为白色
            button1.BackColor = Color.White;
        }

        //转换button
        private void button1_Click(object sender, EventArgs e)
        {
            // ✅ 成功前变白
            button1.BackColor = Color.White;



            if (button2.Enabled ==true)
            {
                tr.inputHtmlPath = this.textBox1.Text;
                tr.outputPath = this.textBox2.Text;
                //检查输入文件路径
                //检查输出文件路径
                //如果都没问题就转换
                if (Transform.Check(tr.inputHtmlPath, tr.outputPath) == true)
                {
                    tr.fileName = Transform.RemoveTrailingDashNumber(tr.inputHtmlPath);
                    tr.outputDocxPath = Path.Combine(tr.outputPath, tr.fileName + ".docx");
                    tr.Translate();
                }
                else
                {
                    textBox4.Text += DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF ") +
                        "输入文件或文件夹不存在。" + "\r\n";
                }
            }
            else
            {
                tr.inputCaseDir=this.textBox3.Text;//案件根目录
                tr.outputPath=this.textBox2.Text;//输出根目录
                //先检查输入输出目录
                //建立案件目录，置于输出目录之下
                //输出日志
                if (Transform.CheckDir(tr.inputCaseDir,tr.outputPath))
                {
                    //htmls是从this.textBox3.Text里根目录中扫出来的
                    string folderName = Path.GetFileName(tr.inputCaseDir);
                    string outputfolder = Path.Combine(tr.outputPath, folderName);
                    Directory.CreateDirectory(outputfolder);
                    tr.outputPath=outputfolder;
                    textBox4.Text += DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF ") +
                        "创建输出案件文件夹。" + "\r\n";

                    for(int i=0;i<htmls.htmPaths.Count;i++)
                    {
                        tr.inputHtmlPath = htmls.htmPaths[i];
                        tr.fileName = Transform.RemoveTrailingDashNumber(tr.inputHtmlPath);
                        tr.outputDocxPath = Path.Combine(tr.outputPath, tr.fileName + ".docx");
                        tr.Translate();
                        //更新进度条
                        // ✅ 更新进度条
                        progressBar1.Value = i + 1;
                        Application.DoEvents();
                    }
                    textBox4.Text += DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF ") +
                        "批量转换完成。" + "\r\n";
                }
                else
                {
                    textBox4.Text += DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF ") +
                        "输入文件或文件夹不存在。" + "\r\n";
                }
            }
            

            // ✅ 成功后变绿
            button1.BackColor = Color.LightGreen;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "HTML 文件|*.htm;*.html|所有文件|*.*";
                dlg.InitialDirectory = Environment.GetFolderPath(
                    Environment.SpecialFolder.Desktop
                );

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = dlg.FileName;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "请选择输出目录";
                dlg.SelectedPath = textBox2.Text;

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = dlg.SelectedPath;

                    //button1背景改为白色
                    button1.BackColor = Color.YellowGreen;
                }
            }
        }

        //输入案件根目录
        private void button4_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "请选择案件根目录";
                dlg.SelectedPath = textBox3.Text;

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    textBox3.Text = dlg.SelectedPath;

                    //禁用button2（输入html路劲
                    //灰色label1、button2、
                    button2.Enabled = false;     // 浏览按钮
                    textBox1.Enabled = false;    // 路径输入框
                    label1.Enabled = false;      // 标签文字变灰
                    textBox1.ReadOnly = true;   //希望“视觉上更像只读”
                    textBox1.BackColor = SystemColors.Control;
                    //textBox4日志框显示搜索htm文件
                    //progressBar1进度条初始化
                    textBox4.Text += DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF")+
                        " 正在案件根目录下搜索所有子文件夹的htm文件。" + "\r\n";
                    // progressBar1 进度条初始化
                    progressBar1.Value = 0;
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = 100;
                    progressBar1.Style = ProgressBarStyle.Continuous;

                    // 扫描
                    htmls.htmPaths.Clear();
                    htmls.Scan(textBox3.Text);

                    // 进度条初始化
                    progressBar1.Value = 0;
                    progressBar1.Maximum = htmls.Count == 0 ? 1 : htmls.Count;

                    textBox4.AppendText(DateTime.Now.ToString("yy-MM-dd HH:mm:ss.FFFF") +
                        $" 扫描完成，共发现 {htmls.Count} 个 htm文件.\r\n");
                }
            }
        }
    }
}
