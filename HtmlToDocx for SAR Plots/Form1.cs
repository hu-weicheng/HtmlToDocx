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

        }

        //转换button
        private void button1_Click(object sender, EventArgs e)
        {
            // ✅ 成功前变白
            button1.BackColor = Color.White;

            tr.inputHtmlPath = this.textBox1.Text;
            tr.outputPath = this.textBox2.Text;
            
            //检查输入文件路径
            //检查输出文件路径
            //如果都没问题就转换
            if(tr.Check(tr.inputHtmlPath,tr.outputPath)==true)
            {
                tr.fileName = Transform.RemoveTrailingDashNumber(tr.inputHtmlPath);
                tr.outputDocxPath = Path.Combine(tr.outputPath, tr.fileName + ".docx");
                tr.Translate();
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
                }
            }
        }
    }
}
