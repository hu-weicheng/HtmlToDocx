using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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
        public Form1()
        {
            InitializeComponent();
        }


        //输入Docx案件目录
        private void button4_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "请选择输出目录";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    textBox_inCase.Text = dlg.SelectedPath;
                }
            }
        }



        //合并button
        private void button6_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Test");
            //传参数，调用控制台程序
            Process.Start(new ProcessStartInfo
            {
                FileName = "MergeDocx.exe",
                Arguments = $"\"{textBox_inCase.Text}\"",
                UseShellExecute = true  //点了就跑，不关心进度
            });
        }
        
    }
}
