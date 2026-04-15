using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ass2Excel
{
    public partial class Form1: Form
    {
        string path;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(AssFileTextBox.Text == "")
            {
                Warning.Text = "请拖入你的ass";
                return;
            }
            if (OutputTextBox.Text == "")
            {
                Warning.Text = "请选择输出路径";
                return;
            }

            var assReader = new AssReader(path);
            assReader.Read();
            var count = assReader.WriteExcel(OutputTextBox.Text, PathToFileName(path).Replace(".ass", ".xlsx"));

            Warning.Text = String.Format("成功，共写入{0}条——{1}", count, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        }

        private void selectButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowNewFolderButton = true;
            if(dialog.ShowDialog() == DialogResult.OK)
            {
                OutputTextBox.Text = dialog.SelectedPath;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && DragFilePath(e).EndsWith(".ass"))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            path = DragFilePath(e);
            AssFileTextBox.Text = PathToFileName(path);
        }

        private string DragFilePath(DragEventArgs e)
        {
            return ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
        }

        private void AutoGetSpeakers(object sender, EventArgs e)
        {
            if (AssFileTextBox.Text == "")
            {
                Warning.Text = "请拖入你的ass";
                return;
            }

            Warning.Text = "Done";
        }

        private string PathToFileName(string path)
        {
            return path.Substring(path.LastIndexOf(@"\") + 1);
        }

        private void PreProcess(object sender, EventArgs e)
        {
            var speakers = AssReader.PrePorcess(path);
            
        }
    }
}
