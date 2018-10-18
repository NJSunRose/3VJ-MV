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

namespace _3VJ_MV
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        private void btnForm1_Click(object sender, EventArgs e)
        {
            Form1 frm = new Form1();
            frm.ShowDialog();
            txt1.Text = Form1.csvpath;
        }

        private void btnFormInvoke_Click(object sender, EventArgs e)
        {
            FormInvoke frm = new FormInvoke();
            frm.ShowDialog();
            //txt2.Text =AppDomain.CurrentDomain.BaseDirectory + "\\" + FormInvoke.guid;
        }

        private void btn_Click(object sender, EventArgs e)
        {
            if (txt1.Text.Trim().Length == 0 || txt2.Text.Trim().Length == 0)
            {
                MessageBox.Show("请指定路径！");
                return;
            }
            string[] p1Files = Directory.GetFiles(txt1.Text);
            string[] p2Files = Directory.GetFiles(txt2.Text);
            if (p1Files.Length != p2Files.Length)
            {
                MessageBox.Show("两个文件夹所生成的csv数量不一致！");
                return;
            }
            for (int i = 0; i < p1Files.Length; i++)
            {
                string filePath = p1Files[i];
                FileStream aFile = new FileStream(filePath, FileMode.Open);
                StreamReader sr = new StreamReader(aFile, System.Text.Encoding.Default);
                string str1 = sr.ReadToEnd();
                sr.Close();
                aFile.Close();

                filePath = p2Files[i];
                FileStream aFile2 = new FileStream(filePath, FileMode.Open);
                StreamReader sr2 = new StreamReader(aFile2, System.Text.Encoding.Default);
                string str2 = sr2.ReadToEnd();
                sr.Close();
                aFile.Close();
                if (str1 != str2) {
                    MessageBox.Show("csv内容不一致："+filePath );
                    return;
                }
            }
            MessageBox.Show("两个目录的内容完全一致！Invoke Interface验证通过！");
        }
    }
}
