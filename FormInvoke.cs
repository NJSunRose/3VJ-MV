using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using WorkList;
using Class;


namespace _3VJ_MV
{
    public partial class FormInvoke : Form
    {
        public FormInvoke()
        {
            InitializeComponent();
            //this.listView1.BackColor = Color.Blue;
            this.listView1.View = View.Details;
            this.listView1.GridLines = true;

            this.listView2.GridLines = true;

            ColumnHeader ch = new ColumnHeader();
            ch.Text = "序号";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "板件名称";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "颜色";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "材料";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "封边面1";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "封边面2";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "封边面3";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "封边面4";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "成品长";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "成品宽";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "成品厚";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "面5加工码";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "面6加工码";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "开料长";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "开料宽";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);
            ch = new ColumnHeader();
            ch.Text = "是否有水平孔";
            ch.TextAlign = HorizontalAlignment.Left;
            this.listView1.Columns.Add(ch);


        }
       public static string csvpath;
        List<SWJXMLToCSV.Cabinet> Cabinetlist = null;

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfiledialog = new OpenFileDialog();
            openfiledialog.Multiselect = false;
            openfiledialog.Filter = "三维家文件|*.xml";
            if (openfiledialog.ShowDialog() == DialogResult.OK)
            {
                //T3VProduct t3vproduct = new T3VProduct();
                string xmlpath = openfiledialog.FileName;
                byte[] bytes = null;
                string fileName = xmlpath;
                FileStream fs = new FileStream(fileName, FileMode.Open);
                int streamLength = (int)fs.Length;
                bytes = new byte[streamLength];
                fs.Read(bytes, 0, streamLength);
                fs.Close();

                string base64Str = Convert.ToBase64String(bytes);
                Cabinetlist = SWJXMLToCSV.T3VProduct .LoadFromXML(base64Str);

                //if (Directory.Exists(Path.GetFullPath(xmlpath).Replace(".xml", "")))
                //{
                //    Directory.Delete(Path.GetFullPath(xmlpath).Replace(".xml", ""),true);
                //}

                //Directory.CreateDirectory(Path.GetFullPath(xmlpath).Replace(".xml", ""));

                csvpath = Path.GetFullPath(xmlpath).Replace(".xml", "");

                //t3vproduct.OutputCSV(xmlpath, csvpath);

                ListViewItem item;
                ListViewItem itemmeter;

                int num = 1;
                int nummeter = 1;

                this.listView1.Items.Clear();
                this.listView1.BeginUpdate();

                this.listView2.Items.Clear();
                this.listView2.BeginUpdate();
                
                for (int i = 0; i < Cabinetlist.Count;i++ )
                {
                    foreach (var panel in Cabinetlist[i].Panellist)
                    {
                        item = new ListViewItem();
                        item.ImageIndex = num;
                        item.Text = num.ToString();
                        item.SubItems.Add(panel.Name);
                        item.SubItems.Add(panel.Material);
                        item.SubItems.Add(panel.BasicMaterial);

                        if (panel.Edgelist.Count == 4)
                        {
                            item.SubItems.Add(panel.Edgelist[0].Thickness);
                            item.SubItems.Add(panel.Edgelist[1].Thickness);
                            item.SubItems.Add(panel.Edgelist[2].Thickness);
                            item.SubItems.Add(panel.Edgelist[3].Thickness);
                        }
                        else
                        {
                            item.SubItems.Add("");
                            item.SubItems.Add("");
                            item.SubItems.Add("");
                            item.SubItems.Add("");
                        }

                        item.SubItems.Add(panel.Length);
                        item.SubItems.Add(panel.Width);
                        item.SubItems.Add(panel.Thickness);
                        item.SubItems.Add(panel.Face5ID);
                        item.SubItems.Add(panel.Face6ID);
                        item.SubItems.Add(panel.ActualLength);
                        item.SubItems.Add(panel.ActualWidth);
                        item.SubItems.Add(panel.HasHorizontalHole);

                        item.Tag = panel;

                        this.listView1.Items.Add(item);

                        num++;

                    }
                    this.listView1.EndUpdate();
                    this.listView1.Tag = Cabinetlist;

                    //五金在listview控件中的显示
                    foreach(var metal in Cabinetlist[i].Metallist)
                    {
                        itemmeter = new ListViewItem();
                        itemmeter.ImageIndex = nummeter;
                        itemmeter.Text = nummeter.ToString();
                        itemmeter.SubItems.Add(metal.Id);
                        itemmeter.SubItems.Add(metal.Name);
                        itemmeter.SubItems.Add(metal.PartNumber);
                        itemmeter.SubItems.Add(metal.Num);
                        itemmeter.SubItems.Add(metal.length);
                        itemmeter.SubItems.Add(metal.width);
                        itemmeter.SubItems.Add(metal.height);

                        this.listView2.Items.Add(itemmeter);
                        nummeter++;
                    }
                }
                this.listView2.EndUpdate();
                //五金在listview控件中的显示

            }
        }

        public static string guid = string.Empty;
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string name = Path.GetFileNameWithoutExtension(csvpath);

                //List<string> fileNameList = SWJXMLToCSV.Invoke.OutputCsvByXml(Cabinetlist, "EMS-845562");

                List<string> fileNameList = SWJXMLToCSV.Invoke.OutputCsvByXml(Cabinetlist, name);  //20180528 如果三维家在MV优化的时候发现有错误，则从这里读取优化文件

                if (fileNameList.Count > 0)
                {
                    string[] dirName = fileNameList.First().Split('\\');
                    guid = dirName[dirName.Length - 2];
                }
                MessageBox.Show("导出成功!共" + fileNameList.Count + "个文件！");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        

        private void button3_Click(object sender, EventArgs e)
        {
            string nestpath = csvpath + ".csv";

            if (File.Exists(nestpath))
                File.Delete(nestpath);

            StreamWriter sw = new StreamWriter(nestpath, false, Encoding.Default);
            sw.WriteLine("Nesting优化板件清单,,,,,,,,,,,,,,,,,,,");
            sw.WriteLine("板件信息,,,,,,,成品尺寸,,,加工信息,,,,,,,,,");
            sw.WriteLine("序号,材料,封边宽1,封边宽2,封边长1,封边长2,名称,高,宽,数量,正面加工码,反面加工码,备注1（批次号）,备注2（任务编码）,备注3（序列号）,备注4（分流）,备注5,备注6,备注7,备注8");
            foreach (ListViewItem item in listView1.Items)
            {
                SWJXMLToCSV. Panel panel = item.Tag as SWJXMLToCSV.Panel;
                SWJXMLToCSV.ClassEntity nest = new SWJXMLToCSV.ClassEntity();
                nest.Index = panel.ID;
                //nest.Material = panel.Material;

                if (panel.Thickness == "9")
                    panel.Thickness = "8";

                nest.Material = panel.Thickness + "mm" + "闪电黑GG7022" + "E0级刨花板";
                nest.EbW1 = "1mm闪电黑GG7022封边条";
                nest.EbW2 = "1mm闪电黑GG7022封边条";
                nest.EbL1 = "1mm闪电黑GG7022封边条";
                nest.EbL2 = "1mm闪电黑GG7022封边条";
                //nest.EbW1 = panel.Edgelist[0].Thickness;
                //nest.EbW2 = panel.Edgelist[1].Thickness;
                //nest.EbL1 = panel.Edgelist[2].Thickness;
                //nest.EbL2 = panel.Edgelist[3].Thickness;
                nest.PartName = panel.Name;
                nest.Length = panel.Length;
                nest.Width = panel.Width;
                nest.Num = "1";

                nest.F5FileName = panel.Face5ID.Substring(0, 6) + panel.Face5ID.Substring(7, 3) + "X";
                nest.F6FileName = panel.Face6ID.Substring(0, 6) + panel.Face6ID.Substring(7, 3) + "Y";

                string F5csv = Path.Combine(csvpath+"\\"+guid, nest.F5FileName + ".csv");
                string F6csv = Path.Combine(csvpath+"\\"+guid, nest.F6FileName + ".csv");

                if (!File.Exists(F5csv))
                {
                    if (File.Exists(F6csv))
                    {
                        


                        nest.F6FileName = "";
                    }
                    else
                    {
                        nest.F5FileName = "";
                        nest.F6FileName = "";
                    }
                }
                else
                {
                    if (!File.Exists(F6csv))
                    {
                        nest.F6FileName = "";
                    }

                }

                nest.BatchNum = Path.GetFileNameWithoutExtension(nestpath);
                nest.BoxNumber = panel.cabinet.Id.Substring(0,2);
                if (nest.BoxNumber == "")
                {
                    nest.BoxNumber = "11111111";
                }
                nest.PartNumber = nest.BatchNum + nest.Index + nest.Num.PadLeft(2,'0');
                nest.ModelName = panel.cabinet.Name;
                nest.F5FTPAdress = "";
                nest.F6FTPAdress = "";
                nest.Nest_Num = "";
                nest.Order = panel.cabinet.OrderNo;
                nest.LineNumber = "";
            
                sw.WriteLine(nest.OutPutCsvString());

            }
            sw.Flush();
            sw.Close();
            MessageBox.Show("排程单生成成功!");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //System.Environment.Exit(System.Environment.ExitCode);
            //this.Dispose();
            this.Close();
        }

    }
}