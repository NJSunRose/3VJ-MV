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
using System.Text.RegularExpressions;
using System.IO;
using WorkList;
using Class;
using Ionic.Zip;
using Dimeng.FTP;
using SpreadsheetGear;
using System.Data.OleDb;



namespace _3VJ_MV
{
    
    public partial class Form1 : Form
    {
        List<Class_PartType> parttype = new List<Class_PartType>();
        public Form1()
        {
            InitializeComponent();

            #region 设软件小工具版本号V3.2  宋新刚电脑是读取本地D:\模板忽删里的配置文件，其他电脑读取1.20服务器上面的

            string currentversion = "V3.2";

            IniFiles inifile_First = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
            if (inifile_First.ExistINIFile())
            {
                inifile_First.IniWriteValue("SAMEX", "ps", "1");
                inifile_First.IniWriteValue("Version", "ver", currentversion);  //小工具版本号
            }
            else
            {
                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + Path.Combine(Environment.CurrentDirectory, "OrderNo.ini") + " 目录下创建！");
                return;
            }
            #endregion

            #region 读取版本号与软件里的小工具所设的版本号比较
            string versionpath = @"\\192.168.1.20\数据源\模板忽删";
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))//下拉菜单只对我有效果
            {
                versionpath = @"D:\模板忽删";
            }

            IniFiles iniVersion = new IniFiles(Path.Combine(versionpath, "OrderNo.ini"));
            string serverversion = string.Empty;
            if (iniVersion.ExistINIFile())
            {
                serverversion = iniVersion.IniReadValue("Version", "ver");
            }
            else
            {
                MessageBox.Show("服务器检测版本的文件丢失!","警告",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }


            if (!currentversion.Equals(serverversion))
            {
                MessageBox.Show("当前软件的版本与服务器上面的版本不同步\n\n请至\\192.168.1.20\\数据源\\3VJ-SAMEX-MV目录下拷贝最新版本的程序!",
                    "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                button_InputXML.Visible = false;
                button_OutCsv.Visible = false;
                button_OutReport.Visible = false;
                button_Plan.Visible = false;
            }
            #endregion

            button1.Visible = false;
            button2.Visible = false;
            button12.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            button11.Visible = false;
            label2.Visible = false;
            textBox1.Visible = false;
            label3.Visible = false;
            textBox2.Visible = false;
            button_PlanOutReport.Visible = false;
            button_PlanTask.Visible = false;
            button_ModifyCsv.Visible = false;
            button_OutSamexNest.Visible = false;

            ComboBox_3VJ_SMAX.Items.Add("3VJ-PS");
            ComboBox_3VJ_SMAX.Items.Add("3VJ-SMAX");

            ComboBox_3VJ_SMAX.SelectedIndex = 1;

            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))//下拉菜单只对我有效果
            {
                toolStripMenuItem1.Visible = true;
            }
            else
            {
                toolStripMenuItem1.Visible = false;
            }

            //this.listView1.BackColor = Color.Blue;
            this.listView1.View = View.Details;
            this.listView1.GridLines = true;

            this.listView2.GridLines = true;

            #region 每次开启软件的时候，将2号车间生产的板件类型导进内存中
            string path_parttype = Path.Combine(Environment.CurrentDirectory, "PartType.csv");
            Class_PartType classparttype = new Class_PartType();
            StreamReader sr = new StreamReader(path_parttype);
            sr.ReadLine();
            string line = string.Empty;
            while((line = sr.ReadLine()) != null)
            {
                classparttype = new Class_PartType();
                string[] kline = line.Split(',');
                classparttype.Num = kline[0];
                classparttype.PartType = kline[1];
                classparttype.PartNumber = kline[2];
                classparttype.ProductModel = kline[3];
                parttype.Add(classparttype);
            }
            sr.Close();
            #endregion


        }
        public static string csvpath;
        private void button1_Click(object sender, EventArgs e)
        {
            IniFiles inifile_First = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
            if (inifile_First.ExistINIFile())
            {
                int lastdrawernum = -1;
                inifile_First.IniWriteValue("DrawerNum", "draw", (lastdrawernum + 1).ToString());
            }
            else
            {
                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + Path.Combine(Environment.CurrentDirectory, "OrderNo.ini") + " 目录下创建！");
                return;
            }

            OpenFileDialog openfiledialog = new OpenFileDialog();
            openfiledialog.Multiselect = false;
            openfiledialog.Filter = "三维家文件|*.xml";
            string EveryNum = string.Empty;

            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))
            {
                if (!Directory.Exists(@"C:\Users\sxg035.000\Desktop"))
                    openfiledialog.InitialDirectory = @"C:\Users\sxg035\Desktop\三维家XML下载\xml";
                else
                    openfiledialog.InitialDirectory = @"C:\Users\sxg035.000\Desktop\三维家XML下载\xml";
            }

            if (openfiledialog.ShowDialog() == DialogResult.OK)
            {
                T3VProduct t3vproduct = new T3VProduct();
                string xmlpath = openfiledialog.FileName;
                t3vproduct.LoadFromXML(xmlpath);

                if (Directory.Exists(Path.GetFullPath(xmlpath).Replace(".xml", "")))
                {
                    Directory.Delete(Path.GetFullPath(xmlpath).Replace(".xml", ""),true);
                }

                Directory.CreateDirectory(Path.GetFullPath(xmlpath).Replace(".xml", ""));

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
                
                for (int i = 0; i < t3vproduct.Cabinetlist.Count;i++ )
                {
                    foreach (var panel in t3vproduct.Cabinetlist[i].Panellist)
                    {
                        item = new ListViewItem();
                        item.ImageIndex = num;
                        item.Text = num.ToString();
                        item.SubItems.Add(panel.Name);
                        item.SubItems.Add(panel.Material);
                        item.SubItems.Add(panel.BasicMaterial);

                        if (panel.Edgelist.Count == 4)  //20180525
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

                        if (ComboBox_3VJ_SMAX.SelectedIndex - 0 < 0.1) //如果选项卡上选择的是索引号为0,则为原先与普实对接的，为1则为与SMAX对接
                        {
                            item.SubItems.Add(panel.Face5ID);
                            item.SubItems.Add(panel.Face6ID);
                        }
                        else
                        {
                            string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                            IniFiles inifile = new IniFiles(inipath);
                            EveryNum = inifile.IniReadValue("CsvNum", "Num");
                            if (inifile.ExistINIFile())
                            {
                                item.SubItems.Add(EveryNum + panel.Face5ID);
                                item.SubItems.Add(EveryNum + panel.Face6ID);
                            }
                            else
                            {
                                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                return;
                            }
                        }

                        item.SubItems.Add(panel.ActualLength);
                        item.SubItems.Add(panel.ActualWidth);
                        item.SubItems.Add(panel.HasHorizontalHole);

                        item.SubItems.Add(panel.drawer.ToString());  //20180914

                        item.Tag = panel;

                        this.listView1.Items.Add(item);

                        num++;

                    }
                    this.listView1.EndUpdate();
                    this.listView1.Tag = t3vproduct;

                    //五金在listview控件中的显示
                    foreach(var metal in  t3vproduct.Cabinetlist[i].Metallist)
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
                        itemmeter.SubItems.Add(metal.Unit);  //20180413

                        itemmeter.Tag = metal; //将Metal里的数据绑到listview控件中  20180413

                        this.listView2.Items.Add(itemmeter);
                        nummeter++;
                    }
                }
                this.listView2.EndUpdate();
                //五金在listview控件中的显示

            }
            else
            {
                return;
            }
        }
        bool HaveLarger = false;
        private void button2_Click(object sender, EventArgs e)
        {
            HaveLarger = false;
            if (listView1.Items.Count == 0)
            {
                MessageBox.Show("没有导入三维家的XML文件或此XML文件没有需要加工的板件!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
 
            if (textBox1.Text == string.Empty)
            {
                textBox1.Text = "1";
                //errorProvider1.SetError(textBox1, "必须要输入当前XML文件所在板件标签上面对应的柜号!");
                //MessageBox.Show("请输入当前XM0L需要生成的柜号!");
                //return;
            }
            else
            {
                //errorProvider1.SetError(textBox1, string.Empty);
            }

            ArrayList face5list = new ArrayList();
            ArrayList face6list = new ArrayList();

            string csvname = "";
            ////T3VProduct product = this.listView1.Tag as T3VProduct;
            foreach(ListViewItem item in listView1.Items)
            {
                string Cao_5 = "0";  //20180807 初始化没有槽
                string Cao_6 = "0";

                string EveryNum = string.Empty;

                List<fourpoint> point4 = new List<fourpoint>();
                face5list = new ArrayList();
                face6list = new ArrayList();
                
                Panel panel = item.Tag as Panel;
                BorderSequenceEntity borderseq = new BorderSequenceEntity();
                borderseq.BorderSequence = "BorderSequence";
                borderseq.PanelWidth = panel.Width;                     
                borderseq.PanelLength = panel.Length;
                borderseq.PanelThickness = panel.Thickness;

                borderseq.RunField = "4";                       
                borderseq.CurrentFace = "";
                borderseq.PreviousFace = "";
                borderseq.CurrentZoneName = "N";           
                borderseq.FieldOffsetX = "";
                borderseq.FieldOffsetY = "";
                borderseq.FieldOffsetZ = "";
                borderseq.JobName = panel.cabinet.Name;
                borderseq.ItemNumber = "1";

                borderseq.FileName = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "X";
                borderseq.Face6FileName = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "Y";

                if (ComboBox_3VJ_SMAX.SelectedIndex - 0 > 0.1) //如果选项卡上选择的是索引号为0,则为原先与普实对接的，为1则为与SMAX对接
                {
                    string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                    IniFiles inifile = new IniFiles(inipath);
                    EveryNum = inifile.IniReadValue("CsvNum", "Num");
                    if (inifile.ExistINIFile())
                    {
                        borderseq.FileName = EveryNum + borderseq.FileName;
                        borderseq.Face6FileName = EveryNum + borderseq.Face6FileName;
                    }
                    else
                    {
                        MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                        return;
                    }
                }


                borderseq.Description = panel.Name;

                borderseq.PartQty = "1";
                borderseq.CutPartWidth = panel.ActualWidth;
                borderseq.CutPartLength = panel.ActualLength;
                borderseq.MaterialName = panel.Thickness + "mm" + panel.Material + panel.BasicMaterial;
                borderseq.MaterialCode = panel.MaterialId;

                if (panel.Edgelist.Count == 4)  // 20180525
                {
                    borderseq.Edgeband1 = panel.Edgelist[0].Thickness + "mm" + panel.Material.Replace("雏菊MM0261", "白山纹SW5001") + "封边条";               //应该要
                    borderseq.Edgeband2 = panel.Edgelist[1].Thickness + "mm" + panel.Material.Replace("雏菊MM0261", "白山纹SW5001") + "封边条";                                   //应该要
                    borderseq.Edgeband3 = panel.Edgelist[2].Thickness + "mm" + panel.Material.Replace("雏菊MM0261", "白山纹SW5001") + "封边条";                                     //应该要
                    borderseq.Edgeband4 = panel.Edgelist[3].Thickness + "mm" + panel.Material.Replace("雏菊MM0261", "白山纹SW5001") + "封边条";                                       //应该要

                }
                else
                {
                    borderseq.Edgeband1 = "";
                    borderseq.Edgeband2 = "";
                    borderseq.Edgeband3 = "";
                    borderseq.Edgeband4 = "";
                }

                borderseq.PartComments = "";
                borderseq.ProductDescription = "";
                borderseq.ProductQty = "";
                borderseq.ProductWidth = "";
                borderseq.ProductHeight = "";
                borderseq.ProductDepth = "";
                borderseq.ProductComments = "";
                borderseq.PerfectGrain = "";  
                borderseq.GrainFlag = "";
                borderseq.PartCounter = "";
                borderseq.FoundHdrill = "FALSE";                 //默认无，如果有则在每个水平 垂直 铣型赋值
                borderseq.FoundVdrill = "FALSE";                 //默认无，如果有则在每个水平 垂直 铣型赋值
                borderseq.FoundVdrillFace6 = "FALSE";            //默认无，如果有则在每个水平 垂直 铣型赋值
                borderseq.FoundRouting = "FALSE";                //默认无，如果有则在每个水平 垂直 铣型赋值
                borderseq.FoundRoutingFace6 = "FALSE";           //默认无，如果有则在每个水平 垂直 铣型赋值
                borderseq.FoundSawing = "FALSE";
                borderseq.FoundSawingFace6 = "FALSE";
                borderseq.FoundFace6Program = "FALSE";
                borderseq.FoundNesting = "FALSE";
                borderseq.FirstPassDepth = "1.5";
                borderseq.SpoilBoardPenetration = "0.1";
                borderseq.BasePoint = "";
                borderseq.MachinePoint = "1";            //需要依据镜像修改
                borderseq.MfgDataPath = "";
                borderseq.Col_EdgeFileNames1 = "";
                borderseq.Col_EdgeFileNames2 = "";
                borderseq.Col_EdgeFileNames3 = "";
                borderseq.Col_EdgeFileNames4 = "";
                borderseq.Col_EdgeBarCodes1 = "";
                borderseq.Col_EdgeBarCodes2 = "";
                borderseq.Col_EdgeBarCodes3 = "";
                borderseq.Col_EdgeBarCodes4 = "";
                borderseq.MaterialNameS = "";
                borderseq.Pointer = "";
                borderseq.PanelWidthS = panel.Width;
                borderseq.PanelLengthS = panel.Length;
                borderseq.TheDataOrigin = "";
                borderseq.ReleaseFolder = "UnNamed";
                borderseq.VHolesCount = "";
                borderseq.HHolesCount = "";
                borderseq.RoutesCount = "";
                borderseq.SawsCount = "0";  //记录面5 槽类型等  20180807
                borderseq.NestRoutesCount = ""; 

                point4.Add(new fourpoint(0, 0));
                point4.Add(new fourpoint(double.Parse(borderseq.PanelLength), 0));
                point4.Add(new fourpoint(double.Parse(borderseq.PanelLength), double.Parse(borderseq.PanelWidth)));
                point4.Add(new fourpoint(0, double.Parse(borderseq.PanelWidth)));


                for (int i = 0; i < panel.Machininglist.Count;i++ )
                {                          
                    if (panel.Machininglist[i].Type == "1")
                    {
                        HDrillSequenceEntity hdrillseq = new HDrillSequenceEntity();
                        hdrillseq.HDrillSequence = "HDrillSequence";

                        if (panel.Machininglist[i].Face == "1")
                        {
                            hdrillseq.CurrentFace = "2";
                            hdrillseq.HDrillX = panel.Machininglist[i].X;
                            hdrillseq.HDrillZ = panel.Machininglist[i].Depth;
                            hdrillseq.HDrillY = panel.Machininglist[i].Z;
                        }

                        else if (panel.Machininglist[i].Face == "2")
                        {
                            hdrillseq.CurrentFace = "1";
                            hdrillseq.HDrillX = panel.Machininglist[i].X;
                            hdrillseq.HDrillZ = panel.Machininglist[i].Depth;
                            hdrillseq.HDrillY = panel.Machininglist[i].Z;
                        }
                        else if (panel.Machininglist[i].Face == "3")
                        {
                            hdrillseq.CurrentFace = "4";
                            hdrillseq.HDrillX = panel.Machininglist[i].Depth;
                            hdrillseq.HDrillZ = panel.Machininglist[i].Y;
                            hdrillseq.HDrillY = panel.Machininglist[i].Z;

                        }
                        else if (panel.Machininglist[i].Face == "4")
                        {
                            hdrillseq.CurrentFace = "3";
                            hdrillseq.HDrillX = panel.Machininglist[i].Depth;
                            hdrillseq.HDrillZ = panel.Machininglist[i].Y;
                            hdrillseq.HDrillY = panel.Machininglist[i].Z;
                        }

                        if  (panel.Machininglist[i].Face != "0")  //20180816 发现水平孔有面为0的情况，将之面为0的情况过滤
                        {
                            hdrillseq.PreviousFace = "";
                            hdrillseq.HDrillDiameter = panel.Machininglist[i].Diameter;
                            hdrillseq.HDrillToolName = "";
                            hdrillseq.HDrillFeedSpeed = "3000";
                            hdrillseq.HDrillEntrySpeed = "2700";
                            hdrillseq.HDrillFirstDrillDone = "";
                            hdrillseq.HDrillPreviousToolName = "";
                            hdrillseq.HDrilleNextToolName = "";
                            hdrillseq.HDrillCounter = "";
                            borderseq.FoundHdrill = "TRUE";

                            face5list.Add(hdrillseq.OutPutCsvString());
                        }
                     
                    }
                    else if (panel.Machininglist[i].Type == "2")
                    {
                        VdrillSequenceEntity vdrillseq = new VdrillSequenceEntity();
                        vdrillseq.VdrillSequence = "VdrillSequence";
                        vdrillseq.VDrillX = panel.Machininglist[i].X;
                        vdrillseq.VDrillY = panel.Machininglist[i].Y;
                        vdrillseq.VDrillZ = panel.Machininglist[i].Depth;
                        vdrillseq.VDrillXOffset = "0";
                        vdrillseq.VDrillYOffset = "0";
                        vdrillseq.VDrillDiameter = panel.Machininglist[i].Diameter;
                        vdrillseq.VDrillToolName = "";
                        vdrillseq.VDrillFeedSpeed = "3000";
                        vdrillseq.VDrillEntrySpeed = "8000";
                        vdrillseq.VDrillBitType = "";
                        vdrillseq.VDrillFirstDrillDone = "";
                        vdrillseq.VDrillPreviousToolName = "";
                        vdrillseq.VDrillCounter = "";

                        if (vdrillseq.VDrillDiameter != "") //20180828 发现垂直孔的直径出现为空的情况，将此情况过滤
                        {
                            #region 垂直孔是25的，将25的孔先用20的孔打掉。然后再用10mm的刀铣成25的直径
                            if (Math.Abs(Convert.ToDouble(vdrillseq.VDrillDiameter) - 25) < 0.41)
                            {
                                vdrillseq.VDrillDiameter = "20";
                                vdrillseq.VDrillZ = (Convert.ToDouble(vdrillseq.VDrillZ) + 0.15).ToString();
                                RouteSetMillSequenceEntity route = new RouteSetMillSequenceEntity();

                                route = RouteProcess(panel.Machininglist[i], 1);

                                if (panel.Machininglist[i].Face == "5")
                                {
                                    borderseq.FoundVdrillFace6 = "TRUE";
                                    borderseq.FoundRoutingFace6 = "TRUE";
                                    face6list.Add(vdrillseq.OutPutCsvString());
                                    face6list.Add(route.OutPutCsvString());
                                }
                                else if (panel.Machininglist[i].Face == "6")
                                {
                                    borderseq.FoundVdrill = "TRUE";
                                    borderseq.FoundRouting = "TRUE";
                                    face5list.Add(vdrillseq.OutPutCsvString());
                                    face5list.Add(route.OutPutCsvString());
                                }

                                route = RouteProcess(panel.Machininglist[i], 2);

                                if (panel.Machininglist[i].Face == "5")
                                {
                                    borderseq.FoundRoutingFace6 = "TRUE";
                                    face6list.Add(route.OutPutCsvString());
                                }
                                else if (panel.Machininglist[i].Face == "6")
                                {
                                    borderseq.FoundRouting = "TRUE";
                                    face5list.Add(route.OutPutCsvString());
                                }

                                route = RouteProcess(panel.Machininglist[i], 3);

                                if (panel.Machininglist[i].Face == "5")
                                {
                                    borderseq.FoundRoutingFace6 = "TRUE";
                                    face6list.Add(route.OutPutCsvString());
                                }
                                else if (panel.Machininglist[i].Face == "6")
                                {
                                    borderseq.FoundRouting = "TRUE";
                                    face5list.Add(route.OutPutCsvString());
                                }
                            }
                            else
                            {
                                if (panel.Machininglist[i].Face == "5")
                                {
                                    borderseq.FoundVdrillFace6 = "TRUE";
                                    face6list.Add(vdrillseq.OutPutCsvString());
                                }
                                else if (panel.Machininglist[i].Face == "6")
                                {
                                    borderseq.FoundVdrill = "TRUE";
                                    face5list.Add(vdrillseq.OutPutCsvString());
                                }
                            }
                            #endregion

                            if (Math.Abs(double.Parse(vdrillseq.VDrillDiameter) - 20) < 0.1 && panel.Name.Contains("二合一层板"))  //20180611  宋新刚 对于二合一层板，需要将垂直孔的部份，用水平孔将封边条给打掉
                            {
                                bool iscorrect = false;
                                HDrillSequenceEntity hdrillseq = new HDrillSequenceEntity();

                                hdrillseq.HDrillSequence = "HDrillSequence";
                                if (Math.Abs(double.Parse(vdrillseq.VDrillX) - 9.5) < 0.51)  //20180613  宋新刚 因为二合一层板时常加成9.5 又时常加成9 所以在这里判断的时候，索性写成9和9.5都认
                                {
                                    hdrillseq.CurrentFace = "3";
                                    hdrillseq.HDrillX = "8";
                                    hdrillseq.HDrillZ = vdrillseq.VDrillY;
                                    iscorrect = true;
                                }
                                else if (Math.Abs(double.Parse(panel.Length) - double.Parse(vdrillseq.VDrillX) - 9.5) < 0.51)
                                {
                                    hdrillseq.CurrentFace = "4";
                                    hdrillseq.HDrillX = "8";
                                    hdrillseq.HDrillZ = vdrillseq.VDrillY;
                                    iscorrect = true;
                                }
                                else if (Math.Abs(double.Parse(vdrillseq.VDrillY) - 9.5) < 0.51)  //发现二合一的层板有三面 故增加三面的数据
                                {
                                    hdrillseq.CurrentFace = "1";
                                    hdrillseq.HDrillX = vdrillseq.VDrillX;
                                    hdrillseq.HDrillZ = "8";
                                    iscorrect = true;
                                }
                                else if (Math.Abs(double.Parse(panel.Width) - double.Parse(vdrillseq.VDrillY) - 9.5) < 0.51)
                                {
                                    hdrillseq.CurrentFace = "2";
                                    hdrillseq.HDrillX = vdrillseq.VDrillX;
                                    hdrillseq.HDrillZ = "8";
                                    iscorrect = true;
                                }

                                if (iscorrect)   //20180627  发现不在四个大面的二合一垂直孔 打二合一的时候，就不产生水平孔
                                {
                                    hdrillseq.PreviousFace = "";
                                    hdrillseq.HDrillY = "9";
                                    hdrillseq.HDrillDiameter = "8";
                                    hdrillseq.HDrillToolName = "";
                                    hdrillseq.HDrillFeedSpeed = "3000";
                                    hdrillseq.HDrillEntrySpeed = "2700";
                                    hdrillseq.HDrillFirstDrillDone = "";
                                    hdrillseq.HDrillPreviousToolName = "";
                                    hdrillseq.HDrilleNextToolName = "";
                                    hdrillseq.HDrillCounter = "";
                                    borderseq.FoundHdrill = "TRUE";

                                    face5list.Add(hdrillseq.OutPutCsvString());
                                }

                            }
                        }
                    }
                    else if (panel.Machininglist[i].Type == "3")
                    {
                        //if (panel.ID == "2513311050")
                        //    MessageBox.Show("他奶奶的!");

                        if (panel.Machininglist[i].ToolName == "开料铣型刀" && panel.Machininglist[i].GrooveType == "1" && panel.Machininglist[i].Linelist.Count == 4 && Math.Abs(float.Parse(panel.Machininglist[i].Linelist[1].EndY) - float.Parse(panel.Machininglist[i].Y)) != 0 && Math.Abs(float.Parse(panel.Machininglist[i].Linelist[1].EndX) - float.Parse(panel.Machininglist[i].X)) != 0) //目前只考虑矩形槽的情况 且踢除掉4段线的圆 20180331  //20180822 增加了X方向点的加入 不然4段线的圆弧踢不掉
                        {
                            double min_x = double.Parse(panel.Machininglist[i].X);
                            double min_y = double.Parse(panel.Machininglist[i].Y);
                            double max_x = double.Parse(panel.Machininglist[i].X);
                            double max_y = double.Parse(panel.Machininglist[i].Y);
                            double depth = double.Parse(panel.Machininglist[i].Depth);

                            for (int j = 0; j < panel.Machininglist[i].Linelist.Count; j++)
                            {
                                min_x = Math.Min(min_x, double.Parse(panel.Machininglist[i].Linelist[j].EndX));
                                min_y = Math.Min(min_y, double.Parse(panel.Machininglist[i].Linelist[j].EndY));
                                max_x = Math.Max(max_x, double.Parse(panel.Machininglist[i].Linelist[j].EndX));
                                max_y = Math.Max(max_y, double.Parse(panel.Machininglist[i].Linelist[j].EndY));
                            }

                            double width = max_y - min_y;
                            double height = max_x - min_x;

                            if (width > height)   // 发现有铣槽有横向和纵向的区别  20180419
                            {
                                bool useD3after129 = true; //20180620 默认值改为true
                                if (Math.Abs(height - 41.5) < 0.1 && Math.Abs(width - 112.5) < 0.1 && Math.Abs(depth - 13) < 0.1) //帕码隐藏长拉手
                                    useD3after129 = true;
                                else if (Math.Abs(height - 48) < 0.1 && Math.Abs(width - 198.2) < 0.1 && Math.Abs(depth - 14) < 0.1) //铁灰帕码内嵌拉手
                                    useD3after129 = true;
                                else if ((Math.Abs(height - 41.5) < 0.1 && Math.Abs(width - 41.5) < 0.1 && Math.Abs(depth - 13) < 0.1))//帕码隐藏方拉手
                                    useD3after129 = true;
                                else if ((Math.Abs(height - 26.3) < 0.1 && Math.Abs(width - 55) < 0.1 && Math.Abs(depth - 14) < 0.1))//意大利铰链
                                    useD3after129 = false;  //20180620 意大利铝框灰玻门铰 不需要D3刀绕圈  51改到55是因为要多切出去4mm

                                RouteSetMillSequenceEntity routefirst = null;
                                RouteSetMillSequenceEntity routesecond = null;

                                if (height < 6) // 取直径为3的double 来做
                                {
                                    double everycuttingdepth = 5;
                                    routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y);
                                    double needcuttingdepth = Convert.ToDouble(routefirst.RouteZ);
                                    double cutingtimes = Math.Ceiling(needcuttingdepth / everycuttingdepth);
                                    if (height >= 3)  //槽宽大于3 小于6 铣一圈 每圈铣的深度为 Z/3 
                                    {
                                        if (double.Parse(panel.Machininglist[i].Depth) > everycuttingdepth)  //对直径为3mm的刀具做分层处理
                                        {
                                            for (int l = 1; l <= cutingtimes; l++)  //分层循环
                                            {
                                                if (height == 3)
                                                {
                                                    routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y - 3/2);

                                                    routefirst.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }  //输出铣型的开始部份  

                                                    routesecond = RouteProcess(panel.Machininglist[i], height, min_x, min_y + 3/2);
                                                    routesecond.RouteSetMillSequence = "RouteSequence";

                                                    routesecond.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routesecond.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routesecond.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routesecond.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }
                                                }
                                                else
                                                {
                                                    double leadin_x = min_x + 3 + (height - 3) / 2;   //20180420
                                                    double leadin_y = max_y - 10;   //20180418

                                                    if (Math.Abs(height - double.Parse(panel.Length)) < 0.1)  //如果铣型发现在X方向铣通的话，则算刀具轨迹的时候多铣5mm 以将因刀产生的圆弧部份铣掉 20180414
                                                    {
                                                        min_y = min_y - 1.5;
                                                        max_y = max_y + 1.5;
                                                    }

                                                    routefirst = RouteProcess(panel.Machininglist[i], height, leadin_x, leadin_y);

                                                    routefirst.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteDiameter = "3";
                                                    routefirst.RouteToolName = "D3";

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }  //输出引线的开始部份  

                                                    routefirst = RouteProcess(panel.Machininglist[i], height, leadin_x, max_y);
                                                    routefirst.RouteSetMillSequence = "RouteSequence";
                                                    routefirst.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteDiameter = "3";
                                                    routefirst.RouteToolName = "D3";

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }  //输出引线的结束部份  


                                                    routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y);
                                                    routefirst.RouteSetMillSequence = "RouteSequence";
                                                    routefirst.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteDiameter = "3";
                                                    routefirst.RouteToolName = "D3";

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }  //输出XML第一个点的开始部份  

                                                    //D3的刀走一圈未测试~！

                                                    for (int j = 0; j < panel.Machininglist[i].Linelist.Count; j++)  //输出铣型部份
                                                    {
                                                        if (j == 0)
                                                            routesecond = RouteProcess(panel.Machininglist[i], height, min_x, min_y);
                                                        else if (j == 1)
                                                            routesecond = RouteProcess(panel.Machininglist[i], height, max_x, min_y);
                                                        else if (j == 2)
                                                            routesecond = RouteProcess(panel.Machininglist[i], height, max_x, max_y);
                                                        else if (j == 3)
                                                            routesecond = RouteProcess(panel.Machininglist[i], height, leadin_x - 3/2, max_y);   //20180420

                                                        routesecond.RouteSetMillSequence = "RouteSequence";
                                                        routesecond.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                        routesecond.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                        routesecond.RouteDiameter = "3";
                                                        routesecond.RouteToolName = "D3";

                                                        if (panel.Machininglist[i].Face == "5")
                                                        {
                                                            face6list.Add(routesecond.OutPutCsvString());
                                                            borderseq.FoundRoutingFace6 = "TRUE";
                                                        }
                                                        else if (panel.Machininglist[i].Face == "6")
                                                        {
                                                            face5list.Add(routesecond.OutPutCsvString());
                                                            borderseq.FoundRouting = "TRUE";
                                                        }
                                                    }
                                                }
                                            }

                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("槽宽为:  " + height.ToString() + "不在有其 3 - 8.5mm的范围之内");
                                        return;
                                    }

                                }
                                else if (height >= 8.5)
                                {
                                    double everycuttingwidth = 8;
                                    double cutingtimes = Math.Ceiling(height / everycuttingwidth);
                                    double needcuttingdepth = double.Parse(panel.Machininglist[i].Depth); // 定义由槽打散成铣型部份的深度。如果深度不大于板厚。则走一字型  20180330                             
                                    if (needcuttingdepth < double.Parse(panel.Thickness) && height > 16)  //如果槽宽小于16 依然一圈绕下来
                                    {
                                        for (int l = 0; l < cutingtimes; l++)  //分段循环
                                        {
                                            if (useD3after129)//20180620 宋新刚 意大利铰链 因为刀路需要从外向内。所以这里做了特殊处理 
                                                routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y - 8.5/2);
                                            else
                                            {
                                                routefirst = RouteProcess(panel.Machininglist[i], height, min_x, min_y + 8.5 / 2);

                                                if (routefirst.RouteToolComp == "1")  //发现转换刀路轨迹时偏置也是需要对调的 20180709
                                                    routefirst.RouteToolComp = "2";
                                                else if (routefirst.RouteToolComp == "2")
                                                    routefirst.RouteToolComp = "1";
                                            }


                                            if (height - everycuttingwidth * l < everycuttingwidth)
                                            {
                                                routefirst.RouteX = (min_x + height - 8.5).ToString();  //8.5毫米是刀具直径
                                                routefirst.RouteSetMillX = (min_x + height - 8.5).ToString();
                                            }
                                            else
                                            {
                                                routefirst.RouteX = (min_x + everycuttingwidth * l).ToString();
                                                routefirst.RouteSetMillX = (min_x + everycuttingwidth * l).ToString();
                                            }

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }  //输出铣型的开始部份  

                                            if (useD3after129)//20180620 宋新刚 意大利铰链 因为刀路需要从外向内。所以这里做了特殊处理 
                                                routesecond = RouteProcess(panel.Machininglist[i], height, min_x, min_y + 8.5/2);
                                            else
                                            {
                                                routesecond = RouteProcess(panel.Machininglist[i], height, min_x, max_y - 8.5 / 2);

                                                if (routesecond.RouteToolComp == "1")  //发现转换刀路轨迹时偏置也是需要对调的 20180709
                                                    routesecond.RouteToolComp = "2";
                                                else if (routesecond.RouteToolComp == "2")
                                                    routesecond.RouteToolComp = "1";
                                            }


                                            routesecond.RouteSetMillSequence = "RouteSequence";

                                            if (height - everycuttingwidth * l < everycuttingwidth)
                                            {
                                                routesecond.RouteX = (min_x + height - 8.5).ToString();  //8.5毫米是刀具直径
                                                routesecond.RouteSetMillX = (min_x + height - 8.5).ToString();
                                            }
                                            else
                                            {
                                                routesecond.RouteX = (min_x + everycuttingwidth * l).ToString();
                                                routesecond.RouteSetMillX = (min_x + everycuttingwidth * l).ToString();
                                            }

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routesecond.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routesecond.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }
                                        }
                                        //useD3after129 = true;  //铣灯槽的时候发现拉三刀 还是需要D3的刀铣一圈的  所以这边先不做是不是帕码拉手处理。只要铣的深度小于板厚 槽宽大于16的都用D3的刀铣一下  20180409

                                        if (useD3after129)
                                        {
                                            double everycuttingdepth1 = 5;
                                            routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y);
                                            double needcuttingdepth1 = Convert.ToDouble(routefirst.RouteZ);
                                            double cutingtimes1 = Math.Ceiling(needcuttingdepth1 / everycuttingdepth1);

                                            for (int l = 1; l <= cutingtimes1; l++)  //分层循环
                                            {
                                                double leadin_x = min_x + 3 + (height - 3) / 2;   //20180420
                                                double leadin_y = max_y - 10;   //20180418

                                                if (Math.Abs(height - double.Parse(panel.Length)) < 0.1)  //如果铣型发现在X方向铣通的话，则算刀具轨迹的时候多铣5mm 以将因刀产生的圆弧部份铣掉 20180414
                                                {
                                                    min_y = min_y - 1.5;
                                                    max_y = max_y + 1.5;
                                                }

                                                routefirst = RouteProcess(panel.Machininglist[i], height, leadin_x, leadin_y);

                                                routefirst.RouteZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteSetMillZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteDiameter = "3";
                                                routefirst.RouteToolName = "D3";

                                                if (panel.Machininglist[i].Face == "5")
                                                {
                                                    face6list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRoutingFace6 = "TRUE";
                                                }
                                                else if (panel.Machininglist[i].Face == "6")
                                                {
                                                    face5list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRouting = "TRUE";
                                                }  //输出引线的开始部份  

                                                routefirst = RouteProcess(panel.Machininglist[i], height, leadin_x, max_y);
                                                routefirst.RouteSetMillSequence = "RouteSequence";
                                                routefirst.RouteZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteSetMillZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteDiameter = "3";
                                                routefirst.RouteToolName = "D3";
                                                if (panel.Machininglist[i].Face == "5")
                                                {
                                                    face6list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRoutingFace6 = "TRUE";
                                                }
                                                else if (panel.Machininglist[i].Face == "6")
                                                {
                                                    face5list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRouting = "TRUE";
                                                }  //输出引线的结束部份  


                                                routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y);
                                                routefirst.RouteSetMillSequence = "RouteSequence";
                                                routefirst.RouteZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteSetMillZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteDiameter = "3";
                                                routefirst.RouteToolName = "D3";
                                                if (panel.Machininglist[i].Face == "5")
                                                {
                                                    face6list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRoutingFace6 = "TRUE";
                                                }
                                                else if (panel.Machininglist[i].Face == "6")
                                                {
                                                    face5list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRouting = "TRUE";
                                                }  //输出XML第一个点的开始部份  

                                                for (int j = 0; j < panel.Machininglist[i].Linelist.Count; j++)  //输出铣型部份
                                                {
                                                    if (j == 0)
                                                        routesecond = RouteProcess(panel.Machininglist[i], height, min_x, min_y);
                                                    else if (j == 1)
                                                        routesecond = RouteProcess(panel.Machininglist[i], height, max_x, min_y);
                                                    else if (j == 2)
                                                        routesecond = RouteProcess(panel.Machininglist[i], height, max_x, max_y);
                                                    else if (j == 3)
                                                        routesecond = RouteProcess(panel.Machininglist[i], height, leadin_x - 3/2, max_y);   //20180420

                                                    routesecond.RouteSetMillSequence = "RouteSequence";
                                                    routesecond.RouteZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                    routesecond.RouteSetMillZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                    routesecond.RouteDiameter = "3";
                                                    routesecond.RouteToolName = "D3";

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routesecond.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routesecond.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (height == 8.5)
                                        {
                                            routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y + 8.5/2);

                                            routefirst.RouteSetMillSequence = "RouteSetMillSequence";   //20180615 宋新刚 护墙板 单独8.5的矩形槽  这里原本错误
                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }  //输出铣型的开始部份 

                                            routesecond = RouteProcess(panel.Machininglist[i], height, min_x, min_y - 8.5 / 2);

                                            routesecond.RouteSetMillSequence = "RouteSequence";

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routesecond.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routesecond.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }
                                        }
                                        else
                                        {   //这边的情况有如下几点：
                                            //1:铣的深度大于等于板厚
                                            //2:铣的宽度大于16
                                            //3：将原先的用的8.5的刀具替换成10的刀具。因为发现有25mm板厚的情况
                                            //4:将原先的所有的8.5的字符都替换成10

                                            double leadin_x = min_x + 10 + (height - 10) / 2;   //20180420
                                            double leadin_y = max_y - 30;   //20180418

                                            if (Math.Abs(height - double.Parse(panel.Length)) < 0.1)  //如果铣型发现在X方向铣通的话，则算刀具轨迹的时候多铣5mm 以将因刀产生的圆弧部份铣掉 20180414
                                            {
                                                min_y = min_y - 4;
                                                max_y = max_y + 4;
                                            }

                                            routefirst = RouteProcess(panel.Machininglist[i], height, leadin_x, leadin_y);

                                            routefirst.RouteDiameter = "10"; // 20180725 将刀具从8.5改成10
                                            routefirst.RouteToolName = "130";

                                            if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                            {
                                                routefirst.RouteEndTangentY = "1";

                                                if (Math.Abs(height - 15.1) < 0.1 || Math.Abs(height - 15.5) < 0.1)  //20180817 汉森专用滑门槽，在NESTING机床过滤。在PTP160上加工
                                                {
                                                    routefirst.RouteDiameter = "10.01";
                                                    routefirst.RouteToolName = "D130";
                                                }
                                            }

                                            if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                            {
                                                routefirst.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                                routefirst.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                            }

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }  //输出引线的开始部份 

                                            routefirst = RouteProcess(panel.Machininglist[i], height, leadin_x, max_y);
                                            routefirst.RouteSetMillSequence = "RouteSequence";
                                            routefirst.RouteDiameter = "10";  //20180725 替换的刀具
                                            routefirst.RouteToolName = "130";

                                            if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                            {
                                                routefirst.RouteEndTangentY = "1";

                                                if (Math.Abs(height - 15.1) < 0.1 || Math.Abs(height - 15.5) < 0.1)  //20180817 汉森专用滑门槽，在NESTING机床过滤。在PTP160上加工
                                                {
                                                    routefirst.RouteDiameter = "10.01";
                                                    routefirst.RouteToolName = "D130";
                                                }

                                            }

                                            if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                            {
                                                routefirst.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                                routefirst.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                            }

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }  //输出引线的结束部份  


                                            routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y);
                                            routefirst.RouteSetMillSequence = "RouteSequence";
                                            routefirst.RouteDiameter = "10";  //20180725 替换的刀具
                                            routefirst.RouteToolName = "130";

                                            if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                            {
                                                routefirst.RouteEndTangentY = "1";

                                                if (Math.Abs(height - 15.1) < 0.1 || Math.Abs(height - 15.5) < 0.1)  //20180817 汉森专用滑门槽，在NESTING机床过滤。在PTP160上加工
                                                {
                                                    routefirst.RouteDiameter = "10.01";
                                                    routefirst.RouteToolName = "D130";
                                                }
                                            }

                                            if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                            {
                                                routefirst.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                                routefirst.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                            }

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }  //输出XML第一个点的开始部份  

                                            for (int j = 0; j < panel.Machininglist[i].Linelist.Count; j++)  //输出铣型部份
                                            {
                                                if (j == 0)
                                                    routesecond = RouteProcess(panel.Machininglist[i], height, min_x, min_y);
                                                else if (j == 1)
                                                    routesecond = RouteProcess(panel.Machininglist[i], height, max_x, min_y);
                                                else if (j == 2)
                                                    routesecond = RouteProcess(panel.Machininglist[i], height, max_x, max_y);
                                                else if (j == 3)
                                                    routesecond = RouteProcess(panel.Machininglist[i], height, leadin_x - 10/2, max_y);   //20180420 //20180818 将这里的8.5改成10

                                                routesecond.RouteSetMillSequence = "RouteSequence";

                                                routesecond.RouteDiameter = "10"; // 20180725 将刀具从8.5改成10
                                                routesecond.RouteToolName = "130";

                                                if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                                {
                                                    routesecond.RouteEndTangentY = "1";

                                                    if (Math.Abs(height - 15.1) < 0.1 || Math.Abs(height - 15.5) < 0.1)  //20180817 汉森专用滑门槽，在NESTING机床过滤。在PTP160上加工
                                                    {
                                                        routesecond.RouteDiameter = "10.01";
                                                        routesecond.RouteToolName = "D130";
                                                    }
                                                }

                                                if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                                {
                                                    routesecond.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                                    routesecond.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                                }

                                                if (panel.Machininglist[i].Face == "5")
                                                {
                                                    face6list.Add(routesecond.OutPutCsvString());
                                                    borderseq.FoundRoutingFace6 = "TRUE";
                                                }
                                                else if (panel.Machininglist[i].Face == "6")
                                                {
                                                    face5list.Add(routesecond.OutPutCsvString());
                                                    borderseq.FoundRouting = "TRUE";
                                                }
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Height 槽宽为:  " + height.ToString() + "没有合适的刀具进行加工!");
                                    return;
                                }
                            }
                            else
                            {

                                bool useD3after129 = true;  //20180620 默认值改为true
                                if (Math.Abs(width - 41.5) < 0.1 && Math.Abs(height - 112.5) < 0.1 && Math.Abs(depth - 13) < 0.1) //帕码隐藏长拉手
                                    useD3after129 = true;
                                else if (Math.Abs(width - 48) < 0.1 && Math.Abs(height - 198.2) < 0.1 && Math.Abs(depth - 14) < 0.1) //铁灰帕码内嵌拉手
                                    useD3after129 = true;
                                else if ((Math.Abs(width - 41.5) < 0.1 && Math.Abs(height - 41.5) < 0.1 && Math.Abs(depth - 13) < 0.1))//帕码隐藏方拉手
                                    useD3after129 = true;
                                else if ((Math.Abs(width - 26.3) < 0.1 && Math.Abs(height - 55) < 0.1 && Math.Abs(depth - 14) < 0.1))//意大利铰链
                                    useD3after129 = false;  //20180620 意大利铝框灰玻门铰 不需要D3刀绕圈  51改到55是因为要多切出去4mm

                                RouteSetMillSequenceEntity routefirst = null;
                                RouteSetMillSequenceEntity routesecond = null;

                                if (width < 6) // 取直径为3的double 来做
                                {
                                    double everycuttingdepth = 5;
                                    routefirst = RouteProcess(panel.Machininglist[i], width, min_x, min_y);
                                    double needcuttingdepth = Convert.ToDouble(routefirst.RouteZ);
                                    double cutingtimes = Math.Ceiling(needcuttingdepth / everycuttingdepth);
                                    if (width >= 3)  //槽宽大于3 小于6 铣一圈 每圈铣的深度为 Z/3 
                                    {
                                        if (double.Parse(panel.Machininglist[i].Depth) > everycuttingdepth)  //对直径为3mm的刀具做分层处理
                                        {
                                            for (int l = 1; l <= cutingtimes; l++)  //分层循环
                                            {
                                                if (width == 3)
                                                {
                                                    routefirst = RouteProcess(panel.Machininglist[i], width, min_x + 3/2, min_y);

                                                    routefirst.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }  //输出铣型的开始部份  

                                                    routesecond = RouteProcess(panel.Machininglist[i], width, max_x - 3/2, min_y);
                                                    routesecond.RouteSetMillSequence = "RouteSequence";

                                                    routesecond.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routesecond.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routesecond.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routesecond.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }
                                                }
                                                else
                                                {
                                                    double leadin_x = min_x + 10;
                                                    double leadin_y = min_y + 3 + (width - 3) / 2;

                                                    if (Math.Abs(height - double.Parse(panel.Length)) < 0.1)  //如果铣型发现在X方向铣通的话，则算刀具轨迹的时候多铣5mm 以将因刀产生的圆弧部份铣掉 20180414
                                                    {
                                                        min_x = min_x - 1.5;
                                                        max_x = max_x + 1.5;
                                                    }

                                                    routefirst = RouteProcess(panel.Machininglist[i], width, leadin_x, leadin_y);

                                                    routefirst.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteDiameter = "3";
                                                    routefirst.RouteToolName = "D3";

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }  //输出引线的开始部份  

                                                    routefirst = RouteProcess(panel.Machininglist[i], width, min_x, leadin_y);
                                                    routefirst.RouteSetMillSequence = "RouteSequence";
                                                    routefirst.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteDiameter = "3";
                                                    routefirst.RouteToolName = "D3";

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }  //输出引线的结束部份  


                                                    routefirst = RouteProcess(panel.Machininglist[i], width, min_x, min_y);
                                                    routefirst.RouteSetMillSequence = "RouteSequence";
                                                    routefirst.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                    routefirst.RouteDiameter = "3";
                                                    routefirst.RouteToolName = "D3";

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routefirst.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }  //输出XML第一个点的开始部份  
  
                                                    //D3的刀走一圈未测试~！

                                                    for (int j = 0; j < panel.Machininglist[i].Linelist.Count; j++)  //输出铣型部份
                                                    {
                                                        if (j == 0)
                                                            routesecond = RouteProcess(panel.Machininglist[i], width, max_x, min_y);
                                                        else if (j == 1)
                                                            routesecond = RouteProcess(panel.Machininglist[i], width, max_x, max_y);
                                                        else if (j == 2)
                                                            routesecond = RouteProcess(panel.Machininglist[i], width, min_x, max_y);
                                                        else if (j == 3)
                                                            routesecond = RouteProcess(panel.Machininglist[i], width, min_x, leadin_y - 3/2);

                                                        routesecond.RouteSetMillSequence = "RouteSequence";

                                                        routesecond.RouteZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                        routesecond.RouteSetMillZ = (needcuttingdepth - everycuttingdepth * (cutingtimes - l)).ToString();
                                                        routesecond.RouteDiameter = "3";
                                                        routesecond.RouteToolName = "D3";

                                                        if (panel.Machininglist[i].Face == "5")
                                                        {
                                                            face6list.Add(routesecond.OutPutCsvString());
                                                            borderseq.FoundRoutingFace6 = "TRUE";
                                                        }
                                                        else if (panel.Machininglist[i].Face == "6")
                                                        {
                                                            face5list.Add(routesecond.OutPutCsvString());
                                                            borderseq.FoundRouting = "TRUE";
                                                        }
                                                    }
                                                }
                                            }

                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("槽宽为:  " + width.ToString() + "不在有其 3 - 8.5mm的范围之内");
                                        return;
                                    }
                                }
                                else if (width >= 8.5)
                                {
                                    double everycuttingwidth = 8;
                                    double cutingtimes = Math.Ceiling(width / everycuttingwidth);
                                    double needcuttingdepth = double.Parse(panel.Machininglist[i].Depth); // 定义由槽打散成铣型部份的深度。如果深度不大于板厚。则走一字型  20180330                             
                                    if (needcuttingdepth < double.Parse(panel.Thickness) && width > 16)  //如果槽宽小于16 依然一圈绕下来
                                    {
                                        for (int l = 0; l < cutingtimes; l++)  //分段循环
                                        {
                                            if (useD3after129) //20180620 宋新刚 意大利铰链 因为刀路需要从外向内。所以这里做了特殊处理 
                                                routefirst = RouteProcess(panel.Machininglist[i], width, min_x + 8.5 / 2, min_y);
                                            else
                                            {
                                                routefirst = RouteProcess(panel.Machininglist[i], width, max_x - 8.5 / 2, min_y);

                                                if (routefirst.RouteToolComp == "1")  //发现转换刀路轨迹时偏置也是需要对调的 20180709
                                                    routefirst.RouteToolComp = "2";
                                                else if (routefirst.RouteToolComp == "2")
                                                    routefirst.RouteToolComp = "1";
                                            }
                                                

                                            if (width - everycuttingwidth * l < everycuttingwidth)
                                            {
                                                routefirst.RouteY = (min_y + width - 8.5).ToString();  //8.5毫米是刀具直径
                                                routefirst.RouteSetMillY = (min_y + width - 8.5).ToString();
                                            }
                                            else
                                            {
                                                routefirst.RouteY = (min_y + everycuttingwidth * l).ToString();
                                                routefirst.RouteSetMillY = (min_y + everycuttingwidth * l).ToString();
                                            }

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }  //输出铣型的开始部份  

                                            if (useD3after129)//20180620 宋新刚 意大利铰链 因为刀路需要从外向内。所以这里做了特殊处理 
                                                routesecond = RouteProcess(panel.Machininglist[i], width, max_x - 8.5/2, min_y);
                                            else
                                            {
                                                routesecond = RouteProcess(panel.Machininglist[i], width, min_x + 8.5 / 2, min_y);

                                                if (routesecond.RouteToolComp == "1")  //发现转换刀路轨迹时偏置也是需要对调的 20180709
                                                    routesecond.RouteToolComp = "2";
                                                else if (routesecond.RouteToolComp == "2")
                                                    routesecond.RouteToolComp = "1";
                                            }


                                            routesecond.RouteSetMillSequence = "RouteSequence";
                                            if (width - everycuttingwidth * l < everycuttingwidth)
                                            {
                                                routesecond.RouteY = (min_y + width - 8.5).ToString();  //8.5毫米是刀具直径
                                                routesecond.RouteSetMillY = (min_y + width - 8.5).ToString();
                                            }
                                            else
                                            {
                                                routesecond.RouteY = (min_y + everycuttingwidth * l).ToString();
                                                routesecond.RouteSetMillY = (min_y + everycuttingwidth * l).ToString();
                                            }

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routesecond.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routesecond.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }
                                        }
                                        //useD3after129 = true;  //铣灯槽的时候发现拉三刀 还是需要D3的刀铣一圈的  所以这边先不做是不是帕码拉手处理。只要铣的深度小于板厚 槽宽大于16的都用D3的刀铣一下  20180409
                                        if (useD3after129)
                                        {
                                            double everycuttingdepth1 = 5;
                                            routefirst = RouteProcess(panel.Machininglist[i], width, min_x, min_y);
                                            double needcuttingdepth1 = Convert.ToDouble(routefirst.RouteZ);
                                            double cutingtimes1 = Math.Ceiling(needcuttingdepth1 / everycuttingdepth1);

                                            for (int l = 1; l <= cutingtimes1; l++)  //分层循环
                                            {
                                                double leadin_x = min_x + 10;
                                                double leadin_y = min_y + 3 + (width - 3) / 2;

                                                if (Math.Abs(height - double.Parse(panel.Length)) < 0.1)  //如果铣型发现在X方向铣通的话，则算刀具轨迹的时候多铣5mm 以将因刀产生的圆弧部份铣掉 20180414
                                                {
                                                    min_x = min_x - 1.5;
                                                    max_x = max_x + 1.5;
                                                }

                                                routefirst = RouteProcess(panel.Machininglist[i], width, leadin_x, leadin_y);

                                                routefirst.RouteZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteSetMillZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteDiameter = "3";
                                                routefirst.RouteToolName = "D3";

                                                if (panel.Machininglist[i].Face == "5")
                                                {
                                                    face6list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRoutingFace6 = "TRUE";
                                                }
                                                else if (panel.Machininglist[i].Face == "6")
                                                {
                                                    face5list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRouting = "TRUE";
                                                }  //输出引线的开始部份  

                                                routefirst = RouteProcess(panel.Machininglist[i], width, min_x, leadin_y);
                                                routefirst.RouteSetMillSequence = "RouteSequence";
                                                routefirst.RouteZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteSetMillZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteDiameter = "3";
                                                routefirst.RouteToolName = "D3";
                                                if (panel.Machininglist[i].Face == "5")
                                                {
                                                    face6list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRoutingFace6 = "TRUE";
                                                }
                                                else if (panel.Machininglist[i].Face == "6")
                                                {
                                                    face5list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRouting = "TRUE";
                                                }  //输出引线的结束部份  


                                                routefirst = RouteProcess(panel.Machininglist[i], width, min_x, min_y);
                                                routefirst.RouteSetMillSequence = "RouteSequence";
                                                routefirst.RouteZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteSetMillZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                routefirst.RouteDiameter = "3";
                                                routefirst.RouteToolName = "D3";
                                                if (panel.Machininglist[i].Face == "5")
                                                {
                                                    face6list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRoutingFace6 = "TRUE";
                                                }
                                                else if (panel.Machininglist[i].Face == "6")
                                                {
                                                    face5list.Add(routefirst.OutPutCsvString());
                                                    borderseq.FoundRouting = "TRUE";
                                                }  //输出XML第一个点的开始部份  


                                                for (int j = 0; j < panel.Machininglist[i].Linelist.Count; j++)  //输出铣型部份
                                                {
                                                    if (j == 0)
                                                        routesecond = RouteProcess(panel.Machininglist[i], width, max_x, min_y);
                                                    else if (j == 1)
                                                        routesecond = RouteProcess(panel.Machininglist[i], width, max_x, max_y);
                                                    else if (j == 2)
                                                        routesecond = RouteProcess(panel.Machininglist[i], width, min_x, max_y);
                                                    else if (j == 3)
                                                        routesecond = RouteProcess(panel.Machininglist[i], width, min_x, leadin_y - 3/2);

                                                    routesecond.RouteSetMillSequence = "RouteSequence";
                                                    routesecond.RouteZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                    routesecond.RouteSetMillZ = (needcuttingdepth1 - everycuttingdepth1 * (cutingtimes1 - l)).ToString();
                                                    routesecond.RouteDiameter = "3";
                                                    routesecond.RouteToolName = "D3";

                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        face6list.Add(routesecond.OutPutCsvString());
                                                        borderseq.FoundRoutingFace6 = "TRUE";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        face5list.Add(routesecond.OutPutCsvString());
                                                        borderseq.FoundRouting = "TRUE";
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (width == 8.5)
                                        {
                                            routefirst = RouteProcess(panel.Machininglist[i], width, min_x - 8.5/2, min_y);

                                            routefirst.RouteSetMillSequence = "RouteSetMillSequence";   //20180615 宋新刚 护墙板 单独8.5的矩形槽  这里原本错误

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }  //输出铣型的开始部份 

                                            routesecond = RouteProcess(panel.Machininglist[i], width, max_x + 8.5 / 2, min_y);

                                            routesecond.RouteSetMillSequence = "RouteSequence";

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routesecond.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routesecond.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }
                                        }
                                        else
                                        {   //这边的情况有如下几点：
                                            //1:铣的深度大于等于板厚
                                            //2:铣的宽度大于16
                                            //3：将原先的用的8.5的刀具替换成10的刀具。因为发现有25mm板厚的情况
                                            //4:将原先的所有的8.5的字符都替换成10


                                            double leadin_x = min_x + 30;   //20180418   //20180420
                                            double leadin_y = min_y + 10 + (width - 10) / 2;

                                            if (Math.Abs(height - double.Parse(panel.Length)) < 0.1)  //如果铣型发现在X方向铣通的话，则算刀具轨迹的时候多铣5mm 以将因刀产生的圆弧部份铣掉 20180414
                                            {
                                                min_x = min_x - 4;
                                                max_x = max_x + 4;
                                            }

                                            routefirst = RouteProcess(panel.Machininglist[i], width, leadin_x, leadin_y);

                                            routefirst.RouteDiameter = "10";  //20180725 替换的刀具
                                            routefirst.RouteToolName = "130";

                                            if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                            {
                                                routefirst.RouteEndTangentY = "1";

                                                if (Math.Abs(width - 15.1) < 0.1 || Math.Abs(width - 15.5) < 0.1)  //20180817 汉森专用滑门槽，在NESTING机床过滤。在PTP160上加工
                                                {
                                                    routefirst.RouteDiameter = "10.01";
                                                    routefirst.RouteToolName = "D130";
                                                }
                                            }


                                            if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                            {
                                                routefirst.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                                routefirst.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                            }

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }  //输出引线的开始部份  

                                            routefirst = RouteProcess(panel.Machininglist[i], width, min_x, leadin_y);
                                            routefirst.RouteSetMillSequence = "RouteSequence";


                                            routefirst.RouteDiameter = "10";  //20180725 替换的刀具
                                            routefirst.RouteToolName = "130";

                                            if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                            {
                                                routefirst.RouteEndTangentY = "1";

                                                if (Math.Abs(width - 15.1) < 0.1 || Math.Abs(width - 15.5) < 0.1)  //20180817 汉森专用滑门槽，在NESTING机床过滤。在PTP160上加工
                                                {
                                                    routefirst.RouteDiameter = "10.01";
                                                    routefirst.RouteToolName = "D130";
                                                }
                                            }

                                            if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                            {
                                                routefirst.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                                routefirst.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                            }

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }  //输出引线的结束部份  


                                            routefirst = RouteProcess(panel.Machininglist[i], width, min_x, min_y);
                                            routefirst.RouteSetMillSequence = "RouteSequence";
                                            routefirst.RouteDiameter = "10";  //20180725 替换的刀具
                                            routefirst.RouteToolName = "130";

                                            if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                            {
                                                routefirst.RouteEndTangentY = "1";

                                                if (Math.Abs(width - 15.1) < 0.1 || Math.Abs(width - 15.5) < 0.1)  //20180817 汉森专用滑门槽，在NESTING机床过滤。在PTP160上加工
                                                {
                                                    routefirst.RouteDiameter = "10.01";
                                                    routefirst.RouteToolName = "D130";
                                                }
                                            }

                                            if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                            {
                                                routefirst.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                                routefirst.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                            }

                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                face6list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                face5list.Add(routefirst.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }  //输出XML第一个点的开始部份  

                                            for (int j = 0; j < panel.Machininglist[i].Linelist.Count; j++)  //输出铣型部份
                                            {
                                                if (j == 0)
                                                    routesecond = RouteProcess(panel.Machininglist[i], width, max_x, min_y);
                                                else if (j == 1)
                                                    routesecond = RouteProcess(panel.Machininglist[i], width, max_x, max_y);
                                                else if (j == 2)
                                                    routesecond = RouteProcess(panel.Machininglist[i], width, min_x, max_y);
                                                else if (j == 3)
                                                    routesecond = RouteProcess(panel.Machininglist[i], width, min_x, (leadin_y - 10/2));

                                                routesecond.RouteSetMillSequence = "RouteSequence";

                                                routesecond.RouteDiameter = "10";  //20180725 替换的刀具
                                                routesecond.RouteToolName = "130";

                                                if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                                {
                                                    routesecond.RouteEndTangentY = "1";

                                                    if (Math.Abs(width - 15.1) < 0.1 || Math.Abs(width - 15.5) < 0.1)  //20180817 汉森专用滑门槽，在NESTING机床过滤。在PTP160上加工
                                                    {
                                                        routesecond.RouteDiameter = "10.01";
                                                        routesecond.RouteToolName = "D130";
                                                    }
                                                }

                                                if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                                {
                                                    routesecond.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                                    routesecond.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                                }

                                                if (panel.Machininglist[i].Face == "5")
                                                {
                                                    face6list.Add(routesecond.OutPutCsvString());
                                                    borderseq.FoundRoutingFace6 = "TRUE";
                                                }
                                                else if (panel.Machininglist[i].Face == "6")
                                                {
                                                    face5list.Add(routesecond.OutPutCsvString());
                                                    borderseq.FoundRouting = "TRUE";
                                                }
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Width 槽宽为:  " + width.ToString() + "没有合适的刀具进行加工!");
                                    return;
                                }
                            }
                        }
                        else
                        {
                            #region 做正常的拉槽 铣型 20180327

                            #region 异形四点从板件所有的点的最大值和最小值来判断。因为三维家有前缩进量等  20180404

                            if (panel.Machininglist[i].Linelist.Count - 1 > 0.1 && panel.Machininglist[i].GrooveType != "1")   //要过滤掉单独铣槽的部份   宋新刚 20180408 //20180713 如果在板上内轮廓上挖个孔等。则需要增加这个判断
                            {
                                double min_x = double.Parse(panel.Machininglist[i].X);
                                double min_y = double.Parse(panel.Machininglist[i].Y);
                                double max_x = double.Parse(panel.Machininglist[i].X);
                                double max_y = double.Parse(panel.Machininglist[i].Y);
                                double depth = double.Parse(panel.Machininglist[i].Depth);

                                for (int j = 0; j < panel.Machininglist[i].Linelist.Count; j++)
                                {
                                    min_x = Math.Min(min_x, double.Parse(panel.Machininglist[i].Linelist[j].EndX));
                                    min_y = Math.Min(min_y, double.Parse(panel.Machininglist[i].Linelist[j].EndY));
                                    max_x = Math.Max(max_x, double.Parse(panel.Machininglist[i].Linelist[j].EndX));
                                    max_y = Math.Max(max_y, double.Parse(panel.Machininglist[i].Linelist[j].EndY));
                                }

                                if (panel.Name.Contains("双向背光灯"))   //20180604发现三维家XML转过来的时候，关于双向背光灯和底板灯 开料尺寸有问题。原因是转换有问题
                                {
                                    max_x = double.Parse(panel.Length);
                                    max_y = double.Parse(panel.Width);
                                }
                                else if (panel.Name.Contains("底板灯"))
                                {
                                    min_x = 0;
                                    min_y = 0;
                                }

                                //if (Math.Abs(panel.Machininglist[i].Linelist.Count - 3) < 0.1)//如果是三角形，则人为的将点的坐标值扩大 脱离判断是顶点的条件 20180724
                                //{
                                //    min_x = min_x - 1;
                                //    min_y = min_y - 1;
                                //    max_x = max_x + 1;
                                //    max_y = max_y + 1;
                                //}

                                point4.Clear();//增加新的4个点时，将原先的点清除掉 20180724
                                point4.Add(new fourpoint(min_x, min_y));
                                point4.Add(new fourpoint(max_x, min_y));
                                point4.Add(new fourpoint(max_x, max_y));
                                point4.Add(new fourpoint(min_x, max_y));
                            }

                            #endregion

                        RouteSetMillSequenceEntity routesetmillseq = new RouteSetMillSequenceEntity();
                        routesetmillseq.RouteSetMillSequence = "RouteSetMillSequence";
                        routesetmillseq.RouteSetMillX = panel.Machininglist[i].X;
                        routesetmillseq.RouteSetMillY = panel.Machininglist[i].Y;
                        routesetmillseq.RouteSetMillZ = panel.Machininglist[i].Depth;
                        routesetmillseq.RouteStartOffsetX = panel.Machininglist[i].X;     //实则上要依据刀具偏置换算
                        routesetmillseq.RouteStartOffsetY = panel.Machininglist[i].Y;     //实则上要依据刀具偏置换算

                        if (panel.Machininglist[i].ToolName == "开料铣型刀") //20180331 //20180725 因为8.5的刀有效的长度只有28mm，所以不能切25厚以上的板。索性全改成10mm的刀
                        {
                            //    if (panel.Machininglist[i].GrooveType == "1")
                            //        {
                            //            routesetmillseq.RouteDiameter = "8.5";
                            //            routesetmillseq.RouteToolName = "129";
                            //        }
                            //    else
                            //        {
                                    routesetmillseq.RouteDiameter = "10";
                                    routesetmillseq.RouteToolName = "130";

                                    // }
                        }
                        else if (panel.Machininglist[i].ToolName == "开槽刀")   // 厨柜因为背板只有5mm，增加了6.35 20180419
                        {
                                if (Math.Abs(double.Parse(panel.Machininglist[i].Width) - 9) < 0.01) //20190102对厨柜的槽宽做了限制。如果不满足要求就报错.这里的逻辑重新写了一下
                                {
                                    routesetmillseq.RouteDiameter = "8.5";
                                    routesetmillseq.RouteToolName = "129";
                                }
                                else
                                {
                                    if (Math.Abs(double.Parse(panel.Machininglist[i].Width) - 6) < 0.01)
                                    {
                                        routesetmillseq.RouteDiameter = "6.35";
                                        routesetmillseq.RouteToolName = "131";
                                    }
                                    else
                                    {
                                        HaveLarger = true;
                                        
                                        if (panel.cabinet.OrderNo.Contains("CG"))
                                        MessageBox.Show("当前橱柜订单号为: " + panel.cabinet.OrderNo + "\n\n有问题的板号为: " + panel.ID + "\n\n有问题的板件名称为: " + panel.Name + "\n\n有问题的板件长宽为: " + panel.Length + " X " + panel.Width + "\n\n槽宽为: " + panel.Machininglist[i].Width
                                            + "\n\n请注意,橱柜的槽宽必须为 6mm，正反面加工码转换终止!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        else
                                            MessageBox.Show("当前订单号为: " + panel.cabinet.OrderNo + "\n\n有问题的板号为: " + panel.ID + "\n\n有问题的板件名称为: " + panel.Name + "\n\n有问题的板件长宽为: " + panel.Length + " X " + panel.Width + "\n\n槽宽为: " + panel.Machininglist[i].Width
                                                + "\n\n请注意,此槽宽并非班尔奇所规定的6mm槽或9mm槽的范围之内，正反面加工码转换终止!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                        return;
                                    }
                                }


                                #region 增加如果开通槽的时候 在通过式PTP上有锯片功能  20180330

                                double d1 = Convert.ToDouble(panel.Thickness);
                                double c1 = Convert.ToDouble(panel.Length);

                                double di = Convert.ToDouble(panel.Machininglist[i].Depth);
                                double pi = Convert.ToDouble(panel.Machininglist[i].X);

                                if (di < d1 && (pi <= 0 || pi >= c1) && Math.Abs(double.Parse(panel.Machininglist[i].Width) - 6) > 0.1)
                                {
                                    VdrillSequenceEntity vdrill = new VdrillSequenceEntity();
                                    double drillx = Convert.ToDouble(panel.Machininglist[i].X);
                                    if (drillx <= 0)
                                        vdrill.VDrillX = "0.1";
                                    else if (drillx >= c1)
                                        vdrill.VDrillX = (c1 - 0.1).ToString();

                                    if (Convert.ToDouble(routesetmillseq.RouteToolComp) == 0)
                                        vdrill.VDrillY = routesetmillseq.RouteSetMillY;
                                    else if (Convert.ToDouble(routesetmillseq.RouteToolComp) == 1)
                                        vdrill.VDrillY = (Convert.ToDouble(routesetmillseq.RouteSetMillY) + 0.5 * Convert.ToDouble(routesetmillseq.RouteDiameter)).ToString();
                                    else if (Convert.ToDouble(routesetmillseq.RouteToolComp) == 2)
                                        vdrill.VDrillY = (Convert.ToDouble(routesetmillseq.RouteSetMillY) - 0.5 * Convert.ToDouble(routesetmillseq.RouteDiameter)).ToString();

                                    vdrill.VdrillSequence = "VdrillSequence";
                                    vdrill.VDrillZ = routesetmillseq.RouteSetMillZ;
                                    vdrill.VDrillXOffset = "0";
                                    vdrill.VDrillYOffset = "0";
                                    vdrill.VDrillDiameter = "12.2";
                                    vdrill.VDrillToolName = "12.2mm";
                                    vdrill.VDrillFeedSpeed = "3000";
                                    vdrill.VDrillEntrySpeed = "8000";

                                    if (panel.Machininglist[i].Face == "5")
                                    {
                                        borderseq.FoundVdrillFace6 = "TRUE";
                                        face6list.Add(vdrill.OutPutCsvString());
                                    }
                                    else if (panel.Machininglist[i].Face == "6")
                                    {
                                        borderseq.FoundVdrill = "TRUE";
                                        face5list.Add(vdrill.OutPutCsvString());
                                    }
                                }
                                #endregion
                        }


                        routesetmillseq.RoutePreviousToolName = "";
                        routesetmillseq.RouteNextToolName = "";
                        routesetmillseq.RouteFeedSpeed = "";
                        routesetmillseq.RouteEntrySpeed = "";
                        routesetmillseq.RouteBitType = "";
                        routesetmillseq.RouteRotation = "";

                        if (panel.Machininglist[i].ToolOffset == " ") // 宋新刚 20180326
                            routesetmillseq.RouteToolComp = "0";
                        else if (panel.Machininglist[i].ToolOffset == "左")
                            routesetmillseq.RouteToolComp = "1";
                        else if (panel.Machininglist[i].ToolOffset == "右")
                            routesetmillseq.RouteToolComp = "2";

                        routesetmillseq.RouteX = panel.Machininglist[i].X;     //实则上要依据刀具偏置换算
                        routesetmillseq.RouteY = panel.Machininglist[i].Y;     //实则上要依据刀具偏置换算
                        routesetmillseq.RouteZ = panel.Machininglist[i].Depth;
                        routesetmillseq.RouteEndOffsetX = "";
                        routesetmillseq.RouteEndOffsetY = "";
                        routesetmillseq.RouteBulge = "0";
                        routesetmillseq.RouteRadius = "";   //如果拱高比不为0 则必须要此值
                        routesetmillseq.RouteCenterX = "";  //如果拱高比不为0 则必须要此值
                        routesetmillseq.RouteCenterY = "";  //如果拱高比不为0 则必须要此值
                        routesetmillseq.RouteNextX = "";
                        routesetmillseq.RouteNextY = "";
                        routesetmillseq.RoutePreviousX = "";
                        routesetmillseq.RoutePreviousY = "";
                        routesetmillseq.RoutePreviousZ = "";
                        routesetmillseq.RouteBulgeNext = "";
                        routesetmillseq.RouteSetMillCounter = "";  //AD里是要的 为了标记连续轮廓
                        routesetmillseq.RouteVectorCounter ="";
                        routesetmillseq.RouteVectorCount = "";
                        routesetmillseq.RouteAngle = "";
                        routesetmillseq.RoutePreviousFeedSpeed = "";
                        routesetmillseq.ArcPeakX = "";
                        routesetmillseq.ArcPeakY = "";
                        routesetmillseq.ArcPeakBulge = "";
                        routesetmillseq.RoutePreviousBulge = "";
                        routesetmillseq.RouteRotation = "";
                        routesetmillseq.RouteSpindleSpeed = "";
                        routesetmillseq.RouteStartTangentX = "";
                        routesetmillseq.RouteStartTangentY = "";
                        routesetmillseq.RouteEndTangentX = "";
                        routesetmillseq.RouteEndTangentY = "";


                        #region 判断第一个点是不是工件的四个顶点。如果是四个顶点的其中一个，则删点。如果不是，则暂时跳过步骤

                        List<fourpoint> pointxy1 = new List<fourpoint>();

                        List<fourpoint> partnotfourpoints = new List<fourpoint>();  //建个容器用来保存不是顶点的点.如果第一个点就不是顶点中的点，则不需要增加，因为一段多段线的最后一个点肯定也是这个点
                                               
                        bool firstcomeout = true;
                        bool firstcomeout1 = true;
                        bool yesornofourpoint = false;
                        bool comeoutallprofile = false;
                        bool needallprofile = false;
                        if (Isforpoint(point4,(new fourpoint(double.Parse(routesetmillseq.RouteX),double.Parse(routesetmillseq.RouteY)))))
                        {                         
                            pointxy1.Add(new fourpoint(double.Parse(routesetmillseq.RouteX), double.Parse(routesetmillseq.RouteY)));
                            yesornofourpoint = true;
                        }
                        else
                        {
                            yesornofourpoint = false;
                            partnotfourpoints.Add(new fourpoint(double.Parse(routesetmillseq.RouteX), double.Parse(routesetmillseq.RouteY)));
                        }
                        #endregion

                        for (int j = 0;j<panel.Machininglist[i].Linelist.Count;j++)
                        {
                            RouteSetMillSequenceEntity routeseq = new RouteSetMillSequenceEntity();
                            routeseq.RouteSetMillSequence = "RouteSequence";
                            routeseq.RouteSetMillX = panel.Machininglist[i].X;
                            routeseq.RouteSetMillY = panel.Machininglist[i].Y;
                            routeseq.RouteSetMillZ = panel.Machininglist[i].Depth;
                            routeseq.RouteStartOffsetX = panel.Machininglist[i].X;     //实则上要依据刀具偏置换算
                            routeseq.RouteStartOffsetY = panel.Machininglist[i].Y;     //实则上要依据刀具偏置换算

                            if (panel.Machininglist[i].ToolName == "开料铣型刀") //20180331
                            {
                                //if (panel.Machininglist[i].GrooveType == "1")  //20180725 因为8.5的刀有效的长度只有28mm，所以不能切25厚以上的板。索性全改成10mm的刀
                                //    {
                                //        routeseq.RouteDiameter = "8.5";
                                //        routeseq.RouteToolName = "129";
                                //    }
                                //else
                                //    {
                                        routeseq.RouteDiameter = "10";
                                        routeseq.RouteToolName = "130";
                                   // }
                            }
                            else if (panel.Machininglist[i].ToolName == "开槽刀")  // 厨柜因为背板只有5mm，增加了6.35 20180419
                                {

                                    if (Math.Abs(double.Parse(panel.Machininglist[i].Width) - 9) < 0.01) //20190102对厨柜的槽宽做了限制。如果不满足要求就报错.这里的逻辑重新写了一下
                                    {
                                        routeseq.RouteDiameter = "8.5";
                                        routeseq.RouteToolName = "129";
                                    }
                                    else
                                    {
                                        if (Math.Abs(double.Parse(panel.Machininglist[i].Width) - 6) < 0.01)
                                        {
                                            routeseq.RouteDiameter = "6.35";
                                            routeseq.RouteToolName = "131";
                                        }
                                        else
                                        {
                                            HaveLarger = true;

                                            if (panel.cabinet.OrderNo.Contains("CG"))
                                                MessageBox.Show("当前橱柜订单号为: " + panel.cabinet.OrderNo + "\n\n有问题的板号为: " + panel.ID + "\n\n有问题的板件名称为: " + panel.Name + "\n\n有问题的板件长宽为: " + panel.Length + " X " + panel.Width + "\n\n槽宽为: " + panel.Machininglist[i].Width
                                                    + "\n\n请注意,橱柜的槽宽必须为 6mm，正反面加工码转换终止!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            else
                                                MessageBox.Show("当前订单号为: " + panel.cabinet.OrderNo + "\n\n有问题的板号为: " + panel.ID + "\n\n有问题的板件名称为: " + panel.Name + "\n\n有问题的板件长宽为: " + panel.Length + " X " + panel.Width + "\n\n槽宽为: " + panel.Machininglist[i].Width
                                                    + "\n\n请注意,此槽宽并非班尔奇所规定的6mm槽或9mm槽的范围之内，正反面加工码转换终止!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                            return;
                                        }
                                    }

                                    #region 增加如果开通槽的时候 在通过式PTP上有锯片功能  20180330

                                    double d1 = Convert.ToDouble(panel.Thickness);
                                    double c1 = Convert.ToDouble(panel.Length);

                                    double di = Convert.ToDouble(panel.Machininglist[i].Depth);
                                    double pi = Convert.ToDouble(panel.Machininglist[i].Linelist[j].EndX);

                                    if (di < d1 && (pi <= 0 || pi >= c1) && Math.Abs(double.Parse(panel.Machininglist[i].Width) - 6) > 0.1)
                                    {
                                        VdrillSequenceEntity vdrill = new VdrillSequenceEntity();
                                        double drillx = Convert.ToDouble(panel.Machininglist[i].Linelist[j].EndX);
                                        if (drillx <= 0)
                                            vdrill.VDrillX = "0.1";
                                        else if (drillx >= c1)
                                            vdrill.VDrillX = (c1 - 0.1).ToString();

                                        if (Convert.ToDouble(routesetmillseq.RouteToolComp) == 0)
                                            vdrill.VDrillY = routesetmillseq.RouteSetMillY;
                                        else if (Convert.ToDouble(routesetmillseq.RouteToolComp) == 1)
                                            vdrill.VDrillY = (Convert.ToDouble(routesetmillseq.RouteSetMillY) + 0.5 * Convert.ToDouble(routesetmillseq.RouteDiameter)).ToString();
                                        else if (Convert.ToDouble(routesetmillseq.RouteToolComp) == 2)
                                            vdrill.VDrillY = (Convert.ToDouble(routesetmillseq.RouteSetMillY) - 0.5 * Convert.ToDouble(routesetmillseq.RouteDiameter)).ToString();

                                        vdrill.VdrillSequence = "VdrillSequence";
                                        vdrill.VDrillZ = routesetmillseq.RouteSetMillZ;
                                        vdrill.VDrillXOffset = "0";
                                        vdrill.VDrillYOffset = "0";
                                        vdrill.VDrillDiameter = "12.2";
                                        vdrill.VDrillToolName = "12.2mm";
                                        vdrill.VDrillFeedSpeed = "3000";
                                        vdrill.VDrillEntrySpeed = "8000";

                                        if (panel.Machininglist[i].Face == "5")
                                        {
                                            borderseq.FoundVdrillFace6 = "TRUE";
                                            face6list.Add(vdrill.OutPutCsvString());
                                        }
                                        else if (panel.Machininglist[i].Face == "6")
                                        {
                                            borderseq.FoundVdrill = "TRUE";
                                            face5list.Add(vdrill.OutPutCsvString());
                                        }
                                    }
                                    #endregion
                                }

                            routeseq.RoutePreviousToolName = "";
                            routeseq.RouteNextToolName = "";
                            routeseq.RouteFeedSpeed = "";
                            routeseq.RouteEntrySpeed = "";
                            routeseq.RouteBitType = "";
                            routeseq.RouteRotation = "";

                            if (panel.Machininglist[i].ToolOffset.Contains(" ")) //宋新刚 20180326
                                routeseq.RouteToolComp = "0";
                            else if (panel.Machininglist[i].ToolOffset.Contains("左"))
                                routeseq.RouteToolComp = "1";
                            else if (panel.Machininglist[i].ToolOffset.Contains("右"))
                                routeseq.RouteToolComp = "2";

                            routeseq.RouteX = panel.Machininglist[i].Linelist[j].EndX;     //实则上要依据刀具偏置换算
                            routeseq.RouteY = panel.Machininglist[i].Linelist[j].EndY;     //实则上要依据刀具偏置换算
                            routeseq.RouteZ = panel.Machininglist[i].Depth;
                            routeseq.RouteEndOffsetX = "";
                            routeseq.RouteEndOffsetY = "";


                            #region 拱高比计算
                            double angle = Convert.ToDouble(panel.Machininglist[i].Linelist[j].Angle);
                            double numflag1 = Math.PI / 180;
                            double numflag2 = numflag1 * (angle / 2);

                            if (numflag2 == 0)
                                routeseq.RouteBulge = "0";
                            else 
                                routeseq.RouteBulge = ((1 - Math.Cos(numflag2)) / Math.Sin(numflag2)* -1).ToString("F5");
                            #endregion

                            #region 计算圆弧半径及圆心坐标
                            if (routeseq.RouteBulge != "0")
                                {
                                    double x1 = 0, x2 = 0, y1 = 0, y2 = 0, l = 0, u = 0, radius = 0, ang = 0, cenx = 0, ceny = 0;
                                    
                                    
                                    double.TryParse(routeseq.RouteX, out x2);
                                    double.TryParse(routeseq.RouteY, out y2);

                                    if (j == 0)  //如果铣一个台面的圆孔 则发现j = 0的时候就有角度是90的情况，则会报错  20180331
                                    {
                                        double.TryParse(panel.Machininglist[i].X, out x1);
                                        double.TryParse(panel.Machininglist[i].Y, out y1);
                                    }
                                    else
                                    {
                                        double.TryParse(panel.Machininglist[i].Linelist[j - 1].EndX, out x1);
                                        double.TryParse(panel.Machininglist[i].Linelist[j - 1].EndY, out y1);
                                    }



                                     double.TryParse(routeseq.RouteBulge, out u);

                                    l = Math.Sqrt(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2));
                                    radius = 0.25 * l * (u + 1 / u);

                                    routeseq.RouteRadius = string.Format("{0:f4}", radius);
                                    if (x1 == x2)
                                    {
                                        ang = Math.PI / 2;
                                        if (u > 0)
                                        {
                                            cenx = (x1 + x2) / 2 + (radius - l * u / 2);
                                            ceny = (y1 + y2) / 2;
                                            if ((x2 - x1) * (ceny - y2) - (y2 - y1) * (cenx - x2) < 0)
                                            {
                                                routeseq.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                                routeseq.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                            }
                                            else
                                            {
                                                cenx = (x1 + x2) / 2 - (radius - l * u / 2);
                                                ceny = (y1 + y2) / 2;
                                                routeseq.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                                routeseq.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                            }
                                        }
                                        else
                                        {
                                            cenx = (x1 + x2) / 2 - (radius - l * u / 2);
                                            ceny = (y1 + y2) / 2;
                                            if ((x2 - x1) * (ceny - y2) - (y2 - y1) * (cenx - x2) < 0)
                                            {
                                                routeseq.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                                routeseq.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                            }
                                            else
                                            {
                                                cenx = (x1 + x2) / 2 + (radius - l * u / 2);
                                                ceny = (y1 + y2) / 2;
                                                routeseq.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                                routeseq.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        ang = Math.Atan((y1 - y2) / (x1 - x2));
                                        if (u > 0)
                                        {
                                            cenx = (x1 + x2) / 2 + (radius - l * u / 2) * Math.Sin(ang);
                                            ceny = (y1 + y2) / 2 - (radius - l * u / 2) * Math.Cos(ang);
                                            if ((x2 - x1) * (ceny - y2) - (y2 - y1) * (cenx - x2) < 0)
                                            {
                                                routeseq.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                                routeseq.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                            }
                                            else
                                            {
                                                cenx = (x1 + x2) / 2 - (radius - l * u / 2) * Math.Sin(ang);
                                                ceny = (y1 + y2) / 2 + (radius - l * u / 2) * Math.Cos(ang);
                                                routeseq.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                                routeseq.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                            }
                                        }
                                        else
                                        {
                                            cenx = (x1 + x2) / 2 - (radius - l * u / 2) * Math.Sin(ang);
                                            ceny = (y1 + y2) / 2 + (radius - l * u / 2) * Math.Cos(ang);
                                            if ((x2 - x1) * (ceny - y2) - (y2 - y1) * (cenx - x2) < 0)
                                            {
                                                routeseq.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                                routeseq.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                            }
                                            else
                                            {
                                                cenx = (x1 + x2) / 2 + (radius - l * u / 2) * Math.Sin(ang);
                                                ceny = (y1 + y2) / 2 - (radius - l * u / 2) * Math.Cos(ang);
                                                routeseq.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                                routeseq.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                            }
                                        }
                                    }
                                    double.TryParse(routeseq.RouteX, out x1);   //记录圆弧的点的上一点
                                    double.TryParse(routeseq.RouteY, out y1);
                                }
                                else
                                {
                                    routeseq.RouteCenterX = "";
                                    routeseq.RouteCenterY = "";
                                    routeseq.RouteRadius = "";
                                }

                            #endregion

                            routeseq.RouteNextX = "";
                            routeseq.RouteNextY = "";
                            routeseq.RoutePreviousX = "";
                            routeseq.RoutePreviousY = "";
                            routeseq.RoutePreviousZ = "";
                            routeseq.RouteBulgeNext = "";
                            routeseq.RouteSetMillCounter = "";  //AD里是要的 为了标记连续轮廓
                            routeseq.RouteVectorCounter = "";
                            routeseq.RouteVectorCount = "";
                            routeseq.RouteAngle = "";
                            routeseq.RoutePreviousFeedSpeed = "";
                            routeseq.ArcPeakX = "";
                            routeseq.ArcPeakY = "";
                            routeseq.ArcPeakBulge = "";
                            routeseq.RoutePreviousBulge = "";
                            routeseq.RouteRotation = "";
                            routeseq.RouteSpindleSpeed = "";
                            routeseq.RouteStartTangentX = "";
                            routeseq.RouteStartTangentY = "";
                            routeseq.RouteEndTangentX = "";
                            routeseq.RouteEndTangentY = "";

                            #region 判断点的特殊处理
                            //如果第1个点在顶点中  处理
                            //如果第1个点不在顶点 第2个点在顶点 处理
                            //如果第1个点不在顶点 第2个点也不在顶点 后面的情况未处理
                            List<fourpoint> pointxy2 = new List<fourpoint>();
                            pointxy2.Add(new fourpoint(double.Parse(routeseq.RouteX), double.Parse(routeseq.RouteY)));

                            if (!yesornofourpoint && !Isforpoint(point4, (new fourpoint(double.Parse(routeseq.RouteX), double.Parse(routeseq.RouteY)))) && !needallprofile)
                            {
                                comeoutallprofile = true;
                                needallprofile = true;
                            }

                            if (comeoutallprofile)
                            {
                                if (firstcomeout1)
                                {
                                    firstcomeout1 = false;
                                    routeseq.RouteSetMillSequence = "RouteSetMillSequence";
                                    if (pointxy1.Count > 0)
                                    {
                                        foreach (fourpoint fpt in pointxy1)
                                        {
                                            routeseq.RouteSetMillX = fpt.x.ToString();
                                            routeseq.RouteSetMillY = fpt.y.ToString();
                                            routeseq.RouteX = fpt.x.ToString();
                                            routeseq.RouteY = fpt.y.ToString();
                                        }
                                    }
                                    else
                                    {
                                        foreach (fourpoint fpt in partnotfourpoints) //如果是槽 则取出槽的顶点 宋新刚20180326
                                        {
                                            routeseq.RouteSetMillX = fpt.x.ToString();
                                            routeseq.RouteSetMillY = fpt.y.ToString();
                                            routeseq.RouteX = fpt.x.ToString();
                                            routeseq.RouteY = fpt.y.ToString();
                                        }

                                    }


                                    #region 用来记录铣型的时候，刀具在切入点时的位置  宋新刚20180320
                                    if (routeseq.RouteDiameter == "10")
                                    {
                                        double a = Convert.ToDouble(panel.Machininglist[i].Linelist[j].EndX) - Convert.ToDouble(routeseq.RouteX);
                                        double b = Convert.ToDouble(panel.Machininglist[i].Linelist[j].EndY) - Convert.ToDouble(routeseq.RouteY);

                                        if (a >= 0 && b >= 0 && routeseq.RouteToolComp == "2")
                                        {
                                            routeseq.RouteStartOffsetX = routeseq.RouteX;
                                            routeseq.RouteStartOffsetY = (Convert.ToDouble(routeseq.RouteY) - (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                        }
                                        else if (a <= 0 && b <= 0 && routeseq.RouteToolComp == "2")
                                        {
                                            routeseq.RouteStartOffsetX = routeseq.RouteX;
                                            routeseq.RouteStartOffsetY = (Convert.ToDouble(routeseq.RouteY) + (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                        }
                                        else if (a <= 0 && b >= 0 && routeseq.RouteToolComp == "2")
                                        {
                                            routeseq.RouteStartOffsetX = (Convert.ToDouble(routeseq.RouteX) + (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString(); ;
                                            routeseq.RouteStartOffsetY = routeseq.RouteY;
                                        }
                                        else if (a >= 0 && b <= 0 && routeseq.RouteToolComp == "2")
                                        {
                                            routeseq.RouteStartOffsetX = (Convert.ToDouble(routeseq.RouteX) - (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString(); ;
                                            routeseq.RouteStartOffsetY = routeseq.RouteY;
                                        }
                                        else if (a >= 0 && b >= 0 && routeseq.RouteToolComp == "1") //20180725已验证
                                        {
                                            routeseq.RouteStartOffsetX = (Convert.ToDouble(routeseq.RouteX) + (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                            routeseq.RouteStartOffsetY = routeseq.RouteY;
                                        }
                                        else if (a <= 0 && b <= 0 && routeseq.RouteToolComp == "1")
                                        {
                                            routeseq.RouteStartOffsetX = (Convert.ToDouble(routeseq.RouteX) - (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                                routeseq.RouteStartOffsetY = routeseq.RouteY;
                                        }
                                        else if (a <= 0 && b >= 0 && routeseq.RouteToolComp == "1")
                                        {
                                            routeseq.RouteStartOffsetX = routeseq.RouteX;
                                            routeseq.RouteStartOffsetY = (Convert.ToDouble(routeseq.RouteY) - (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                        }
                                        else if (a >= 0 && b <= 0 && routeseq.RouteToolComp == "1")
                                        {
                                            routeseq.RouteStartOffsetX = routeseq.RouteX;
                                            routeseq.RouteStartOffsetY = (Convert.ToDouble(routeseq.RouteY) + (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                        }
                                        else
                                        {
                                            MessageBox.Show("开始点超出统计范围(一)，请与宋新刚 18913812043联系！");
                                        }
                                    }
                                        #endregion

                                        #region 判断内轮廓的前3个点来确定是顺时针走向还是逆时针走向 20180713 宋新刚


                                        if (panel.Machininglist[i].GrooveType == "1" && panel.Machininglist[i].Linelist.Count == 2)  //仅仅是圆才需要修改 20180801
                                        {
                                            //三维家的苏镇城说：
                                            //1。孔的制作，有两种方式，一种是用写点的孔，一种是直接生成，
                                            //2。如果用直接生成的孔，固定是逆时针的，用写点的孔，是顺逆都会生成的。
                                          
                                            
                                            if (routeseq.RouteToolComp == "2")
                                            {
                                                routeseq.RouteToolComp = "1";
                                            }
                                            if (Convert.ToDouble(routeseq.RouteBulge) - 1 < 0.01)
                                            {
                                                routeseq.RouteBulge = "-1";
                                            }
                                        }

                                        #endregion

                                        if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                            routeseq.RouteEndTangentY = "1";

                                        #region 0 = 无槽  1 = 标准20的槽   2 = 非标20的槽  20180807
                                        if (panel.cabinet.OrderNo.Contains("CG") && panel.Machininglist[i].ToolName.Equals("开槽刀") && panel.Machininglist[i].Linelist.Count == 1) // 20180806   0 = 无槽  1 = 标准20的槽   2 = 非标20的槽
                                        {
                                            if (Convert.ToDouble(panel.Machininglist[i].Width) - 6 < 0.1)
                                            {
                                                double panelwidth = Convert.ToDouble(panel.Width);
                                                double bbwidth = Convert.ToDouble(panel.Machininglist[i].Width) - 1;
                                                double standcao = 20;
                                                if (Math.Abs((Convert.ToDouble(panel.Machininglist[i].Y) - bbwidth / 2 + standcao) - panelwidth) < 0.1)
                                                {
                                                    if (panel.Machininglist[i].Face == "5" && Cao_6 != "2")
                                                    {
                                                        Cao_6 = "1";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6" && Cao_5 != "2")
                                                    {
                                                        Cao_5 = "1";
                                                    }
                                                }
                                                else
                                                {
                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        Cao_6 = "2";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        Cao_5 = "2";
                                                    }
                                                }
                                            }
                                            else if (Convert.ToDouble(panel.Machininglist[i].Width) - 9 < 0.1) //20180818 增加如果是8.5的标准槽，无论是不是20 都认为是非标槽
                                            {
                                                if (panel.Machininglist[i].Face == "5")
                                                {
                                                    Cao_6 = "2";
                                                }
                                                else if (panel.Machininglist[i].Face == "6")
                                                {
                                                    Cao_5 = "2";
                                                }
                                            }

                                            borderseq.SawsCount = Cao_5 + ";" + Cao_6;
                                        }
                                        #endregion

                                        if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                        {
                                            routeseq.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                            routeseq.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                        }

                                        if (panel.Machininglist[i].Face == "5")
                                        {
                                            if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  
                                            {
                                                face5list.Add(routeseq.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }
                                            else
                                            {
                                                face6list.Add(routeseq.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }

                                        }
                                        else if (panel.Machininglist[i].Face == "6")
                                        {
                                            face5list.Add(routeseq.OutPutCsvString());
                                            borderseq.FoundRouting = "TRUE";
                                        }
                                    }
                            }
                            else if (Isforpoint(point4, (new fourpoint(double.Parse(routeseq.RouteX), double.Parse(routeseq.RouteY)))))
                            {
                                needallprofile = true;
                                yesornofourpoint = true;
                                pointxy1.Clear();
                                pointxy1.Add(new fourpoint(double.Parse(routeseq.RouteX), double.Parse(routeseq.RouteY)));
                                continue;
                            }
                            else if (twopointisHorVline(pointxy1, pointxy2) && yesornofourpoint)
                            {
                                needallprofile = true;
                                yesornofourpoint = false;
                                pointxy1.Clear();
                                pointxy1.Add(new fourpoint(double.Parse(routeseq.RouteX), double.Parse(routeseq.RouteY)));
                                firstcomeout = true;  //转角柜带切柱的板 相当于切的部份有两段线  宋新刚 20180213
                                continue;

                            }
                            else
                            {
                                needallprofile = true;
                                if (firstcomeout)
                                {
                                    firstcomeout = false;
                                    routeseq.RouteSetMillSequence = "RouteSetMillSequence";
                                    foreach (fourpoint fpt in pointxy1)
                                    {
                                        routeseq.RouteSetMillX = fpt.x.ToString();
                                        routeseq.RouteSetMillY = fpt.y.ToString();
                                        routeseq.RouteX = fpt.x.ToString();
                                        routeseq.RouteY = fpt.y.ToString();
                                    }

                                    #region 用来记录铣型的时候，刀具在切入点时的位置  宋新刚20180320
                                    if (routeseq.RouteDiameter == "10")
                                    {
                                        double a = Convert.ToDouble(panel.Machininglist[i].Linelist[j].EndX) - Convert.ToDouble(routeseq.RouteX);
                                        double b = Convert.ToDouble(panel.Machininglist[i].Linelist[j].EndY) - Convert.ToDouble(routeseq.RouteY);

                                        if (a >= 0 && b >= 0 && routeseq.RouteToolComp == "2")
                                        {
                                            routeseq.RouteStartOffsetX = routeseq.RouteX;
                                            routeseq.RouteStartOffsetY = (Convert.ToDouble(routeseq.RouteY) - (Convert.ToDouble(routeseq.RouteDiameter))/2).ToString();
                                        }
                                        else if (a <= 0 && b <= 0 && routeseq.RouteToolComp == "2")
                                        {
                                            routeseq.RouteStartOffsetX = routeseq.RouteX;
                                            routeseq.RouteStartOffsetY = (Convert.ToDouble(routeseq.RouteY) + (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                        }
                                        else if (a <= 0 && b >= 0 && routeseq.RouteToolComp == "2")
                                        {
                                            routeseq.RouteStartOffsetX = (Convert.ToDouble(routeseq.RouteX) + (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString(); ;
                                            routeseq.RouteStartOffsetY = routeseq.RouteY;
                                        }
                                        else if (a >= 0 && b <= 0 && routeseq.RouteToolComp == "2")
                                        {
                                            routeseq.RouteStartOffsetX = (Convert.ToDouble(routeseq.RouteX) - (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString(); ;
                                            routeseq.RouteStartOffsetY = routeseq.RouteY;
                                        }
                                        else if (a >= 0 && b >= 0 && routeseq.RouteToolComp == "1") //20180725已验证
                                        {
                                            routeseq.RouteStartOffsetX = (Convert.ToDouble(routeseq.RouteX) + (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                            routeseq.RouteStartOffsetY = routeseq.RouteY;
                                        }
                                        else if (a <= 0 && b <= 0 && routeseq.RouteToolComp == "1")
                                        {
                                            routeseq.RouteStartOffsetX = (Convert.ToDouble(routeseq.RouteX) - (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                            routeseq.RouteStartOffsetY = routeseq.RouteY;
                                        }
                                        else if (a <= 0 && b >= 0 && routeseq.RouteToolComp == "1")
                                        {
                                            routeseq.RouteStartOffsetX = routeseq.RouteX;
                                            routeseq.RouteStartOffsetY = (Convert.ToDouble(routeseq.RouteY) - (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                        }
                                        else if (a >= 0 && b <= 0 && routeseq.RouteToolComp == "1")
                                        {
                                            routeseq.RouteStartOffsetX = routeseq.RouteX;
                                            routeseq.RouteStartOffsetY = (Convert.ToDouble(routeseq.RouteY) + (Convert.ToDouble(routeseq.RouteDiameter)) / 2).ToString();
                                        }
                                        else
                                        {
                                            MessageBox.Show("开始点超出统计范围(二)，请与宋新刚 18913812043联系！");
                                        }
                                    }
                                        #endregion

                                        #region 判断内轮廓的前3个点来确定是顺时针走向还是逆时针走向 20180713 宋新刚

                                        if (panel.Machininglist[i].GrooveType == "1" && panel.Machininglist[i].Linelist.Count == 2)//仅仅是圆才需要修改 20180801
                                        {
                                            //三维家的苏镇城说：
                                            //1。孔的制作，有两种方式，一种是用写点的孔，一种是直接生成，
                                            //2。如果用直接生成的孔，固定是逆时针的，用写点的孔，是顺逆都会生成的。
                                            //所以，这个顺时针与逆时针，是不固定的。需要生产程序做处理。

                                            if (routeseq.RouteToolComp == "2")
                                            {
                                                routeseq.RouteToolComp = "1";
                                            }

                                            if (Convert.ToDouble(routeseq.RouteBulge) - 1 < 0.01)
                                            {
                                                routeseq.RouteBulge = "-1";
                                            }
                                        }

                                        #endregion

                                        if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                            routeseq.RouteEndTangentY = "1";

                                        #region 0 = 无槽  1 = 标准20的槽   2 = 非标20的槽  20180807
                                        if (panel.cabinet.OrderNo.Contains("CG") && panel.Machininglist[i].ToolName.Equals("开槽刀") && panel.Machininglist[i].Linelist.Count == 1) // 20180806   0 = 无槽  1 = 标准20的槽   2 = 非标20的槽
                                        {
                                            if (Convert.ToDouble(panel.Machininglist[i].Width) - 6 < 0.1)
                                            {
                                                double panelwidth = Convert.ToDouble(panel.Width);
                                                double bbwidth = Convert.ToDouble(panel.Machininglist[i].Width) - 1;
                                                double standcao = 20;
                                                if (Math.Abs((Convert.ToDouble(panel.Machininglist[i].Y) - bbwidth / 2 + standcao) - panelwidth) < 0.1)
                                                {
                                                    if (panel.Machininglist[i].Face == "5" && Cao_6 != "2")
                                                    {
                                                        Cao_6 = "1";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6" && Cao_5 != "2")
                                                    {
                                                        Cao_5 = "1";
                                                    }
                                                }
                                                else
                                                {
                                                    if (panel.Machininglist[i].Face == "5")
                                                    {
                                                        Cao_6 = "2";
                                                    }
                                                    else if (panel.Machininglist[i].Face == "6")
                                                    {
                                                        Cao_5 = "2";
                                                    }
                                                }
                                            }
                                            else if (Convert.ToDouble(panel.Machininglist[i].Width) - 9 < 0.1) //20180818 增加如果是8.5的标准槽，无论是不是20 都认为是非标槽
                                            {
                                                if (panel.Machininglist[i].Face == "5")
                                                {
                                                    Cao_6 = "2";
                                                }
                                                else if (panel.Machininglist[i].Face == "6")
                                                {
                                                    Cao_5 = "2";
                                                }
                                            }

                                            borderseq.SawsCount = Cao_5 + ";" + Cao_6;
                                        }
                                        #endregion

                                        if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                        {
                                            routeseq.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                            routeseq.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                        }

                                        if (panel.Machininglist[i].Face == "5")
                                        {
                                            if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  
                                            {
                                                face5list.Add(routeseq.OutPutCsvString());
                                                borderseq.FoundRouting = "TRUE";
                                            }
                                            else
                                            {
                                                face6list.Add(routeseq.OutPutCsvString());
                                                borderseq.FoundRoutingFace6 = "TRUE";
                                            }

                                        }
                                        else if (panel.Machininglist[i].Face == "6")
                                        {
                                            face5list.Add(routeseq.OutPutCsvString());
                                            borderseq.FoundRouting = "TRUE";
                                        }
                                    }

                             }
                            #endregion

                            routeseq.RouteSetMillSequence = "RouteSequence";
                            routeseq.RouteX = panel.Machininglist[i].Linelist[j].EndX;
                            routeseq.RouteY = panel.Machininglist[i].Linelist[j].EndY;

                                #region 判断内轮廓的前3个点来确定是顺时针走向还是逆时针走向 20180713 宋新刚

                                if (panel.Machininglist[i].GrooveType == "1" && panel.Machininglist[i].Linelist.Count == 2)//仅仅是圆才需要修改 20180801
                                {
                                    //三维家的苏镇城说：
                                    //1。孔的制作，有两种方式，一种是用写点的孔，一种是直接生成，
                                    //2。如果用直接生成的孔，固定是逆时针的，用写点的孔，是顺逆都会生成的。
                                    //所以，这个顺时针与逆时针，是不固定的。需要生产程序做处理。

                                    if (routeseq.RouteToolComp == "2")
                                    {
                                        routeseq.RouteToolComp = "1";
                                    }
                                    if (Convert.ToDouble(routeseq.RouteBulge) - 1 < 0.01)
                                    {
                                        routeseq.RouteBulge = "-1";
                                    }
                                }

                                #endregion

                                if (panel.Machininglist[i].GrooveType == "1")  // 20180726 如果是内轮廓 则不做减封边处理 
                                    routeseq.RouteEndTangentY = "1";

                                #region 0 = 无槽  1 = 标准20的槽   2 = 非标20的槽  20180807
                                if (panel.cabinet.OrderNo.Contains("CG") && panel.Machininglist[i].ToolName.Equals("开槽刀") && panel.Machininglist[i].Linelist.Count == 1) // 20180806   0 = 无槽  1 = 标准20的槽   2 = 非标20的槽
                                {
                                    if (Convert.ToDouble(panel.Machininglist[i].Width) - 6 < 0.1)
                                    {
                                        double panelwidth = Convert.ToDouble(panel.Width);
                                        double bbwidth = Convert.ToDouble(panel.Machininglist[i].Width) - 1;
                                        double standcao = 20;
                                        if (Math.Abs((Convert.ToDouble(panel.Machininglist[i].Y) - bbwidth / 2 + standcao) - panelwidth) < 0.1)
                                        {
                                            if (panel.Machininglist[i].Face == "5" && Cao_6 != "2")
                                            {
                                                Cao_6 = "1";
                                            }
                                            else if (panel.Machininglist[i].Face == "6" && Cao_5 != "2")
                                            {
                                                Cao_5 = "1";
                                            }
                                        }
                                        else
                                        {
                                            if (panel.Machininglist[i].Face == "5")
                                            {
                                                Cao_6 = "2";
                                            }
                                            else if (panel.Machininglist[i].Face == "6")
                                            {
                                                Cao_5 = "2";
                                            }
                                        }
                                    }
                                    else if (Convert.ToDouble(panel.Machininglist[i].Width) - 9 < 0.1) //20180818 增加如果是8.5的标准槽，无论是不是20 都认为是非标槽
                                    {
                                        if (panel.Machininglist[i].Face == "5")
                                        {
                                            Cao_6 = "2";
                                        }
                                        else if (panel.Machininglist[i].Face == "6")
                                        {
                                            Cao_5 = "2";
                                        }
                                    }

                                    borderseq.SawsCount = Cao_5 + ";" + Cao_6;
                                }
                                #endregion

                                if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  将切透的变成板厚+0.1的深度
                                {
                                    routeseq.RouteZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                    routeseq.RouteSetMillZ = (Convert.ToDouble(panel.Thickness) + 0.15).ToString();
                                }

                                if (panel.Machininglist[i].Face == "5")
                                {
                                    if (Math.Abs(Convert.ToDouble(panel.Machininglist[i].Depth) - Convert.ToDouble(panel.Thickness)) < 0.5)  //20180808 认为是切透的  
                                    {
                                        face5list.Add(routeseq.OutPutCsvString());
                                        borderseq.FoundRouting = "TRUE";
                                    }
                                    else
                                    {
                                        face6list.Add(routeseq.OutPutCsvString());
                                        borderseq.FoundRoutingFace6 = "TRUE";
                                    }

                                }
                                else if (panel.Machininglist[i].Face == "6")
                                {
                                    face5list.Add(routeseq.OutPutCsvString());
                                    borderseq.FoundRouting = "TRUE";
                                }

                            }
                            #endregion
                        }
                    }
                }

                // 因为有些是需要在水平孔、垂直孔判读，故在最后再增加板件信息
                face5list.Add(borderseq.OutPutCsvString());

                borderseq.FileName = "";   //至反面加工码的时候，不需要正面加工码的名字
                face6list.Add(borderseq.OutPutCsvString());

                // 因为有些是需要在水平孔、垂直孔判读，故在最后再增加板件信息
                if (ComboBox_3VJ_SMAX.SelectedIndex - 0 < 0.1) //如果选项卡上选择的是索引号为0,则为原先与普实对接的，为1则为与SMAX对接
                {
                    //csvname = csvname = panel.ID.Substring(0, 6) + panel.ID.Substring(7, 3) + "X";
                    //csvname = borderseq.FileName;
                    csvname = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "X";
                    OutFace5Face6Csv(face5list, csvname, double.Parse(borderseq.PanelThickness));   // 20180418
                                                                                                    //csvname = csvname = panel.ID.Substring(0, 6) + panel.ID.Substring(7, 3) + "Y";
                                                                                                    //csvname = borderseq.Face6FileName;
                    csvname = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "Y";
                    OutFace5Face6Csv(face6list, csvname, double.Parse(borderseq.PanelThickness));  //20180418
                }
                else
                {
                    //csvname = csvname = panel.ID.Substring(0, 6) + panel.ID.Substring(7, 3) + "X";
                    //csvname = borderseq.FileName;
                    csvname = EveryNum + "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "X";
                    OutFace5Face6Csv(face5list, csvname, double.Parse(borderseq.PanelThickness));   // 20180418
                                                                                                    //csvname = csvname = panel.ID.Substring(0, 6) + panel.ID.Substring(7, 3) + "Y";
                                                                                                    //csvname = borderseq.Face6FileName;
                    csvname = EveryNum + "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "Y";
                    OutFace5Face6Csv(face6list, csvname, double.Parse(borderseq.PanelThickness));  //20180418
                }

                                          
            }

            if (ComboBox_3VJ_SMAX.SelectedIndex - 0 < 0.1)
            MessageBox.Show("导出成功!","信息",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        public RouteSetMillSequenceEntity RouteProcess(Machining machining, double width, double pointx , double pointy)  //20180327
        {
            RouteSetMillSequenceEntity routesetmillseq = new RouteSetMillSequenceEntity();
            routesetmillseq.RouteSetMillSequence = "RouteSetMillSequence";
            routesetmillseq.RouteSetMillX = pointx.ToString();
            routesetmillseq.RouteSetMillY = pointy.ToString();
            routesetmillseq.RouteSetMillZ = machining.Depth;
            routesetmillseq.RouteStartOffsetX = pointx.ToString();
            if (machining.ToolName == "开料铣型刀")
            {
                if (width < 6)  //直径为3mm的刀
                {
                    routesetmillseq.RouteDiameter = "3";
                    routesetmillseq.RouteToolName = "D3";
                }
                else
                {
                    routesetmillseq.RouteDiameter = "8.5";
                    routesetmillseq.RouteToolName = "129";
                }
            }
            else if (machining.ToolName == "开槽刀")
            {
                routesetmillseq.RouteDiameter = "8.5";
                routesetmillseq.RouteToolName = "129";
            }
            routesetmillseq.RouteStartOffsetY = (pointy + float.Parse(routesetmillseq.RouteDiameter)/2).ToString();
            routesetmillseq.RoutePreviousToolName = "";
            routesetmillseq.RouteNextToolName = "";
            routesetmillseq.RouteFeedSpeed = "";
            routesetmillseq.RouteEntrySpeed = "";
            routesetmillseq.RouteBitType = "";
            routesetmillseq.RouteRotation = "";

            if (machining.ToolOffset == " ") // 宋新刚 20180326
                routesetmillseq.RouteToolComp = "0";
            else if (machining.ToolOffset == "左")
                routesetmillseq.RouteToolComp = "1";
            else if (machining.ToolOffset == "右")
                routesetmillseq.RouteToolComp = "2";

            routesetmillseq.RouteX = pointx.ToString();    //实则上要依据刀具偏置换算
            routesetmillseq.RouteY = pointy.ToString();     //实则上要依据刀具偏置换算
            routesetmillseq.RouteZ = machining.Depth;
            routesetmillseq.RouteEndOffsetX = "";
            routesetmillseq.RouteEndOffsetY = "";
            routesetmillseq.RouteBulge = "0";
            routesetmillseq.RouteRadius = "";   //如果拱高比不为0 则必须要此值
            routesetmillseq.RouteCenterX = "";  //如果拱高比不为0 则必须要此值
            routesetmillseq.RouteCenterY = "";  //如果拱高比不为0 则必须要此值
            routesetmillseq.RouteNextX = "";
            routesetmillseq.RouteNextY = "";
            routesetmillseq.RoutePreviousX = "";
            routesetmillseq.RoutePreviousY = "";
            routesetmillseq.RoutePreviousZ = "";
            routesetmillseq.RouteBulgeNext = "";
            routesetmillseq.RouteSetMillCounter = "";  //AD里是要的 为了标记连续轮廓
            routesetmillseq.RouteVectorCounter = "";
            routesetmillseq.RouteVectorCount = "";
            routesetmillseq.RouteAngle = "";
            routesetmillseq.RoutePreviousFeedSpeed = "";
            routesetmillseq.ArcPeakX = "";
            routesetmillseq.ArcPeakY = "";
            routesetmillseq.ArcPeakBulge = "";
            routesetmillseq.RoutePreviousBulge = "";
            routesetmillseq.RouteRotation = "";
            routesetmillseq.RouteSpindleSpeed = "";
            routesetmillseq.RouteStartTangentX = "";
            routesetmillseq.RouteStartTangentY = "";
            routesetmillseq.RouteEndTangentX = "";
            routesetmillseq.RouteEndTangentY = "";  //记录是否需要减封边值

            return (routesetmillseq);
        }

        public RouteSetMillSequenceEntity RouteProcess(Machining machining, int times)  //20181030
        {
            RouteSetMillSequenceEntity routesetmillseq = new RouteSetMillSequenceEntity();

            string RouteSetMillX = (Convert.ToDouble(machining.X) - Convert.ToDouble(machining.Diameter) / 2).ToString();

            if (times == 1)
            {
                routesetmillseq.RouteSetMillSequence = "RouteSetMillSequence";
                routesetmillseq.RouteX = RouteSetMillX;
            }
            else if (times == 2)
            {
                routesetmillseq.RouteSetMillSequence = "RouteSequence";
                routesetmillseq.RouteX = (Convert.ToDouble(machining.X) + Convert.ToDouble(machining.Diameter) / 2).ToString();
            }
            else if (times == 3)
            {
                routesetmillseq.RouteSetMillSequence = "RouteSequence";
                routesetmillseq.RouteX = RouteSetMillX;
            }

            routesetmillseq.RouteSetMillX = RouteSetMillX;
            routesetmillseq.RouteSetMillY = machining.Y;
            routesetmillseq.RouteSetMillZ = (Convert.ToDouble(machining.Depth) + 0.15).ToString();
            // routesetmillseq.RouteStartOffsetX = (Convert.ToDouble(RouteSetMillX) + 5).ToString();
            routesetmillseq.RouteStartOffsetX = RouteSetMillX;
            routesetmillseq.RouteStartOffsetY = machining.Y;
            routesetmillseq.RouteDiameter = "10";
            routesetmillseq.RouteToolName = "130";
            routesetmillseq.RoutePreviousToolName = "";
            routesetmillseq.RouteNextToolName = "";
            routesetmillseq.RouteFeedSpeed = "";
            routesetmillseq.RouteEntrySpeed = "";
            routesetmillseq.RouteBitType = "";
            routesetmillseq.RouteRotation = "";
            routesetmillseq.RouteToolComp = "1";

            routesetmillseq.RouteY = machining.Y;
            routesetmillseq.RouteZ = (Convert.ToDouble(machining.Depth) + 0.15).ToString();
            routesetmillseq.RouteEndOffsetX = "";
            routesetmillseq.RouteEndOffsetY = "";
            routesetmillseq.RouteBulge = "-1";
            routesetmillseq.RouteRadius = (Convert.ToDouble(machining.Diameter) / 2).ToString();
            routesetmillseq.RouteCenterX = machining.X;
            routesetmillseq.RouteCenterY = machining.Y;
            routesetmillseq.RouteNextX = "";
            routesetmillseq.RouteNextY = "";
            routesetmillseq.RoutePreviousX = "";
            routesetmillseq.RoutePreviousY = "";
            routesetmillseq.RoutePreviousZ = "";
            routesetmillseq.RouteBulgeNext = "";
            routesetmillseq.RouteSetMillCounter = "";  //AD里是要的 为了标记连续轮廓
            routesetmillseq.RouteVectorCounter = "";
            routesetmillseq.RouteVectorCount = "";
            routesetmillseq.RouteAngle = "";
            routesetmillseq.RoutePreviousFeedSpeed = "";
            routesetmillseq.ArcPeakX = "";
            routesetmillseq.ArcPeakY = "";
            routesetmillseq.ArcPeakBulge = "";
            routesetmillseq.RoutePreviousBulge = "";
            routesetmillseq.RouteRotation = "";
            routesetmillseq.RouteSpindleSpeed = "";
            routesetmillseq.RouteStartTangentX = "";
            routesetmillseq.RouteStartTangentY = "";
            routesetmillseq.RouteEndTangentX = "";
            routesetmillseq.RouteEndTangentY = "1";  //记录是否需要减封边值

            return (routesetmillseq);
        }

        private bool twopointisHorVline(List<fourpoint> pointxy1, List<fourpoint> pointxy2)
        {
            try
            {
                double px1 = 0;
                double py1 = 0;
                double px2 = 0;
                double py2 = 0;

                foreach(fourpoint fp in pointxy1)
                {
                    px1 = fp.x;
                    py1 = fp.y;
                }

                foreach (fourpoint fp in pointxy2)
                {
                    px2 = fp.x;
                    py2 = fp.y;
                }

                double angle = 0;
                double x = px2 - px1;
                double y = py2 - py1;

                double hypotenuse = Math.Sqrt(Math.Pow(x, 2) + Math.Pow(y, 2));

                double cos = x / hypotenuse;
                double radian = Math.Acos(cos);

                angle = 180 / (Math.PI / radian);

                if (y < 0)
                {
                    angle = -angle;
                }
                else if ((y == 0) && (x < 0))
                {
                    angle = 180;
                }

                if (angle < 0)
                    angle = angle + 360;

                if (angle == 0 || angle == 90 || angle == 180 || angle == 270)
                    return true;
                else 
                    return false;
            }
            catch
            {
                throw new NotImplementedException();
            }

        }

        private bool Isforpoint(List<fourpoint> point4, fourpoint fourpoint)
        {
            try
            {
                foreach(fourpoint FPT in point4)
                {
                    if (Math.Abs(FPT.x - fourpoint.x) < 0.01 && Math.Abs(FPT.y - fourpoint.y) < 0.01)
                        return true;
                }
                return false;
            }
            catch
            {
                throw new NotImplementedException();
            }

        }

        public void OutFace5Face6Csv(ArrayList list,string csvname,double Thickness)
        {
            ArrayList orderBorderSequence = new ArrayList();
            ArrayList orderHDrillSequence = new ArrayList();
            ArrayList orderVdrillSequence = new ArrayList();
            ArrayList orderRouteSequence = new ArrayList();
            ArrayList orderRouteSequencemodify = new ArrayList();
            ArrayList calculate = new ArrayList();

            foreach (string str in list)
            {
                if (str.Contains("BorderSequence"))
                    orderBorderSequence.Add(str);
                else if (str.Contains("HDrillSequence"))
                    orderHDrillSequence.Add(str);
                else if (str.Contains("VdrillSequence"))
                    orderVdrillSequence.Add(str);
                else
                    orderRouteSequence.Add(str);

            }

            if (orderHDrillSequence.Count == 0 && orderVdrillSequence.Count == 0 && orderRouteSequence.Count == 0)
                return;

            string path = Path.Combine(csvpath, csvname + ".csv");

            StreamWriter sw = new StreamWriter(path, false, Encoding.Default);

            foreach (string str in orderBorderSequence)
            {
                sw.WriteLine(str);
            }

            foreach (string str in orderHDrillSequence)
            {
                sw.WriteLine(str);
            }

            foreach (string str in orderVdrillSequence)
            {
                sw.WriteLine(str);
            }

            //foreach (string str in orderRouteSequence)
            //{
            //    sw.WriteLine(str);
            //}

            #region 统计一组铣型中一共有多少个拐点同时按照顺序填加进index容器中 宋新刚 20180320
            ArrayList index = new ArrayList();
            int num = 0;
            foreach (string str in orderRouteSequence)
            {
                if (str.Contains("RouteSetMillSequence"))
                {
                    index.Add(num.ToString());
                    num = 0;
                }
                else if (str.Contains("RouteSequence"))
                {
                    num++;
                }
            }
            index.Add(num.ToString());
            #endregion

            #region 在原数据中 增加表格AE、AF、AG列的值  宋新刚 20180320
            int RouteSetMillCounterNum = 0;
            int RouteVectorCounterNum = 0;
            int k = 0;
            foreach (string str in orderRouteSequence)
            {
                RouteSetMillSequenceEntity fororder = new RouteSetMillSequenceEntity(str);

                if (fororder.RouteSetMillSequence.StartsWith("RouteSetMillSequence"))
                {
                    RouteSetMillCounterNum++;
                    RouteVectorCounterNum = 0;
                    k++;
                }

                if (fororder.RouteSetMillSequence.StartsWith("RouteSequence"))
                {
                    RouteVectorCounterNum++;
                }

                fororder.RouteSetMillCounter = RouteSetMillCounterNum.ToString();
                fororder.RouteVectorCounter = RouteVectorCounterNum.ToString();
                fororder.RouteVectorCount = index[k].ToString();
                orderRouteSequencemodify.Add(fororder.OutPutCsvString());
            }
            #endregion

            #region 三维家里做异型是不减封边的值的。在这里减去封边的值，默认为1mm   20180328

            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            for (int m = 0 ;m < orderRouteSequencemodify.Count;m++)
            {
                string line = orderRouteSequencemodify[m].ToString();
                RouteSetMillSequenceEntity subvalue = new RouteSetMillSequenceEntity(line);

                if (subvalue.RouteDiameter == "10" && Thickness - 8 > 0.1 && subvalue.RouteEndTangentY != "1")   //20180418  //20180726 增加减不减封边的识别
                {
                    if (line.StartsWith("RouteSetMillSequence"))
                    {
                        if (m != 0)
                        {
                            string line2 = orderRouteSequencemodify[m - 1].ToString();
                            RouteSetMillSequenceEntity subvalue2 = new RouteSetMillSequenceEntity(line2);
                            if (subvalue2.RouteDiameter == "10" && subvalue2.RouteEndTangentY != "1")  //上一段铣型也必须要是用10mm的刀具  宋新刚 20180326
                            {
                                subvalue2.RouteX = x2.ToString();
                                subvalue2.RouteY = y2.ToString();
                                if (subvalue2.RouteBulge != "0")
                                {
                                    subvalue2.RouteRadius = (Convert.ToDouble(subvalue2.RouteRadius) + 1).ToString();
                                }
                                calculate.Add(subvalue2.OutPutCsvString());
                            }
                        }

                        x1 = Convert.ToDouble(subvalue.RouteX);
                        y1 = Convert.ToDouble(subvalue.RouteY);
                        continue;
                    }

                   x2 = Convert.ToDouble(subvalue.RouteX);
                   y2 = Convert.ToDouble(subvalue.RouteY);

                    if ((x2 - x1) > 0 && (y2 - y1) > 0)
                    {
                        if (!subvalue.RouteBulge.StartsWith("-"))
                        {
                            x1 = x1 - 1;
                            y2 = y2 + 1;
                        }
                        else
                        {
                            y1 = y1 + 1;
                            x2 = x2 - 1;
                        }
                    }
                    else if ((x2 - x1) > 0 && (y2 - y1) < 0)
                    {
                        if (!subvalue.RouteBulge.StartsWith("-"))
                        {
                            y1 = y1 + 1;
                            x2 = x2 + 1;
                        }
                        else
                        {
                            x1 = x1 + 1;
                            y2 = y2 + 1;
                        }

                    }
                    else if ((x2 - x1) < 0 && (y2 - y1) > 0)
                    {
                        if (!subvalue.RouteBulge.StartsWith("-"))
                        {
                            y1 = y1 - 1;
                            x2 = x2 - 1;
                        }
                        else
                        {
                            x1 = x1 - 1;
                            y2 = y2 - 1;
                        }
                    }
                    else if ((x2 - x1) < 0 && (y2 - y1) < 0) 
                    {
                        if (!subvalue.RouteBulge.StartsWith("-")) //未验算
                        {
                            x1 = x1 + 1;
                            y2 = y2 - 1;
                        }
                        else
                        {
                            y1 = y1 - 1;
                            x2 = x2 + 1;
                        }

                    }
                    else if ((x2 - x1) == 0 && (y2 - y1) > 0)
                    {
                        x1 = x1 - 1;
                        x2 = x2 - 1;
                    }
                    else if ((x2 - x1) > 0 && (y2 - y1) == 0)
                    {
                        if (subvalue.RouteBulge == "1")
                        {
                            x1 = x1 - 1;
                            x2 = x2 + 1;
                        }
                        else
                        {
                            y1 = y1 + 1;
                            y2 = y2 + 1;
                        }

                    }
                    else if ((x2 - x1) < 0 && (y2 - y1) == 0)
                    {
                        if (subvalue.RouteBulge == "1")
                        {
                            x1 = x1 + 1;
                            x2 = x2 - 1;
                        }
                        else
                        {
                            y1 = y1 - 1;
                            y2 = y2 - 1;
                        }

                    }
                    else if ((x2 - x1) == 0 && (y2 - y1) < 0)
                    {
                        x1 = x1 + 1;
                        x2 = x2 + 1;
                    }
                    else
                    {
                        MessageBox.Show("异形减封边不在上述的范围之内.请将XML文件准备好，与宋新刚18913812043联系!");
                    }

                    string line1 = orderRouteSequencemodify[m-1].ToString();
                    RouteSetMillSequenceEntity subvalue1 = new RouteSetMillSequenceEntity(line1);
                    subvalue1.RouteX = x1.ToString();
                    subvalue1.RouteY = y1.ToString();
                    subvalue1.RouteSetMillX = x1.ToString();
                    subvalue1.RouteSetMillY = y1.ToString();
                    if (subvalue1.RouteBulge != "0")
                    {
                        subvalue1.RouteRadius = (Convert.ToDouble(subvalue1.RouteRadius) + 1).ToString();
                    }
                    calculate.Add(subvalue1.OutPutCsvString());

                    if ((orderRouteSequencemodify.Count - (m + 1)) > 0.1)//20180726 如下面还有内轮廓的造型。则需要将最后的点输出！
                    {
                        string line111 = orderRouteSequencemodify[m + 1].ToString();
                        RouteSetMillSequenceEntity subvalue111 = new RouteSetMillSequenceEntity(line111);
                        if (subvalue111.RouteEndTangentY == "1")  
                        {
                            subvalue.RouteX = x2.ToString();
                            subvalue.RouteY = y2.ToString();
                            if (subvalue.RouteBulge != "0")
                            {
                                subvalue.RouteRadius = (Convert.ToDouble(subvalue.RouteRadius) + 1).ToString();
                            }
                            calculate.Add(subvalue.OutPutCsvString());
                        }
                    }

                    if ((orderRouteSequencemodify.Count - (m + 1)) < 0.1)
                    {
                        subvalue.RouteX = x2.ToString();
                        subvalue.RouteY = y2.ToString();
                        if (subvalue.RouteBulge != "0")
                        {
                            subvalue.RouteRadius = (Convert.ToDouble(subvalue.RouteRadius) + 1).ToString();
                        }
                        calculate.Add(subvalue.OutPutCsvString());
                    }
                    else
                    {
                        if (subvalue.RouteBulge == "0" && subvalue1.RouteBulge == "0")  //当前为拱高为0 上一拱高也为0 可得出是连续直线
                        {
                            string line3 = orderRouteSequencemodify[m + 1].ToString();
                            RouteSetMillSequenceEntity subvalue3 = new RouteSetMillSequenceEntity(line3);

                            if (subvalue3.RouteBulge == "0")  //下一段线的拱高是不是为0
                            {
                                x1 = x2;
                                y1 = y2;
                            }
                            else
                            {
                                x1 = Convert.ToDouble(subvalue.RouteX);
                                y1 = Convert.ToDouble(subvalue.RouteY);
                            }

                        }
                        else  //中间存在圆弧
                        {
                            string line3 = orderRouteSequencemodify[m + 1].ToString();
                            RouteSetMillSequenceEntity subvalue3 = new RouteSetMillSequenceEntity(line3);

                            if (subvalue1.RouteBulge == "1")
                            {
                                x1 = x2;
                                y1 = y2;
                            }
                            else if (subvalue1.RouteBulge.StartsWith("-") && subvalue.RouteBulge == "0" && subvalue3.RouteBulge == "0")   //上一拱高为负 当前拱高为0 下一拱高为0   宋新刚 20180407
                            {
                                x1 = x2;
                                y1 = y2;
                            }
                            else
                            {
                                x1 = Convert.ToDouble(subvalue.RouteX);
                                y1 = Convert.ToDouble(subvalue.RouteY);
                            }

                        }

                    }

                } 
                else
                {
                    subvalue.RouteEndTangentY = ""; //20180726 如果不减封边，将此值再为空
                    calculate.Add(subvalue.OutPutCsvString());
                }
            }
            #endregion

            foreach (string str in calculate)
            {
                sw.WriteLine(str);
            }

            sw.WriteLine("EndSequence,,,,");
            sw.Flush();
            sw.Close();
        }


        ArrayList list = new ArrayList();
        private void button3_Click(object sender, EventArgs e)
        {
            #region 订单号的创建 20180412
            string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
            string orderno = "";
            IniFiles inifile = new IniFiles(inipath);

            if (inifile.ExistINIFile())
            {
                orderno = inifile.IniReadValue("OrderNo", "Order");
                DateTime dt = DateTime.Now;
                string nowdate = dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0');
                string olddate = orderno.Substring(6, 8);
                string orderline = orderno.Substring(14, 3);
                string fix = orderno.Substring(0, 6);
                if (nowdate == olddate)
                {
                    orderno = fix + olddate + (int.Parse(orderline) + 1).ToString().PadLeft(3, '0');
                }
                else
                {
                    orderno = fix + nowdate + "001";
                }
                inifile.IniWriteValue("OrderNo", "Order", orderno);

                string oldcsvnum = inifile.IniReadValue("CsvNum", "Num");
                inifile.IniWriteValue("CsvNum", "Num", oldcsvnum.Substring(0, 1) + (int.Parse(oldcsvnum.Substring(1, 4)) + 1).ToString().PadLeft(4, '0'));
            }
            else
            {
                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                return;
            }

            #endregion

            list.Clear();

            string nestpath = csvpath + ".csv";

            if (File.Exists(nestpath))
                File.Delete(nestpath);

            StreamWriter sw = new StreamWriter(nestpath, false, Encoding.Default);
            sw.WriteLine("序号, 材料, 封边宽1, 封边宽2, 封边长1, 封边长2, 名称, 长度, 宽度, 数量, 加工代码, 反面加工代码, 备注1(批次号), 备注2(分拣号), 备注3(板件号), 备注4(分流), 备注5(优化号), 备注6(FTP目录), 备注7(反面FTP目录), 备注8, 订单号, 行号");

            foreach (ListViewItem item in listView1.Items)
            {
                Panel panel = item.Tag as Panel;
                ClassEntity nest = new ClassEntity();
                nest.Index = panel.ID.Substring(panel.ID.Length - 3, 3);   // 20180416-1
                nest.Material = panel.Material;
                nest.Material = nest.Material.Replace("白山纹", "白色");
                nest.Material = nest.Material.Replace("珠光白", "PV9001");

                string name = panel.Material;

                #region 对颜色号的提取从配置参数里读 20181103
                string checkpathmvdata = @"\\192.168.1.20\数据源\模板忽删";
                if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))
                {
                    checkpathmvdata = @"D:\模板忽删";
                }
                string checkinimvdata = Path.Combine(checkpathmvdata, "OrderNo.ini");
                IniFiles checkinifilemvdata = new IniFiles(checkinimvdata);
                string partname = nest.Material;

                if (panel.BasicMaterial.Equals("水晶板"))
                {
                    partname = nest.Material + "SJ";
                }

                if (checkinifilemvdata.ExistINIFile())
                {
                    nest.Material = checkinifilemvdata.IniReadValue("MVDATA", partname);
                    name = nest.Material;
                }
                else
                {
                    MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                    return;
                }
                #endregion

                nest.Material = panel.Thickness + "mm" + nest.Material.Replace("月石白PW", "月石白皮纹") + "E0级刨花板";
                nest.EbW1 = "";
                nest.EbW2 = "";
                nest.EbL1 = "";
                nest.EbL2 = "";

                int i = -1;
                //foreach (var str in panel.Edgelist.Where(p => p.Thickness != "0"))
                //{
                //    i++;
                //    if (i == 0)
                //        nest.EbW1 = str.Thickness + "mm" + panel.Material.Replace("PW5101", "皮纹") + "封边条";
                //    else if (i == 1)
                //        nest.EbW2 = str.Thickness + "mm" + panel.Material.Replace("PW5101", "皮纹") + "封边条";
                //    else if (i == 2)
                //        nest.EbL1 = str.Thickness + "mm" + panel.Material.Replace("PW5101", "皮纹") + "封边条";
                //    else if (i == 3)
                //        nest.EbL2 = str.Thickness + "mm" + panel.Material.Replace("PW5101", "皮纹") + "封边条";
                //}

                foreach (var str in panel.Edgelist.Where(p => p.Thickness != "0"))  // 20180718
                {
                    i++;
                    if (i == 0)
                        nest.EbW1 = str.Thickness + "mm" + name.Replace("月石白PW", "月石白皮纹").Replace("白色", "白山纹") + "封边条";
                    else if (i == 1)
                        nest.EbW2 = str.Thickness + "mm" + name.Replace("月石白PW", "月石白皮纹").Replace("白色", "白山纹") + "封边条";
                    else if (i == 2)
                        nest.EbL1 = str.Thickness + "mm" + name.Replace("月石白PW", "月石白皮纹").Replace("白色", "白山纹") + "封边条";
                    else if (i == 3)
                        nest.EbL2 = str.Thickness + "mm" + name.Replace("月石白PW", "月石白皮纹").Replace("白色", "白山纹") + "封边条";
                }

                nest.PartName = panel.Name;
                nest.Length = panel.Length;
                nest.Width = panel.Width;
                nest.Num = "1";

                nest.F5FileName = "P" + panel.ID.Substring((panel.ID.Length - 3),3) + "X";
                nest.F6FileName = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "Y";

                string F5csv = Path.Combine(csvpath, nest.F5FileName + ".csv");
                string F6csv = Path.Combine(csvpath, nest.F6FileName + ".csv");

                if (!File.Exists(F5csv))
                {
                    if (File.Exists(F6csv))
                    {
                        ArrayList face6_face5 = new ArrayList();
                        StreamReader sr = new StreamReader(F6csv, Encoding.Default);
                        string line = "";
                        BorderSequenceEntity border = null;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.StartsWith("BorderSequence"))
                            {
                                border = new BorderSequenceEntity(line);
                                border.FoundRouting = "TRUE";
                                border.RunField = "3";
                                border.CurrentZoneName = "M";
                                border.MachinePoint = "2M";
                                border.FileName = border.Face6FileName.Replace("Y", "X");  // 20180829 转的时候增加面5的名字

                                if (border.SawsCount.Contains(";"))  //20180807 
                                {
                                    string[] a = border.SawsCount.Split(';');
                                    border.SawsCount = a[1] + ";" + a[0];
                                }
                            }
                            else if (line.StartsWith("HDrillSequence"))
                            {
                                HDrillSequenceEntity hdrill = new HDrillSequenceEntity(line);
                                face6_face5.Add(hdrill.OutPutCsvString());
                            }
                            else if (line.StartsWith("VdrillSequence"))
                            {
                                VdrillSequenceEntity vdrill = new VdrillSequenceEntity(line);
                                face6_face5.Add(vdrill.OutPutCsvString());
                                border.FoundVdrill = "TRUE";
                                border.FoundVdrillFace6 = "FALSE";
                            }
                            else if (line.StartsWith("RouteSetMillSequence") || line.StartsWith("RouteSequence"))
                            {
                                RouteSetMillSequenceEntity route = new RouteSetMillSequenceEntity(line);
                                face6_face5.Add(route.OutPutCsvString());
                                border.FoundRouting = "TRUE";
                                border.FoundRoutingFace6 = "FALSE";
                            }
                        }
                        face6_face5.Insert(0, border.OutPutCsvString());
                        sr.Close();
                        File.Delete(F6csv);

                        StreamWriter swF6_F5 = new StreamWriter(F5csv, false, Encoding.Default);

                        foreach (string str in face6_face5)
                        {
                            swF6_F5.WriteLine(str);
                        }
                        swF6_F5.WriteLine("EndSequence,,,,");
                        swF6_F5.Flush();
                        swF6_F5.Close();


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
                    else //经过标签实际发现 如果面5无垂直铣型加工。面6有垂直铣型加工。则也是需要将面6的图素转到面5上  20180412
                    {
                        ArrayList face6_face5 = new ArrayList();
                        StreamReader F5_sr = new StreamReader(F5csv, Encoding.Default);
                        string F5_line = "";
                        bool needtochange = false;
                        while ((F5_line = F5_sr.ReadLine()) != null)
                        {
                            if (F5_line.StartsWith("HDrillSequence"))
                            {
                                HDrillSequenceEntity hdrill = new HDrillSequenceEntity(F5_line);
                                hdrill.HDrillY = (double.Parse(panel.Thickness) - double.Parse(hdrill.HDrillY)).ToString();   //20180612 发现符合条件转换过来的时候，水平孔也需要减
                                face6_face5.Add(hdrill.OutPutCsvString());
                                needtochange = true;
                            }
                            else if (F5_line.StartsWith("VdrillSequence"))
                            {
                                VdrillSequenceEntity vdrill = new VdrillSequenceEntity(F5_line);
                                face6_face5.Add(vdrill.OutPutCsvString());
                                needtochange = false;
                            }
                            else if (F5_line.StartsWith("RouteSetMillSequence") || F5_line.StartsWith("RouteSequence"))
                            {
                                RouteSetMillSequenceEntity route = new RouteSetMillSequenceEntity(F5_line);
                                face6_face5.Add(route.OutPutCsvString());
                                needtochange = false;
                            }
                        }
                        F5_sr.Close();

                        if (needtochange)
                        {
                            StreamReader sr = new StreamReader(F6csv, Encoding.Default);
                            string line = "";
                            BorderSequenceEntity border = null;
                            while ((line = sr.ReadLine()) != null)
                            {
                                if (line.StartsWith("BorderSequence"))
                                {
                                    border = new BorderSequenceEntity(line);
                                    border.FoundRouting = "TRUE";
                                    border.RunField = "3";
                                    border.CurrentZoneName = "M";
                                    border.MachinePoint = "2M";
                                    border.FileName = border.Face6FileName.Replace("Y", "X");  // 20180829 转的时候增加面5的名字

                                    if (border.SawsCount.Contains(";"))  //20180807 
                                    {
                                        string[] a = border.SawsCount.Split(';');
                                        border.SawsCount = a[1] + ";" + a[0];
                                    }
                                }
                                else if (line.StartsWith("HDrillSequence"))
                                {
                                    HDrillSequenceEntity hdrill = new HDrillSequenceEntity(line);
                                    face6_face5.Add(hdrill.OutPutCsvString());
                                }
                                else if (line.StartsWith("VdrillSequence"))
                                {
                                    VdrillSequenceEntity vdrill = new VdrillSequenceEntity(line);
                                    face6_face5.Add(vdrill.OutPutCsvString());
                                    border.FoundVdrill = "TRUE";
                                    border.FoundVdrillFace6 = "FALSE";
                                }
                                else if (line.StartsWith("RouteSetMillSequence") || line.StartsWith("RouteSequence"))
                                {
                                    RouteSetMillSequenceEntity route = new RouteSetMillSequenceEntity(line);
                                    face6_face5.Add(route.OutPutCsvString());
                                    border.FoundRouting = "TRUE";
                                    border.FoundRoutingFace6 = "FALSE";
                                }
                            }
                            face6_face5.Insert(0, border.OutPutCsvString());
                            sr.Close();
                            File.Delete(F6csv);
                            File.Delete(F5csv);

                            StreamWriter swF6_F5 = new StreamWriter(F5csv, false, Encoding.Default);

                            foreach (string str in face6_face5)
                            {
                                swF6_F5.WriteLine(str);
                            }
                            swF6_F5.WriteLine("EndSequence,,,,");
                            swF6_F5.Flush();
                            swF6_F5.Close();

                            nest.F6FileName = "";
                        }

                    }
                }

                nest.BatchNum = Path.GetFileNameWithoutExtension(nestpath);
                //nest.BoxNumber = panel.cabinet.Id.Substring(0,2);
                //string[] boxnumber = panel.cabinet.Name.Split('_');

                //if (panel.cabinet.CabinetNo != "")
                //    nest.BoxNumber = panel.cabinet.CabinetNo;
                //else
                //    MessageBox.Show("有柜号为空的情况，请检查XML文件!");

                nest.BoxNumber = textBox1.Text;
                    IniFiles inifile2 = new IniFiles(inipath);
                    string fixcsvnum = inifile2.IniReadValue("CsvNum", "Num");

                    nest.PartNumber = fixcsvnum + "-" + nest.BoxNumber + "-" + nest.Index + "-" + nest.Num;

                    nest.ModelName = panel.cabinet.Name;

                    nest.NestingNumber = fixcsvnum + "P" + panel.ID.Substring(panel.ID.Length - 3, 3);   // 20180416-1

                    nest.F5FTPAdress = "ftp://139.196.188.94/cdwj/" + orderno + "/Project/";
                    nest.F6FTPAdress = panel.cabinet.RoomName;
                    nest.Nest_Num = nest.BatchNum + "-" + nest.Index;
                    //nest.Order = panel.cabinet.OrderNo;

                    nest.Order = orderno;
                    nest.LineNumber = "1";

                    list.Add(nest.OutPutCsvString());
                sw.WriteLine(nest.OutPutCsvString());
            }
            sw.Flush();
            sw.Close();
            MessageBox.Show("总的排程单生成成功,共有 " + list.Count.ToString() + " 块工件!");

            #region 压缩正反面加工码

            string newfilepath = Path.Combine(Path.GetDirectoryName(csvpath), orderno + "_" + "1" + "_" + "DMSCSV.zip");

            if (File.Exists(csvpath))
                File.Delete(csvpath);

            ZipFile zip = new ZipFile(newfilepath, Encoding.Default);
            zip.AddDirectory(csvpath);

            zip.Save();

            MessageBox.Show("备份成功!");

            #endregion

            #region 上传至FTP139.196.188.94服务器

            string ftpserver = @"ftp://139.196.188.94";
            string ftpuser = "rckj";
            string ftppassword = "123456";

            string b = "ftp://139.196.188.94/cdwj/";

            string needupfile = Path.Combine(b, orderno + "/Project/");

            string orderfolder = Path.Combine(b, orderno);

            FTPclient client = new FTPclient(ftpserver, ftpuser, ftppassword, true);

            if (!Directory.Exists(orderfolder))
                client.FtpCreateDirectory(orderfolder);

            if (!Directory.Exists(needupfile))
                client.FtpCreateDirectory(needupfile);

            string needupfile2 = needupfile +"/" + orderno + "_" + "1" + "_" + "DMSCSV.zip";

            client.Upload(newfilepath, needupfile2);

            MessageBox.Show("上传成功");
            #endregion

            #region 五金清单的导出

            StreamWriter swmetal = new StreamWriter(nestpath.Replace(".csv","_Metal.csv"), false, Encoding.Default);
            swmetal.WriteLine(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", "序号", "编码", "名称", "型号", "数量", "长", "宽", "高", "单位"));
            int ii = 0;
            foreach (ListViewItem item2 in listView2.Items)
            {
                ii++;
                Metal metal = item2.Tag as Metal;
                swmetal.WriteLine(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", ii.ToString(), metal.Id, metal.Name, metal.PartNumber, metal.Num, metal.length, metal.width, metal.height,metal.Unit));
            }
            swmetal.Flush();
            swmetal.Close();

            MessageBox.Show("五金清单生成成功!");
            #endregion


        }

        /// <summary>
        /// 拷贝文件夹
        /// </summary>
        /// <param name="srcdir"></param>
        /// <param name="desdir"></param>
        private void CopyDirectory(string srcdir, string desdir)
        {
            string folderName = srcdir.Substring(srcdir.LastIndexOf("\\") + 1);

            //string desfolderdir = desdir + "\\" + folderName;

            string desfolderdir = desdir;

            if (desdir.LastIndexOf("\\") == (desdir.Length - 1))
            {
                desfolderdir = desdir + folderName;
            }
            string[] filenames = Directory.GetFileSystemEntries(srcdir);

            foreach (string file in filenames)// 遍历所有的文件和目录
            {
                if (Directory.Exists(file))// 先当作目录处理如果存在这个目录就递归Copy该目录下面的文件
                {

                    string currentdir = desfolderdir + "\\" + file.Substring(file.LastIndexOf("\\") + 1);
                    if (!Directory.Exists(currentdir))
                    {
                        Directory.CreateDirectory(currentdir);
                    }

                    CopyDirectory(file, desfolderdir);
                }

                else // 否则直接copy文件
                {
                    string srcfileName = file.Substring(file.LastIndexOf("\\") + 1);

                    srcfileName = desfolderdir + "\\" + srcfileName;


                    if (!Directory.Exists(desfolderdir))
                    {
                        Directory.CreateDirectory(desfolderdir);
                    }


                    File.Copy(file, srcfileName);
                }
            }//foreach 
        }//function end

        public static long GetDirectoryLength(string dirPath)
        {
            //判断给定的路径是否存在,如果不存在则退出
            if (!Directory.Exists(dirPath))
                return 0;
            long len = 0;
            //定义一个DirectoryInfo对象
            DirectoryInfo di = new DirectoryInfo(dirPath);
            //通过GetFiles方法,获取di目录中的所有文件的大小
            foreach (FileInfo fi in di.GetFiles())
            {
                len += fi.Length;
            }
            //获取di中所有的文件夹,并存到一个新的对象数组中,以进行递归
            DirectoryInfo[] dis = di.GetDirectories();
            if (dis.Length > 0)
            {
                for (int i = 0; i < dis.Length; i++)
                {
                    len += GetDirectoryLength(dis[i].FullName);
                }
            }
            return len;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string path = @"\\192.168.1.20\Optimizing\";
            string pathptp160 = @"\\192.168.1.20\ptp160\";
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035"))  //如果是宋新刚的虚拟机电脑，则重新对路径赋值
            {
                path = @"C:\Users\sxg035\Desktop\MV\Optimizing\";
                pathptp160 = @"C:\Users\sxg035\Desktop\MV\ptp160\";
            }

            DirectoryInfo TheFolder = new DirectoryInfo(path);

            foreach (DirectoryInfo NextFolder in TheFolder.GetDirectories())
            {
                string sourcepath = Path.Combine(path, NextFolder.Name + "\\Machine");
                string copypath = Path.Combine(pathptp160, NextFolder.Name);

                if (Directory.Exists(copypath))
                Directory.Delete(copypath, true);

                CopyDirectory(sourcepath, copypath);
               
            }

            MessageBox.Show("PTP160文件夹数据生成成功!");

            

        }

        private void button5_Click(object sender, EventArgs e)
        {
            int num = 0;
            ClassEntity nest_P = new ClassEntity();
            string path = @"\\192.168.1.20\Optimizing\P" + Path.GetFileName(csvpath);
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035"))  //如果是宋新刚的虚拟机电脑，则重新对路径赋值
            {
                path = @"C:\Users\sxg035\Desktop\MV\Optimizing\P" + Path.GetFileName(csvpath); 
            }

            if (Directory.Exists(path))
                Directory.Delete(path,true);

            Directory.CreateDirectory(path);
            Directory.CreateDirectory(Path.Combine(path, "ERP"));
            Directory.CreateDirectory(Path.Combine(path, "Machine"));
            Directory.CreateDirectory(Path.Combine(path, "Other"));

            string nestpath = Path.Combine(path, "ERP\\") + Path.GetFileNameWithoutExtension(path) + ".csv";
            StreamWriter sw_P = new StreamWriter(nestpath, false, Encoding.Default);
            sw_P.WriteLine("序号, 材料, 封边宽1, 封边宽2, 封边长1, 封边长2, 名称, 长度, 宽度, 数量, 加工代码, 反面加工代码, 备注1(批次号), 备注2(分拣号), 备注3(板件号), 备注4(分流), 备注5(优化号), 备注6(FTP目录), 备注7(反面FTP目录), 备注8, 订单号, 行号");
            foreach (var str in list)
            {
                nest_P = new ClassEntity(str.ToString());
                string[] thickness = nest_P.Material.Split('m');
                if (double.Parse(thickness[0]) > 17.9 && double.Parse(thickness[0]) < 35.1 && double.Parse(nest_P.Width) > 60.1 && !nest_P.PartName.Contains("门") &&  !nest_P.PartName.Contains("台面") && !nest_P.PartName.Contains("水晶") && !nest_P.PartName.Contains("PET") && !nest_P.Material.Contains("SJ") && !nest_P.Material.Contains("PV"))
                {
                    num++;
                    nest_P.BatchNum = "P" + nest_P.BatchNum;
                    nest_P.Nest_Num = "P" + nest_P.Nest_Num;
                    sw_P.WriteLine(nest_P.OutPutCsvString());
                }

            }
            sw_P.Flush();
            sw_P.Close();

            if (num == 0)
            {
                MessageBox.Show("无符合要求的【P】批次板件!");
                Directory.Delete(path, true);
            }
            else
                MessageBox.Show("板件【P】生成成功,共有 " + num.ToString() + " 块工件!");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int num = 0;
            ClassEntity nest_P = new ClassEntity();
            string path = @"\\192.168.1.20\Optimizing\B" + Path.GetFileName(csvpath);
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035"))  //如果是宋新刚的虚拟机电脑，则重新对路径赋值
            {
                path = @"C:\Users\sxg035\Desktop\MV\Optimizing\B" + Path.GetFileName(csvpath);
            }

            if (Directory.Exists(path))
                Directory.Delete(path, true);

            Directory.CreateDirectory(path);
            Directory.CreateDirectory(Path.Combine(path, "ERP"));
            Directory.CreateDirectory(Path.Combine(path, "Machine"));
            Directory.CreateDirectory(Path.Combine(path, "Other"));

            string nestpath = Path.Combine(path, "ERP\\") + Path.GetFileNameWithoutExtension(path) + ".csv";
            StreamWriter sw_P = new StreamWriter(nestpath, false, Encoding.Default);
            sw_P.WriteLine("序号, 材料, 封边宽1, 封边宽2, 封边长1, 封边长2, 名称, 长度, 宽度, 数量, 加工代码, 反面加工代码, 备注1(批次号), 备注2(分拣号), 备注3(板件号), 备注4(分流), 备注5(优化号), 备注6(FTP目录), 备注7(反面FTP目录), 备注8, 订单号, 行号");
            foreach (var str in list)
            {
                nest_P = new ClassEntity(str.ToString());
                string[] thickness = nest_P.Material.Split('m');
                string[] namethickness = nest_P.PartName.Split('m');
                if (double.Parse(thickness[0]) > 0.1 && double.Parse(thickness[0]) < 8.1 && !nest_P.PartName.Contains("CT-") && !nest_P.Material.Contains("SJ") && !nest_P.Material.Contains("PV") && !System.Text.RegularExpressions.Regex.IsMatch(namethickness[0], "^\\d+$"))
                {
                    num++;
                    nest_P.BatchNum = "B" + nest_P.BatchNum;
                    nest_P.Nest_Num = "B" + nest_P.Nest_Num;
                    sw_P.WriteLine(nest_P.OutPutCsvString());
                }

            }
            sw_P.Flush();
            sw_P.Close();
            if (num == 0)
            {
                MessageBox.Show("无符合要求的【B】批次板件!");
                Directory.Delete(path, true);
            }
            else
                MessageBox.Show("背板【B】生成成功,共有 " + num.ToString() + " 块工件!");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int num = 0;
            ClassEntity nest_P = new ClassEntity();
            string path = @"\\192.168.1.20\Optimizing\J" + Path.GetFileName(csvpath);
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035"))  //如果是宋新刚的虚拟机电脑，则重新对路径赋值
            {
                path = @"C:\Users\sxg035\Desktop\MV\Optimizing\J" + Path.GetFileName(csvpath);
            }

            if (Directory.Exists(path))
                Directory.Delete(path, true);

            Directory.CreateDirectory(path);
            Directory.CreateDirectory(Path.Combine(path, "ERP"));
            Directory.CreateDirectory(Path.Combine(path, "Machine"));
            Directory.CreateDirectory(Path.Combine(path, "Other"));

            string nestpath = Path.Combine(path, "ERP\\") + Path.GetFileNameWithoutExtension(path) + ".csv";
            StreamWriter sw_P = new StreamWriter(nestpath, false, Encoding.Default);
            sw_P.WriteLine("序号, 材料, 封边宽1, 封边宽2, 封边长1, 封边长2, 名称, 长度, 宽度, 数量, 加工代码, 反面加工代码, 备注1(批次号), 备注2(分拣号), 备注3(板件号), 备注4(分流), 备注5(优化号), 备注6(FTP目录), 备注7(反面FTP目录), 备注8, 订单号, 行号");
            foreach (var str in list)
            {
                nest_P = new ClassEntity(str.ToString());
                string[] thickness = nest_P.Material.Split('m');
                if (double.Parse(nest_P.Width) < 60.1 && !nest_P.PartName.Contains("CT-") && !nest_P.Material.Contains("SJ"))
                {
                    num++;
                    nest_P.BatchNum = "J" + nest_P.BatchNum;
                    nest_P.Nest_Num = "J" + nest_P.Nest_Num;
                    sw_P.WriteLine(nest_P.OutPutCsvString());
                }

            }
            sw_P.Flush();
            sw_P.Close();
            if (num == 0)
            {
                MessageBox.Show("无符合要求的【J】批次板件!");
                Directory.Delete(path, true);
            }
            else
                MessageBox.Show("脚板【J】生成成功,共有 " + num.ToString() + " 块工件!");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int num = 0;
            ClassEntity nest_P = new ClassEntity();
            string path = @"\\192.168.1.20\Optimizing\M" + Path.GetFileName(csvpath);
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035"))  //如果是宋新刚的虚拟机电脑，则重新对路径赋值
            {
                path = @"C:\Users\sxg035\Desktop\MV\Optimizing\M" + Path.GetFileName(csvpath);
            }

            if (Directory.Exists(path))
                Directory.Delete(path, true);

            Directory.CreateDirectory(path);
            Directory.CreateDirectory(Path.Combine(path, "ERP"));
            Directory.CreateDirectory(Path.Combine(path, "Machine"));
            Directory.CreateDirectory(Path.Combine(path, "Other"));

            string nestpath = Path.Combine(path, "ERP\\") + Path.GetFileNameWithoutExtension(path) + ".csv";
            StreamWriter sw_P = new StreamWriter(nestpath, false, Encoding.Default);
            sw_P.WriteLine("序号, 材料, 封边宽1, 封边宽2, 封边长1, 封边长2, 名称, 长度, 宽度, 数量, 加工代码, 反面加工代码, 备注1(批次号), 备注2(分拣号), 备注3(板件号), 备注4(分流), 备注5(优化号), 备注6(FTP目录), 备注7(反面FTP目录), 备注8, 订单号, 行号");
            foreach (var str in list)
            {
                nest_P = new ClassEntity(str.ToString());
                string[] thickness = nest_P.Material.Split('m');
                if ((nest_P.PartName.Contains("门") || nest_P.PartName.Contains("台面") || double.Parse(thickness[0]) > 49.9) && !nest_P.PartName.Contains("水晶") && !nest_P.PartName.Contains("PET") && !nest_P.Material.Contains("PV") && double.Parse(nest_P.Width) > 60.1 && !nest_P.Material.Contains("SJ"))
                {
                    num++;
                    nest_P.BatchNum = "M" + nest_P.BatchNum;
                    nest_P.Nest_Num = "M" + nest_P.Nest_Num;
                    sw_P.WriteLine(nest_P.OutPutCsvString());
                }

            }
            sw_P.Flush();
            sw_P.Close();
            if (num == 0)
            {
                MessageBox.Show("无符合要求的【M】批次板件!");
                Directory.Delete(path, true);
            }
            else
                MessageBox.Show("门板【M】生成成功,共有 " + num.ToString() + " 块工件!");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            int num = 0;
            ClassEntity nest_P = new ClassEntity();
            string path = @"\\192.168.1.20\Optimizing\M" + Path.GetFileName(csvpath);
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035"))  //如果是宋新刚的虚拟机电脑，则重新对路径赋值
            {
                path = @"C:\Users\sxg035\Desktop\MV\Optimizing\S" + Path.GetFileName(csvpath);
            }

            if (Directory.Exists(path))
                Directory.Delete(path, true);

            Directory.CreateDirectory(path);
            Directory.CreateDirectory(Path.Combine(path, "ERP"));
            Directory.CreateDirectory(Path.Combine(path, "Machine"));
            Directory.CreateDirectory(Path.Combine(path, "Other"));

            string nestpath = Path.Combine(path, "ERP\\") + Path.GetFileNameWithoutExtension(path) + ".csv";
            StreamWriter sw_P = new StreamWriter(nestpath, false, Encoding.Default);
            sw_P.WriteLine("序号, 材料, 封边宽1, 封边宽2, 封边长1, 封边长2, 名称, 长度, 宽度, 数量, 加工代码, 反面加工代码, 备注1(批次号), 备注2(分拣号), 备注3(板件号), 备注4(分流), 备注5(优化号), 备注6(FTP目录), 备注7(反面FTP目录), 备注8, 订单号, 行号");
            foreach (var str in list)
            {
                nest_P = new ClassEntity(str.ToString());
                string[] thickness = nest_P.Material.Split('m');
                if ((nest_P.PartName.Contains("水晶") || nest_P.PartName.Contains("PET") || nest_P.Material.Contains("PV") || nest_P.Material.Contains("SJ")))
                {
                    num++;
                    nest_P.BatchNum = "S" + nest_P.BatchNum;
                    nest_P.Nest_Num = "S" + nest_P.Nest_Num;
                    sw_P.WriteLine(nest_P.OutPutCsvString());
                }

            }
            sw_P.Flush();
            sw_P.Close();
            if (num == 0)
            {
                MessageBox.Show("无符合要求的【S】批次板件!");
                Directory.Delete(path, true);
            }
            else
                MessageBox.Show("水晶【S】生成成功,共有 " + num.ToString() + " 块工件!");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int num = 0;
            ClassEntity nest_P = new ClassEntity();
            string path = @"\\192.168.1.20\Optimizing\M" + Path.GetFileName(csvpath);
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035"))  //如果是宋新刚的虚拟机电脑，则重新对路径赋值
            {
                path = @"C:\Users\sxg035\Desktop\MV\Optimizing\C" + Path.GetFileName(csvpath);
            }

            if (Directory.Exists(path))
                Directory.Delete(path, true);

            Directory.CreateDirectory(path);
            Directory.CreateDirectory(Path.Combine(path, "ERP"));
            Directory.CreateDirectory(Path.Combine(path, "Machine"));
            Directory.CreateDirectory(Path.Combine(path, "Other"));

            string nestpath = Path.Combine(path, "ERP\\") + Path.GetFileNameWithoutExtension(path) + ".csv";
            StreamWriter sw_P = new StreamWriter(nestpath, false, Encoding.Default);
            sw_P.WriteLine("序号, 材料, 封边宽1, 封边宽2, 封边长1, 封边长2, 名称, 长度, 宽度, 数量, 加工代码, 反面加工代码, 备注1(批次号), 备注2(分拣号), 备注3(板件号), 备注4(分流), 备注5(优化号), 备注6(FTP目录), 备注7(反面FTP目录), 备注8, 订单号, 行号");
            foreach (var str in list)
            {
                nest_P = new ClassEntity(str.ToString());
                string[] thickness = nest_P.Material.Split('m');
                if (nest_P.PartName.Contains("CT-"))
                {
                    num++;
                    nest_P.BatchNum = "C" + nest_P.BatchNum;
                    nest_P.Nest_Num = "C" + nest_P.Nest_Num;
                    sw_P.WriteLine(nest_P.OutPutCsvString());
                }

            }
            sw_P.Flush();
            sw_P.Close();
            if (num == 0)
            {
                MessageBox.Show("无符合要求的【C】批次板件!");
                Directory.Delete(path, true);
            }
            else
                MessageBox.Show("抽屉【C】生成成功,共有 " + num.ToString() + " 块工件!");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            int num = 0;
            ClassEntity nest_P = new ClassEntity();
            string path = @"\\192.168.1.20\Optimizing\M" + Path.GetFileName(csvpath);
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035"))  //如果是宋新刚的虚拟机电脑，则重新对路径赋值
            {
                path = @"C:\Users\sxg035\Desktop\MV\Optimizing\N" + Path.GetFileName(csvpath);
            }

            if (Directory.Exists(path))
                Directory.Delete(path, true);

            Directory.CreateDirectory(path);
            Directory.CreateDirectory(Path.Combine(path, "ERP"));
            Directory.CreateDirectory(Path.Combine(path, "Machine"));
            Directory.CreateDirectory(Path.Combine(path, "Other"));

            string nestpath = Path.Combine(path, "ERP\\") + Path.GetFileNameWithoutExtension(path) + ".csv";
            StreamWriter sw_P = new StreamWriter(nestpath, false, Encoding.Default);
            sw_P.WriteLine("序号, 材料, 封边宽1, 封边宽2, 封边长1, 封边长2, 名称, 长度, 宽度, 数量, 加工代码, 反面加工代码, 备注1(批次号), 备注2(分拣号), 备注3(板件号), 备注4(分流), 备注5(优化号), 备注6(FTP目录), 备注7(反面FTP目录), 备注8, 订单号, 行号");
            foreach (var str in list)
            {
                nest_P = new ClassEntity(str.ToString());
                string[] thickness = nest_P.Material.Split('m');
                string[] namethickness = nest_P.PartName.Split('m');
                if (System.Text.RegularExpressions.Regex.IsMatch(namethickness[0], "^\\d+$"))
                {
                    num++;
                    nest_P.BatchNum = "N" + nest_P.BatchNum;
                    nest_P.Nest_Num = "N" + nest_P.Nest_Num;
                    sw_P.WriteLine(nest_P.OutPutCsvString());
                }

            }
            sw_P.Flush();
            sw_P.Close();
            if (num == 0)
            {
                MessageBox.Show("无符合要求的【N】成品批次板件!");
                Directory.Delete(path, true);
            }
            else
                MessageBox.Show("成品【N】生成成功,共有 " + num.ToString() + " 块工件!");

        }

        List<SAMXPanel> Part = new List<SAMXPanel>();
        string Fisrt_Csv_Path = @"\\192.168.1.20\数据源\三维家订单正反面加工码";
        private void button12_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count == 0 || HaveLarger)
            {
                return;
            }

            Part.Clear();

            #region 订单号的创建 20180412
            string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
            string orderno = "";
            string oldcsvnum = string.Empty;
            IniFiles inifile = new IniFiles(inipath);

            if (inifile.ExistINIFile())
            {
                orderno = inifile.IniReadValue("OrderNo", "Order");
                DateTime dt = DateTime.Now;
                string nowdate = dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0');
                string olddate = orderno.Substring(6, 8);
                string orderline = orderno.Substring(14, 3);
                string fix = orderno.Substring(0, 6);

                if (nowdate == olddate)
                {
                    orderno = fix + olddate + (int.Parse(orderline) + 1).ToString().PadLeft(3, '0');
                }
                else
                {
                    orderno = fix + nowdate + "001";
                }
                inifile.IniWriteValue("OrderNo", "Order", orderno);

                oldcsvnum = inifile.IniReadValue("CsvNum", "Num");

                if (Math.Abs(int.Parse(oldcsvnum.Substring(1, 4)) - 9999) < 0.1) //如果号超过了9999 则重新归1处理
                {
                    inifile.IniWriteValue("CsvNum", "Num", oldcsvnum.Substring(0, 1) + 1.ToString().PadLeft(4, '0'));
                }
                else
                {
                    inifile.IniWriteValue("CsvNum", "Num", oldcsvnum.Substring(0, 1) + (int.Parse(oldcsvnum.Substring(1, 4)) + 1).ToString().PadLeft(4, '0'));
                }
 
            }
            else
            {
                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                return;
            }

            #endregion

            list.Clear();
            int donotneedpart = 0;
            string nestpath = csvpath + ".csv";

            if (File.Exists(nestpath))
                File.Delete(nestpath);

            StreamWriter sw = new StreamWriter(nestpath, false, Encoding.Default);
            sw.WriteLine("序号, 材料, 封边宽1, 封边宽2, 封边长1, 封边长2, 名称, 长度, 宽度, 数量, 加工代码, 反面加工代码, 备注1(批次号), 备注2(分拣号), 备注3(板件号), 备注4(分流), 备注5(优化号), 备注6(FTP目录), 备注7(反面FTP目录), 备注8, 订单号, 行号");

            foreach (ListViewItem item in listView1.Items)
            {
                Panel panel = item.Tag as Panel;
                ClassEntity nest = new ClassEntity();
                string Currentpanelcolor = string.Empty;
                nest.Index = panel.ID.Substring(panel.ID.Length - 3, 3);   // 20180416-1
                nest.Material = panel.Material;
                string name = panel.Material;

                #region 对颜色号的提取从配置参数里读 20181103
                string checkpathmvdata = @"\\192.168.1.20\数据源\模板忽删";
                if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))
                {
                    checkpathmvdata = @"D:\模板忽删";
                }
                string checkinimvdata = Path.Combine(checkpathmvdata, "OrderNo.ini");
                IniFiles checkinifilemvdata = new IniFiles(checkinimvdata);
                string partname = nest.Material;

                if (panel.BasicMaterial.Equals("水晶板"))
                {
                    partname = nest.Material + "SJ";
                }

                if (checkinifilemvdata.ExistINIFile())
                {
                    nest.Material = checkinifilemvdata.IniReadValue("MVDATA", partname);
                    name = nest.Material;
                }
                else
                {
                    MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                    return;
                }
                #endregion

                Currentpanelcolor = nest.Material;

                nest.Material = panel.Thickness + "mm" + nest.Material + "E0级刨花板";


                nest.EbW1 = "";
                nest.EbW2 = "";
                nest.EbL1 = "";
                nest.EbL2 = "";

                int i = -1;

                foreach (var str in panel.Edgelist.Where(p => p.Thickness != "0"))  // 20180718
                {
                    i++;
                    if (i == 0)
                        nest.EbW1 = str.Thickness + "mm" + name.Replace("雏菊MM0261", "白山纹SW5001") + "封边条";
                    else if (i == 1)
                        nest.EbW2 = str.Thickness + "mm" + name.Replace("雏菊MM0261", "白山纹SW5001") + "封边条";
                    else if (i == 2)
                        nest.EbL1 = str.Thickness + "mm" + name.Replace("雏菊MM0261", "白山纹SW5001") + "封边条";
                    else if (i == 3)
                        nest.EbL2 = str.Thickness + "mm" + name.Replace("雏菊MM0261", "白山纹SW5001") + "封边条";
                }

                nest.PartName = panel.Name;
                nest.Length = panel.Length;
                nest.Width = panel.Width;
                nest.Num = "1";

                nest.F5FileName = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "X";
                nest.F6FileName = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "Y";

                if (ComboBox_3VJ_SMAX.SelectedIndex - 0 > 0.1) //如果选项卡上选择的是索引号为0,则为原先与普实对接的，为1则为与SMAX对接
                {
                    nest.F5FileName = oldcsvnum + "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "X";
                    nest.F6FileName = oldcsvnum + "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "Y";
                }

                string F5csv = Path.Combine(csvpath, nest.F5FileName + ".csv");
                string F6csv = Path.Combine(csvpath, nest.F6FileName + ".csv");

                if (!File.Exists(F5csv))
                {
                    if (File.Exists(F6csv))
                    {
                        ArrayList face6_face5 = new ArrayList();
                        StreamReader sr = new StreamReader(F6csv, Encoding.Default);
                        string line = "";
                        BorderSequenceEntity border = null;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.StartsWith("BorderSequence"))
                            {
                                border = new BorderSequenceEntity(line);
                                border.FoundRouting = "TRUE";
                                border.RunField = "3";
                                border.CurrentZoneName = "M";
                                border.MachinePoint = "2M";
                                border.FileName = border.Face6FileName.Replace("Y", "X");  // 20180829 转的时候增加面5的名字

                                if (border.SawsCount.Contains(";"))  //20180807 
                                {
                                    string[] a = border.SawsCount.Split(';');
                                    border.SawsCount = a[1] + ";" + a[0];
                                }
                            }
                            else if (line.StartsWith("HDrillSequence"))
                            {
                                HDrillSequenceEntity hdrill = new HDrillSequenceEntity(line);
                                face6_face5.Add(hdrill.OutPutCsvString());
                            }
                            else if (line.StartsWith("VdrillSequence"))
                            {
                                VdrillSequenceEntity vdrill = new VdrillSequenceEntity(line);
                                face6_face5.Add(vdrill.OutPutCsvString());
                                border.FoundVdrill = "TRUE";
                                border.FoundVdrillFace6 = "FALSE";
                            }
                            else if (line.StartsWith("RouteSetMillSequence") || line.StartsWith("RouteSequence"))
                            {
                                RouteSetMillSequenceEntity route = new RouteSetMillSequenceEntity(line);
                                face6_face5.Add(route.OutPutCsvString());
                                border.FoundRouting = "TRUE";
                                border.FoundRoutingFace6 = "FALSE";
                            }
                        }
                        face6_face5.Insert(0, border.OutPutCsvString());
                        sr.Close();
                        File.Delete(F6csv);

                        StreamWriter swF6_F5 = new StreamWriter(F5csv, false, Encoding.Default);

                        foreach (string str in face6_face5)
                        {
                            swF6_F5.WriteLine(str);
                        }
                        swF6_F5.WriteLine("EndSequence,,,,");
                        swF6_F5.Flush();
                        swF6_F5.Close();


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
                    else //经过标签实际发现 如果面5无垂直铣型加工。面6有垂直铣型加工。则也是需要将面6的图素转到面5上  20180412
                    {
                        ArrayList face6_face5 = new ArrayList();
                        StreamReader F5_sr = new StreamReader(F5csv, Encoding.Default);
                        string F5_line = "";
                        bool needtochange = false;
                        while ((F5_line = F5_sr.ReadLine()) != null)
                        {
                            if (F5_line.StartsWith("HDrillSequence"))
                            {
                                HDrillSequenceEntity hdrill = new HDrillSequenceEntity(F5_line);
                                hdrill.HDrillY = (double.Parse(panel.Thickness) - double.Parse(hdrill.HDrillY)).ToString();   //20180612 发现符合条件转换过来的时候，水平孔也需要减
                                face6_face5.Add(hdrill.OutPutCsvString());
                                needtochange = true;
                            }
                            else if (F5_line.StartsWith("VdrillSequence"))
                            {
                                VdrillSequenceEntity vdrill = new VdrillSequenceEntity(F5_line);
                                face6_face5.Add(vdrill.OutPutCsvString());
                                needtochange = false;
                            }
                            else if (F5_line.StartsWith("RouteSetMillSequence") || F5_line.StartsWith("RouteSequence"))
                            {
                                RouteSetMillSequenceEntity route = new RouteSetMillSequenceEntity(F5_line);
                                face6_face5.Add(route.OutPutCsvString());
                                needtochange = false;
                            }
                        }
                        F5_sr.Close();

                        if (needtochange)
                        {
                            StreamReader sr = new StreamReader(F6csv, Encoding.Default);
                            string line = "";
                            BorderSequenceEntity border = null;
                            while ((line = sr.ReadLine()) != null)
                            {
                                if (line.StartsWith("BorderSequence"))
                                {
                                    border = new BorderSequenceEntity(line);
                                    border.FoundRouting = "TRUE";
                                    border.RunField = "3";
                                    border.CurrentZoneName = "M";
                                    border.MachinePoint = "2M";
                                    border.FileName = border.Face6FileName.Replace("Y", "X");  // 20180829 转的时候增加面5的名字

                                    if (border.SawsCount.Contains(";"))  //20180807 
                                    {
                                        string[] a = border.SawsCount.Split(';');
                                        border.SawsCount = a[1] + ";" + a[0];
                                    }
                                }
                                else if (line.StartsWith("HDrillSequence"))
                                {
                                    HDrillSequenceEntity hdrill = new HDrillSequenceEntity(line);
                                    face6_face5.Add(hdrill.OutPutCsvString());
                                }
                                else if (line.StartsWith("VdrillSequence"))
                                {
                                    VdrillSequenceEntity vdrill = new VdrillSequenceEntity(line);
                                    face6_face5.Add(vdrill.OutPutCsvString());
                                    border.FoundVdrill = "TRUE";
                                    border.FoundVdrillFace6 = "FALSE";
                                }
                                else if (line.StartsWith("RouteSetMillSequence") || line.StartsWith("RouteSequence"))
                                {
                                    RouteSetMillSequenceEntity route = new RouteSetMillSequenceEntity(line);
                                    face6_face5.Add(route.OutPutCsvString());
                                    border.FoundRouting = "TRUE";
                                    border.FoundRoutingFace6 = "FALSE";
                                }
                            }
                            face6_face5.Insert(0, border.OutPutCsvString());
                            sr.Close();
                            File.Delete(F6csv);
                            File.Delete(F5csv);

                            StreamWriter swF6_F5 = new StreamWriter(F5csv, false, Encoding.Default);

                            foreach (string str in face6_face5)
                            {
                                swF6_F5.WriteLine(str);
                            }
                            swF6_F5.WriteLine("EndSequence,,,,");
                            swF6_F5.Flush();
                            swF6_F5.Close();

                            nest.F6FileName = "";
                        }

                    }
                }

                nest.BatchNum = Path.GetFileNameWithoutExtension(nestpath);
                //nest.BoxNumber = panel.cabinet.Id.Substring(0,2);
                //string[] boxnumber = panel.cabinet.Name.Split('_');

                //if (panel.cabinet.CabinetNo != "")
                //    nest.BoxNumber = panel.cabinet.CabinetNo;
                //else
                //    MessageBox.Show("有柜号为空的情况，请检查XML文件!");

                nest.BoxNumber = textBox1.Text;
                IniFiles inifile2 = new IniFiles(inipath);
                string fixcsvnum = inifile2.IniReadValue("CsvNum", "Num");

                nest.PartNumber = fixcsvnum + "-" + nest.BoxNumber + "-" + nest.Index + "-" + nest.Num;

                nest.ModelName = panel.cabinet.Name;

                nest.NestingNumber = fixcsvnum + "P" + panel.ID.Substring(panel.ID.Length - 3, 3);   // 20180416-1

                nest.F5FTPAdress = "ftp://139.196.188.94/cdwj/" + orderno + "/Project/";
                nest.F6FTPAdress = panel.cabinet.RoomName;
                nest.Nest_Num = nest.BatchNum + "-" + nest.Index;
                //nest.Order = panel.cabinet.OrderNo;

                nest.Order = orderno;
                nest.LineNumber = "1";

                list.Add(nest.OutPutCsvString());
                sw.WriteLine(nest.OutPutCsvString());


                if (ComboBox_3VJ_SMAX.SelectedIndex - 0 > 0.1) //如果选项卡上选择的是索引号为0,则为原先与普实对接的，为1则为与SMAX对接
                {
                    bool istworoompart = false;

                    foreach(var str in parttype)//判断哪些板需要加工 20180918
                    {
                        if (str.PartNumber.Equals(panel.PartNumber))
                        {
                            istworoompart = true;
                            break;
                        }
                    }

                    if (!istworoompart)
                    {
                        donotneedpart++;
                        continue;
                    }

                    SAMXPanel sampanel = new SAMXPanel();

                    sampanel.Index = nest.Index;
                    sampanel.ColorSku = Regex.Replace(name, @"[\u4e00-\u9fa5]", ""); //去除中文汉字

                    #region 当前板件颜色厚度是否在加工范围里的检查 20181014
                    string checkpath = @"\\192.168.1.20\数据源\模板忽删";
                    if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))
                    {
                        checkpath = @"D:\模板忽删";
                    }
                    string checkinipath = Path.Combine(checkpath, "OrderNo.ini");
                    IniFiles checkinifile = new IniFiles(checkinipath);
                    bool Workable = false;
                    string ThicknessShow = string.Empty;
                    if (checkinifile.ExistINIFile())
                    {
                        ThicknessShow = checkinifile.IniReadValue("PanelThickness", Currentpanelcolor);
                        string[] Thickness = ThicknessShow.Split(',');
                        for (int thicknessnum = 0; thicknessnum < Thickness.Length; thicknessnum++)
                        {
                            if (panel.Thickness.Equals(Thickness[thicknessnum]))
                            {
                                Workable = true;
                                break;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                        return;
                    }

                    if (!Workable)
                    {
                        if (string.IsNullOrEmpty(Currentpanelcolor))
                        {                
                            MessageBox.Show("三维家的板件名称为: " + panel.Name + "\n\n板件ID号为: " + panel.ID + "\n\n此板号的颜色为: " + partname.Replace("SJ","水晶板") + " 厚度为: " + panel.Thickness
                                + "\n\n请注意此颜色的板不在我们班尔奇可加工的范围内!\n\n程序终止,请修改后重新执行!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            MessageBox.Show("三维家的板件名称为: " + panel.Name + "\n\n板件ID号为: " + panel.ID + "\n\n此板号的颜色为: " + Currentpanelcolor + " 厚度为: " + panel.Thickness
                                + "\n\n请注意此颜色厚度不在可加工的 " + ThicknessShow + " 范围内!\n\n程序终止,请修改后重新执行!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        return;
                    }

                    #endregion

                    #region 对接SAMEX增加对板件可加工范围的限制
                    if (ComboBox_3VJ_SMAX.SelectedIndex - 0 > 0.1)
                    {
                        double panellength = Convert.ToDouble(panel.Length);
                        double panelwidth = Convert.ToDouble(panel.Width);
                        double panelthickness = Convert.ToDouble(panel.Thickness);

                        if (Math.Abs(panelthickness - 8) < 0.1)
                        {
                            if (panellength > 2420)
                            {
                                HaveLarger = true;
                                MessageBox.Show("三维家XML中的板件名称为: " + panel.Name + " ID号为: " + panel.ID + "\n板件长度【顺着纹理的方向】: " + panel.Length + "\n板件宽度【垂直纹理的方向】: " + panel.Width + "\n厚度为" + panel.Thickness + "mm的板在板件长度【顺着纹理的方向】超过了可加工尺寸的2420\n请先修改模型!\n\n板件正反面加工码的生成已终止!", "警告",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                            else if (panelwidth > 1200)
                            {
                                HaveLarger = true;
                                MessageBox.Show("三维家XML中的板件名称为: " + panel.Name + " ID号为: " + panel.ID + "\n板件长度【顺着纹理的方向】: " + panel.Length + "\n板件宽度【垂直纹理的方向】: " + panel.Width + "\n厚度为" + panel.Thickness + "mm的板在板件宽度【垂直纹理的方向】超过了可加工尺寸的1200\n请先修改模型!\n\n板件正反面加工码的生成已终止!", "警告",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                        else
                        {
                            if (panellength > 2720)
                            {
                                HaveLarger = true;
                                MessageBox.Show("三维家XML中的板件名称为: " + panel.Name + " ID号为: " + panel.ID + "\n板件长度【顺着纹理的方向】: " + panel.Length + "\n板件宽度【垂直纹理的方向】: " + panel.Width + "\n厚度为" + panel.Thickness + "mm的板在板件长度【顺着纹理的方向】超过了可加工尺寸的2720\n请先修改模型!\n\n板件正反面加工码的生成已终止!", "警告",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                            else if (panelwidth > 1200)
                            {
                                HaveLarger = true;
                                MessageBox.Show("三维家XML中的板件名称为: " + panel.Name + " ID号为: " + panel.ID + "\n板件长度【顺着纹理的方向】: " + panel.Length + "\n板件宽度【垂直纹理的方向】: " + panel.Width + "\n厚度为" + panel.Thickness + "mm的板在板件宽度【垂直纹理的方向】超过了可加工尺寸的1200\n请先修改模型!\n\n板件正反面加工码的生成已终止!", "警告",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }

                        #region 20181030 50mm厚的板做长宽规格做了限制，只生成标准的几种。规格写在配置参数中
                        if (Math.Abs(panelthickness - 50) < 0.1)
                        {
                            string checkinipath1 = Path.Combine(checkpath, "OrderNo.ini");
                            IniFiles checkinifile1 = new IniFiles(checkinipath1);
                            bool LengthWorkable = false;
                            bool WidthWorkable = false;

                            if (checkinifile1.ExistINIFile())
                            {
                                string[] LengthNum = checkinifile.IniReadValue("PanelLengthWidthfor50", "Length").Split(',');
                                string[] WidthNum = checkinifile.IniReadValue("PanelLengthWidthfor50", "Width").Split(',');
                                for (int num = 0; num < LengthNum.Length; num++)
                                {
                                    if (Math.Abs(Convert.ToDouble(panel.Length) - Convert.ToDouble(LengthNum[num])) < 0.1)
                                    {
                                        LengthWorkable = true;
                                    }
                                }

                                if (!LengthWorkable)
                                {
                                    MessageBox.Show("三维家XML中的板件名称为: " + panel.Name + " ID号为: " + panel.ID + "\n板件长度【顺着纹理的方向】: " + panel.Length + "\n板件宽度【垂直纹理的方向】: " + panel.Width + "\n厚度为" + panel.Thickness + "mm的板在板件长度【顺着纹理的方向】不在可加工的 " + checkinifile.IniReadValue("PanelLengthWidthfor50", "Length") + " 几个标准的尺寸内!\n请先修改模型!\n\n板件正反面加工码的生成已终止!", "警告",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    return;
                                }

                                for (int num = 0; num < WidthNum.Length; num++)
                                {
                                    if (Math.Abs(Convert.ToDouble(panel.Width) - Convert.ToDouble(WidthNum[num])) < 0.1)
                                    {
                                        WidthWorkable = true;
                                    }
                                }

                                if (!WidthWorkable)
                                {
                                    MessageBox.Show("三维家XML中的板件名称为: " + panel.Name + " ID号为: " + panel.ID + "\n板件长度【顺着纹理的方向】: " + panel.Length + "\n板件宽度【垂直纹理的方向】: " + panel.Width + "\n厚度为" + panel.Thickness + "mm的板在板件宽度【垂直纹理的方向】不在可加工的 " + checkinifile.IniReadValue("PanelLengthWidthfor50", "Width") + " 几个标准的尺寸内!\n请先修改模型!\n\n板件正反面加工码的生成已终止!", "警告",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    return;
                                }
                            }
                            else
                            {
                                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                return;
                            }
                        }
                        #endregion
                    }
                    #endregion

                    sampanel.Name = nest.PartName;
                    sampanel.CuttingLength = panel.ActualLength;
                    sampanel.CuttingWidth = panel.ActualWidth;
                    sampanel.CuttingThickness = panel.Thickness;
                    sampanel.CuttingNum = nest.Num;

                    int Num = 0;
                    sampanel.EL1 = nest.EbL1;
                    Num += sampanel.GetEdgeNum(sampanel.EL1);
                    sampanel.EL2 = nest.EbL2;
                    Num += sampanel.GetEdgeNum(sampanel.EL2);
                    sampanel.EW1 = nest.EbW1;
                    Num += sampanel.GetEdgeNum(sampanel.EW1);
                    sampanel.EW2 = nest.EbW2;
                    Num += sampanel.GetEdgeNum(sampanel.EW2);
                    if (Num != 0)
                        sampanel.EageNum = Num.ToString() + "薄";
                    else
                        sampanel.EageNum = string.Empty;

                    sampanel.Face6FileName = nest.F6FileName;
                    sampanel.Face5FileName = nest.F5FileName;
                    sampanel.YiXing = string.Empty;
                    sampanel.Length = nest.Length;
                    sampanel.Width = nest.Width;
                    sampanel.Thickness = panel.Thickness;
                    sampanel.Num = nest.Num;
                    sampanel.Area = sampanel.GetArea(sampanel.Length, sampanel.Width).ToString("F3");
                    sampanel.PackNo = string.Empty;
                    sampanel.CoderNo = string.Empty;
                    sampanel.CabinetNo = "1";
                    sampanel.HoleNum = "0";
                    sampanel.Material = nest.Material;
                    sampanel.DrawerNo = panel.drawer.ToString();
                    sampanel.PartNumber = panel.PartNumber; //20181102

                    Part.Add(sampanel);
                }
            }
            sw.Flush();
            sw.Close();

            if (ComboBox_3VJ_SMAX.SelectedIndex - 0 < 0.1) //如果选项卡上选择的是索引号为0,则为原先与普实对接的，为1则为与SMAX对接
            {
                if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))
                {
                    MessageBox.Show("总的排程单生成成功,共有 " + list.Count.ToString() + " 块工件! \n 接下来会检测此订单在服务器上面有多少订单号。请稍等！！");


                    #region 检测订单号对应服务器上面有哪些订单行 20180719
                    string ftpserver = @"ftp://139.196.188.94";
                    string ftpuser = "rckj";
                    string ftppassword = "123456";

                    string ftpserveraddcdwj = "ftp://139.196.188.94/cdwj/";

                    string needcheckfolder_temp = Path.Combine(ftpserveraddcdwj, Path.GetFileNameWithoutExtension(csvpath));
                    string needcheckfolder = Path.Combine(needcheckfolder_temp, "Project");

                    FTPclient client = new FTPclient(ftpserver, ftpuser, ftppassword, true);

                    ArrayList orderlinenumbers = new ArrayList();
                    for (int orderlinenum = 1; orderlinenum < 51; orderlinenum++)
                    {
                        string needcheckfile_temp = Path.Combine(needcheckfolder, Path.GetFileNameWithoutExtension(csvpath));
                        string needcheckfile = needcheckfile_temp + "_" + orderlinenum.ToString() + "_" + "DMSCSV.zip";
                        if (client.FtpFileExists(needcheckfile))
                        {
                            orderlinenumbers.Add(orderlinenum.ToString());
                        }
                    }

                    string strs = string.Empty;
                    foreach (string str in orderlinenumbers)
                    {
                        strs = strs + str + ",";
                    }

                    textBox2.Text = strs.TrimEnd(',');

                    if (strs != string.Empty)
                        MessageBox.Show("订单行检测成功！分别有 " + strs.TrimEnd(',') + "订单行 !");
                    else
                    {
                        MessageBox.Show("此订单服务器上无订单号_订单行信息，无须打包!");
                        return;
                    }
                    #endregion

                    #region 批量打压缩包 20180719
                    if (File.Exists(csvpath))
                        File.Delete(csvpath);

                    foreach (string str in orderlinenumbers)
                    {
                        string writedir = csvpath + "新CSV压缩包";
                        string newfilepath = Path.Combine(writedir, Path.GetFileNameWithoutExtension(csvpath)) + "_" + str + "_" + "DMSCSV.zip";

                        if (!Directory.Exists(writedir))
                            Directory.CreateDirectory(writedir);

                        if (File.Exists(newfilepath))
                            File.Delete(newfilepath);

                        ZipFile zip = new ZipFile(newfilepath, Encoding.Default);
                        zip.AddDirectory(csvpath);

                        zip.Save();
                    }


                    MessageBox.Show("新CSV压缩包生成成功!");
                    #endregion
                }
                else
                    MessageBox.Show("总的排程单生成成功,共有 " + list.Count.ToString() + " 块工件!");
            }
            else
            {              
                if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))
                {
                    Fisrt_Csv_Path = @"C: \Users\sxg035\Desktop\MV\三维家订单正反面加工码";
                }

                try
                {
                    if (!Path.GetFileNameWithoutExtension(csvpath).StartsWith("EMS-"))
                    {
                        DirectoryInfo direinfo = new DirectoryInfo(csvpath);

                        string Second_Csv_Path = Path.Combine(Fisrt_Csv_Path, Path.GetFileNameWithoutExtension(csvpath));
                        if (Directory.Exists(Second_Csv_Path))
                        {
                            DialogResult resault = MessageBox.Show("当前订单号: " + Path.GetFileNameWithoutExtension(csvpath) + " 已经生成过正反面加工码!" +
                                "\n\n点击 确定 按钮重新生成!并确定SAMEX系统中的 " + Path.GetFileNameWithoutExtension(csvpath) + ".xls 文件是否是当前的最新版本!\n若不是," +
                                "请上传最新版本的 " + Path.GetFileNameWithoutExtension(csvpath) + ".xls文件!" +
                                "\n\n点击 取消 按钮直接返回", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                            if (resault == DialogResult.OK)
                            {
                                Directory.Delete(Second_Csv_Path, true);
                                Directory.CreateDirectory(Second_Csv_Path);
                            }
                            else
                            {
                                listView1.Items.Clear();
                                listView2.Items.Clear();
                                return;

                            }

                        }
                        else
                        {
                            Directory.CreateDirectory(Second_Csv_Path);
                        }

                        foreach (FileSystemInfo fileinfo in direinfo.GetFiles())
                        {
                            if (fileinfo is DirectoryInfo)
                            {
                                MessageBox.Show(csvpath + " \n这文件夹下不应该再有文件夹啊!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                            else
                            {
                                FileInfo file = new FileInfo(fileinfo.FullName);
                                file.Attributes = FileAttributes.Normal;
                                file.CopyTo(Path.Combine(Second_Csv_Path, Path.GetFileName(fileinfo.FullName)));
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("订单号: " + Path.GetFileNameWithoutExtension(csvpath) + " 是三维家的订单号\n\n请将此订单号修改为SAMEX订单号重新执行!\n\n" +
                            "请注意程序终止!","警告",MessageBoxButtons.OK,MessageBoxIcon.Error);
                        return;
                    }
                

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }


                if (donotneedpart == 0)
                    MessageBox.Show("输出板件的正反面加工码csv文件成功!\n\n共有" + list.Count.ToString() + " 块板件!", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show("输出板件的正反面加工码csv文件成功!\n\n共有" + list.Count.ToString() + " 块板件!\n\n但有" + donotneedpart.ToString() + " 块板被过滤掉了!", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }


        }


        private void 导入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button1_Click(this, e);
        }

        private void 输出CSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button2_Click(this, e);
        }

        private void 修复CSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button12_Click(this,e);
        }

        private void 压缩及上传ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button3_Click(this, e);
        }

        private void 板件PToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button5_Click(this, e);
        }

        private void 背板BToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button6_Click(this, e);
        }

        private void 脚板JToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button7_Click(this,e);
        }

        private void 门板MToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button8_Click(this, e);
        }

        private void 水晶SToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button9_Click(this, e);
        }

        private void 抽屉CToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button10_Click(this, e);
        }

        private void 成品清单ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button11_Click(this, e);
        }

        private void pTP160数据ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button4_Click(this, e);
        }

        private void button_InputXML_Click(object sender, EventArgs e)
        {
            this.button1_Click(this, e);
        }

        private void button_OutCsv_Click(object sender, EventArgs e)
        {
            this.button2_Click(this, e);
            this.button_ModifyCsv_Click(this, e);
        }

        private void button_ModifyCsv_Click(object sender, EventArgs e)
        {
            this.button12_Click(this, e);
        }

        private void button_OutReport_Click(object sender, EventArgs e)
        {
            if (Part.Count == 0)
            {
                MessageBox.Show("没有导入三维家的XML文件或此XML文件没有需要加工的板件!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int drawerNum = 0;
            string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
            IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
            if (inifile.ExistINIFile())
            {
                drawerNum = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
            }
            else
            {
                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                return;
            }

            string Part_template = @"\\192.168.1.20\数据源\模板忽删\板件清单.xls";
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))
            {
                Part_template = @"D:\模板忽删\板件清单.xls";
            }
            IWorkbook bookproduct = null;
            bookproduct = Factory.GetWorkbook(Part_template);

            #region EXCEL表格的连接
            DataTable ExcelTable1;
            DataTable ExcelTable2;
            DataTable ExcelTable3;
            DataSet ds1 = new DataSet();
            DataSet ds2 = new DataSet();
            DataSet ds3 = new DataSet();

            //Excel的连接
            OleDbConnection objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Part_template + ";" + "Extended Properties=Excel 8.0;");  //需要在选择的文件里循环
            objConn.Open();
            DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

            string tableName1 = schemaTable.Rows[0][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 
            string tableName2 = schemaTable.Rows[2][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 
            string tableName3 = schemaTable.Rows[4][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 

            string strSql1 = "select * from [" + tableName1 + "]";
            OleDbCommand objCmd1 = new OleDbCommand(strSql1, objConn);
            OleDbDataAdapter myData1 = new OleDbDataAdapter(strSql1, objConn);
            myData1.Fill(ds1, tableName1);//填充数据

            string strSql2 = "select * from [" + tableName2 + "]";
            OleDbCommand objCmd2 = new OleDbCommand(strSql2, objConn);
            OleDbDataAdapter myData2 = new OleDbDataAdapter(strSql2, objConn);
            myData2.Fill(ds2, tableName2);//填充数据

            string strSql3 = "select * from [" + tableName3 + "]";
            OleDbCommand objCmd3 = new OleDbCommand(strSql3, objConn);
            OleDbDataAdapter myData3 = new OleDbDataAdapter(strSql3, objConn);
            myData3.Fill(ds3, tableName3);//填充数据

            objConn.Close();

            ExcelTable1 = ds1.Tables[tableName1];
            int iColums1 = ExcelTable1.Columns.Count;//列数
            int iRows1 = ExcelTable1.Rows.Count;//行数

            ExcelTable2 = ds2.Tables[tableName2];
            int iColums2 = ExcelTable2.Columns.Count;//列数
            int iRows2 = ExcelTable2.Rows.Count;//行数

            ExcelTable3 = ds3.Tables[tableName3];
            int iColums3 = ExcelTable3.Columns.Count;//列数
            int iRows3 = ExcelTable3.Rows.Count;//行数
            #endregion

            IRange range1 = bookproduct.Worksheets[0].Cells;//第一页
            IRange range2 = bookproduct.Worksheets[1].Cells;//第二页
            IRange range3 = bookproduct.Worksheets[2].Cells;//第三页
            var sheet_three = Part.Where(it_three => it_three.Name.Contains("CT-")).ToList();
            var after_sheet_three = Part.Where(it_three => !it_three.Name.Contains("CT-")).ToList();

            var sheet_two = after_sheet_three.Where(it_two => (Convert.ToDouble(it_two.Width) <= 60 || it_two.Thickness == "5" || it_two.Thickness == "15" || it_two.Thickness == "8" || it_two.Thickness == "50" || it_two.Material.Contains("SJ") || it_two.Material.Contains("PET"))).ToList();
            var after_sheet_two = after_sheet_three.Where(it_two => (Convert.ToDouble(it_two.Width) > 60 && it_two.Thickness != "5" && it_two.Thickness != "15" && it_two.Thickness != "8" && it_two.Thickness != "50" && !it_two.Material.Contains("SJ") && !it_two.Material.Contains("PET"))).ToList();

            //var sheet_one = Part.Where(it_one => it_one.Thickness == 18.ToString()).ToList();

            int i = 0;
            foreach(var str in after_sheet_two)
            {
                i++;
                range1[iRows1 + i, 0].Value = str.Index;
                range1[iRows1 + i, 1].Value = str.ColorSku;
                range1[iRows1 + i, 2].Value = str.Name;
                range1[iRows1 + i, 3].Value = str.CuttingLength;
                range1[iRows1 + i, 4].Value = str.CuttingWidth;
                range1[iRows1 + i, 5].Value = str.CuttingThickness;
                range1[iRows1 + i, 6].Value = str.CuttingNum;
                range1[iRows1 + i, 7].Value = str.EageNum;
                range1[iRows1 + i, 8].Value = str.Face6FileName;
                range1[iRows1 + i, 9].Value = str.Face5FileName;
                range1[iRows1 + i, 10].Value = str.YiXing;
                range1[iRows1 + i, 11].Value = str.Length;
                range1[iRows1 + i, 12].Value = str.Width;
                range1[iRows1 + i, 13].Value = str.Thickness;
                range1[iRows1 + i, 14].Value = str.Num;
                range1[iRows1 + i, 15].Value = str.Area;
                range1[iRows1 + i, 16].Value = str.PackNo;
                range1[iRows1 + i, 17].Value = str.CoderNo;
                range1[iRows1 + i, 18].Value = str.CabinetNo;
                range1[iRows1 + i, 19].Value = str.HoleNum;
                range1[iRows1 + i, 20].Value = str.Material;
                range1[iRows1 + i, 21].Value = str.EL1;
                range1[iRows1 + i, 22].Value = str.EL2;
                range1[iRows1 + i, 23].Value = str.EW1;
                range1[iRows1 + i, 24].Value = str.EW2;
                range1[iRows1 + i, 25].Value = str.PartNumber; //20181102
            }
            int j = 0;
            foreach (var str in sheet_two)
            {
                j++;
                range2[iRows2 + j, 0].Value = str.Index;
                range2[iRows2 + j, 1].Value = str.ColorSku;
                range2[iRows2 + j, 2].Value = str.Name;
                range2[iRows2 + j, 3].Value = str.CuttingLength;
                range2[iRows2 + j, 4].Value = str.CuttingWidth;
                range2[iRows2 + j, 5].Value = str.CuttingThickness;
                range2[iRows2 + j, 6].Value = str.CuttingNum;
                range2[iRows2 + j, 7].Value = str.EageNum;
                range2[iRows2 + j, 8].Value = str.Face6FileName;
                range2[iRows2 + j, 9].Value = str.Face5FileName;
                range2[iRows2 + j, 10].Value = str.YiXing;
                range2[iRows2 + j, 11].Value = str.Length;
                range2[iRows2 + j, 12].Value = str.Width;
                range2[iRows2 + j, 13].Value = str.Thickness;
                range2[iRows2 + j, 14].Value = str.Num;
                range2[iRows2 + j, 15].Value = str.Area;
                range2[iRows2 + j, 16].Value = str.PackNo;
                range2[iRows2 + j, 17].Value = str.CoderNo;
                range2[iRows2 + j, 18].Value = str.CabinetNo;
                range2[iRows2 + j, 19].Value = str.HoleNum;
                range2[iRows2 + j, 20].Value = str.Material;
                range2[iRows2 + j, 21].Value = str.EL1;
                range2[iRows2 + j, 22].Value = str.EL2;
                range2[iRows2 + j, 23].Value = str.EW1;
                range2[iRows2 + j, 24].Value = str.EW2;
                range2[iRows2 + j, 25].Value = str.PartNumber; //20181102
            }
            int k = 0;
            foreach (var str in sheet_three)
            {
                k++;
                range3[iRows3 + k, 0].Value = str.Index;
                range3[iRows3 + k, 1].Value = str.ColorSku;
                range3[iRows3 + k, 2].Value = str.Name;
                range3[iRows3 + k, 3].Value = str.CuttingLength;
                range3[iRows3 + k, 4].Value = str.CuttingWidth;
                range3[iRows3 + k, 5].Value = str.CuttingThickness;
                range3[iRows3 + k, 6].Value = str.CuttingNum;
                range3[iRows3 + k, 7].Value = str.EageNum;
                range3[iRows3 + k, 8].Value = str.Face6FileName;
                range3[iRows3 + k, 9].Value = str.Face5FileName;
                range3[iRows3 + k, 10].Value = str.YiXing;
                range3[iRows3 + k, 11].Value = str.Length;
                range3[iRows3 + k, 12].Value = str.Width;
                range3[iRows3 + k, 13].Value = str.Thickness;
                range3[iRows3 + k, 14].Value = str.Num;
                range3[iRows3 + k, 15].Value = str.Area;
                range3[iRows3 + k, 16].Value = str.PackNo;
                range3[iRows3 + k, 17].Value = str.CoderNo;
                range3[iRows3 + k, 18].Value = str.CabinetNo;
                range3[iRows3 + k, 19].Value = str.HoleNum;
                range3[iRows3 + k, 20].Value = str.DrawerNo;
                range3[iRows3 + k, 21].Value = str.Material;
                range3[iRows3 + k, 22].Value = str.EL1;
                range3[iRows3 + k, 23].Value = str.EL2;
                range3[iRows3 + k, 24].Value = str.EW1;
                range3[iRows3 + k, 25].Value = str.EW2;
                range3[iRows3 + k, 26].Value = str.PartNumber;//20181102
            }

            range1 = bookproduct.Worksheets[0].Range[0,0,iRows1 + i, 25];//20181102
            range1.Borders.Color = Color.Black;
            range2 = bookproduct.Worksheets[1].Range[0, 0, iRows2 + j, 25];//20181102
            range2.Borders.Color = Color.Black;
            range3 = bookproduct.Worksheets[2].Range[0, 0, iRows3 + k, 26];//20181102
            range3.Borders.Color = Color.Black;

            saveFileDialog1.Filter = "SAMEX板件清单|*.xls";
            saveFileDialog1.FileName = Path.GetFileNameWithoutExtension(csvpath);

            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))
            {
                if (!Directory.Exists(@"C:\Users\sxg035.000\Desktop"))
                    saveFileDialog1.InitialDirectory = @"C:\Users\sxg035\Desktop";
                else
                    saveFileDialog1.InitialDirectory = @"C:\Users\sxg035.000\Desktop";
            }

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                bookproduct.SaveAs(saveFileDialog1.FileName, FileFormat.Excel8);

                if (Part.Count == after_sheet_two.Count + sheet_two.Count + sheet_three.Count)
                {
                    MessageBox.Show("导出成功!" + "\n" + "\n共有 " + Part.Count + " 块板件!" , "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("总共的板件数目 " + Part.Count + "\n第一页的板件数目 " +
                        after_sheet_two.Count + "\n第二页的板件数目 " + sheet_two.Count +
                        "\n第三页的板件数目" + sheet_three.Count + "\n" + "\n" + "\n请注意板件分类与总数不相等!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }


        }

        List<SAMXPanel> Sheet1_Panel = new List<SAMXPanel>();
        List<SAMXPanel> Sheet2_Panel = new List<SAMXPanel>();
        List<SAMXPanel> Sheet3_Panel = new List<SAMXPanel>();
        List<MangerOrderCabinetNo> mangerordercabinetno = new List<MangerOrderCabinetNo>();
        string SavePlanName = string.Empty;
        List<NestlistEntity> Nest_Sheet1 = new List<NestlistEntity>();
        List<NestlistEntity> Nest_Sheet2 = new List<NestlistEntity>();
        private void button_Plan_Click(object sender, EventArgs e)
        {
            Nest_Sheet1.Clear();
            Nest_Sheet2.Clear();
            Sheet1_Panel.Clear();
            Sheet2_Panel.Clear();
            Sheet3_Panel.Clear();
            mangerordercabinetno.Clear();
            SavePlanName = string.Empty;
            int PanelNum = 0;
            SAMXPanel samxpanel = new SAMXPanel();
            MangerOrderCabinetNo manger = new MangerOrderCabinetNo();
            openFileDialog1.Filter = "SAMEX板件清单|*.xls";
            openFileDialog1.Multiselect = true;
            openFileDialog1.FileName = string.Empty;
            IWorkbook bookproduct = null;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                List<string> OrderNum = new List<string>();

                string[] OrderFileNum = openFileDialog1.FileNames;

                for (int num = 0; num < OrderFileNum.Length; num++)
                {
                    OrderNum.Add(Path.GetFileNameWithoutExtension(OrderFileNum[num]));
                }

                OrderNum.Sort(); //对订单号进行从小到大的排序 //20180918

                string[] FileNum = new string[OrderNum.Count];

                for (int num = 0;num<OrderNum.Count;num++)
                {
                    FileNum[num] = Path.Combine(Path.GetDirectoryName(OrderFileNum[num]), OrderNum[num] + ".xls");
                }


                for (int num = 0; num < FileNum.Length; num++)
                {
                    manger = new MangerOrderCabinetNo();
                    manger.CabinetNo = (num + 1).ToString();
                    manger.OrderNo = Path.GetFileNameWithoutExtension(FileNum[num]);
                    mangerordercabinetno.Add(manger);

                    #region EXCEL表格的连接
                    DataTable ExcelTable1;
                    DataTable ExcelTable2;
                    DataTable ExcelTable3;
                    DataSet ds1 = new DataSet();
                    DataSet ds2 = new DataSet();
                    DataSet ds3 = new DataSet();

                    //Excel的连接
                    OleDbConnection objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileNum[num] + ";" + "Extended Properties=Excel 8.0;");  //需要在选择的文件里循环
                    objConn.Open();
                    DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

                    string tableName1 = schemaTable.Rows[0][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 
                    string tableName2 = schemaTable.Rows[2][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 
                    string tableName3 = schemaTable.Rows[4][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 

                    string strSql1 = "select * from [" + tableName1 + "]";
                    OleDbCommand objCmd1 = new OleDbCommand(strSql1, objConn);
                    OleDbDataAdapter myData1 = new OleDbDataAdapter(strSql1, objConn);
                    myData1.Fill(ds1, tableName1);//填充数据

                    string strSql2 = "select * from [" + tableName2 + "]";
                    OleDbCommand objCmd2 = new OleDbCommand(strSql2, objConn);
                    OleDbDataAdapter myData2 = new OleDbDataAdapter(strSql2, objConn);
                    myData2.Fill(ds2, tableName2);//填充数据

                    string strSql3 = "select * from [" + tableName3 + "]";
                    OleDbCommand objCmd3 = new OleDbCommand(strSql3, objConn);
                    OleDbDataAdapter myData3 = new OleDbDataAdapter(strSql3, objConn);
                    myData3.Fill(ds3, tableName3);//填充数据

                    objConn.Close();

                    ExcelTable1 = ds1.Tables[tableName1];
                    int iColums1 = ExcelTable1.Columns.Count;//列数
                    int iRows1 = ExcelTable1.Rows.Count;//行数

                    ExcelTable2 = ds2.Tables[tableName2];
                    int iColums2 = ExcelTable2.Columns.Count;//列数
                    int iRows2 = ExcelTable2.Rows.Count;//行数

                    ExcelTable3 = ds3.Tables[tableName3];
                    int iColums3 = ExcelTable3.Columns.Count;//列数
                    int iRows3 = ExcelTable3.Rows.Count;//行数
                    #endregion

                    bookproduct = Factory.GetWorkbook(FileNum[num]);

                    IRange range1 = bookproduct.Worksheets[0].Cells;//第一页
                    IRange range2 = bookproduct.Worksheets[1].Cells;//第二页
                    IRange range3 = bookproduct.Worksheets[2].Cells;//第三页
                    int i = 9;
                    for (int sheet1_num = 0;sheet1_num <= iRows1 - i; sheet1_num++)
                    {
                        samxpanel = new SAMXPanel();
                        samxpanel.Index = (++PanelNum).ToString();
                        samxpanel.ColorSku = range1[i + sheet1_num, 1].Value == null ? string.Empty : range1[i + sheet1_num, 1].Value.ToString();
                        samxpanel.Name = range1[i + sheet1_num, 2].Value == null ? string.Empty : range1[i + sheet1_num, 2].Value.ToString();
                        samxpanel.CuttingLength = range1[i + sheet1_num, 3].Value == null ? string.Empty : range1[i + sheet1_num, 3].Value.ToString(); 
                        samxpanel.CuttingWidth = range1[i + sheet1_num, 4].Value == null ? string.Empty : range1[i + sheet1_num, 4].Value.ToString();
                        samxpanel.CuttingThickness = range1[i + sheet1_num, 5].Value == null ? string.Empty : range1[i + sheet1_num, 5].Value.ToString();
                        samxpanel.CuttingNum = range1[i + sheet1_num, 6].Value == null ? string.Empty : range1[i + sheet1_num, 6].Value.ToString();
                        samxpanel.EageNum = range1[i + sheet1_num, 7].Value == null ? string.Empty : range1[i + sheet1_num, 7].Value.ToString(); 
                        samxpanel.Face6FileName = range1[i + sheet1_num, 8].Value == null ? string.Empty : range1[i + sheet1_num, 8].Value.ToString();
                        samxpanel.Face5FileName = range1[i + sheet1_num, 9].Value == null ? string.Empty : range1[i + sheet1_num, 9].Value.ToString();
                        samxpanel.YiXing = range1[i + sheet1_num, 10].Value == null ? string.Empty : range1[i + sheet1_num, 10].Value.ToString();
                        samxpanel.Length = range1[i + sheet1_num, 11].Value == null ? string.Empty : range1[i + sheet1_num, 11].Value.ToString();
                        samxpanel.Width = range1[i + sheet1_num, 12].Value == null ? string.Empty : range1[i + sheet1_num, 12].Value.ToString();
                        samxpanel.Thickness = range1[i + sheet1_num, 13].Value == null ? string.Empty : range1[i + sheet1_num, 13].Value.ToString();
                        samxpanel.Num = range1[i + sheet1_num, 14].Value == null ? string.Empty : range1[i + sheet1_num, 14].Value.ToString();
                        samxpanel.Area = range1[i + sheet1_num, 15].Value == null ? string.Empty : range1[i + sheet1_num, 15].Value.ToString();
                        samxpanel.PackNo = range1[i + sheet1_num, 16].Value == null ? string.Empty : range1[i + sheet1_num, 16].Value.ToString();
                        samxpanel.CoderNo = range1[i + sheet1_num, 17].Value == null ? string.Empty : range1[i + sheet1_num, 17].Value.ToString();
                        samxpanel.CabinetNo = (num + 1).ToString();
                        samxpanel.HoleNum = range1[i + sheet1_num, 19].Value == null ? string.Empty : range1[i + sheet1_num, 19].Value.ToString();
                        samxpanel.Material = range1[i + sheet1_num, 20].Value == null ? string.Empty : range1[i + sheet1_num, 20].Value.ToString();
                        samxpanel.EL1 = range1[i + sheet1_num, 21].Value == null ? string.Empty : range1[i + sheet1_num, 21].Value.ToString();
                        samxpanel.EL2 = range1[i + sheet1_num, 22].Value == null ? string.Empty : range1[i + sheet1_num, 22].Value.ToString();
                        samxpanel.EW1 = range1[i + sheet1_num, 23].Value == null ? string.Empty : range1[i + sheet1_num, 23].Value.ToString();
                        samxpanel.EW2 = range1[i + sheet1_num, 24].Value == null ? string.Empty : range1[i + sheet1_num, 24].Value.ToString();
                        Sheet1_Panel.Add(samxpanel);
                    }
                    int j = 9;
                    for (int sheet2_num = 0; sheet2_num <= iRows2 - j; sheet2_num++)
                    {
                        samxpanel = new SAMXPanel();
                        samxpanel.Index = (++PanelNum).ToString();
                        samxpanel.ColorSku = range2[j + sheet2_num, 1].Value == null ? string.Empty : range2[j + sheet2_num, 1].Value.ToString();
                        samxpanel.Name = range2[j + sheet2_num, 2].Value == null ? string.Empty : range2[j + sheet2_num, 2].Value.ToString();
                        samxpanel.CuttingLength = range2[j + sheet2_num, 3].Value == null ? string.Empty : range2[j + sheet2_num, 3].Value.ToString();
                        samxpanel.CuttingWidth = range2[j + sheet2_num, 4].Value == null ? string.Empty : range2[j + sheet2_num, 4].Value.ToString();
                        samxpanel.CuttingThickness = range2[j + sheet2_num, 5].Value == null ? string.Empty : range2[j + sheet2_num, 5].Value.ToString();
                        samxpanel.CuttingNum = range2[j + sheet2_num, 6].Value == null ? string.Empty : range2[j + sheet2_num, 6].Value.ToString();
                        samxpanel.EageNum = range2[j + sheet2_num, 7].Value == null ? string.Empty : range2[j + sheet2_num, 7].Value.ToString();
                        samxpanel.Face6FileName = range2[j + sheet2_num, 8].Value == null ? string.Empty : range2[j + sheet2_num, 8].Value.ToString();
                        samxpanel.Face5FileName = range2[j + sheet2_num, 9].Value == null ? string.Empty : range2[j + sheet2_num, 9].Value.ToString();
                        samxpanel.YiXing = range2[j + sheet2_num, 10].Value == null ? string.Empty : range2[j + sheet2_num, 10].Value.ToString();
                        samxpanel.Length = range2[j + sheet2_num, 11].Value == null ? string.Empty : range2[j + sheet2_num, 11].Value.ToString();
                        samxpanel.Width = range2[j + sheet2_num, 12].Value == null ? string.Empty : range2[j + sheet2_num, 12].Value.ToString();
                        samxpanel.Thickness = range2[j + sheet2_num, 13].Value == null ? string.Empty : range2[j + sheet2_num, 13].Value.ToString();
                        samxpanel.Num = range2[j + sheet2_num, 14].Value == null ? string.Empty : range2[j + sheet2_num, 14].Value.ToString();
                        samxpanel.Area = range2[j + sheet2_num, 15].Value == null ? string.Empty : range2[j + sheet2_num, 15].Value.ToString();
                        samxpanel.PackNo = range2[j + sheet2_num, 16].Value == null ? string.Empty : range2[j + sheet2_num, 16].Value.ToString();
                        samxpanel.CoderNo = range2[j + sheet2_num, 17].Value == null ? string.Empty : range2[j + sheet2_num, 17].Value.ToString();
                        samxpanel.CabinetNo = (num + 1).ToString();
                        samxpanel.HoleNum = range2[j + sheet2_num, 19].Value == null ? string.Empty : range2[j + sheet2_num, 19].Value.ToString();
                        samxpanel.Material = range2[j + sheet2_num, 20].Value == null ? string.Empty : range2[j + sheet2_num, 20].Value.ToString();
                        samxpanel.EL1 = range2[j + sheet2_num, 21].Value == null ? string.Empty : range2[j + sheet2_num, 21].Value.ToString();
                        samxpanel.EL2 = range2[j + sheet2_num, 22].Value == null ? string.Empty : range2[j + sheet2_num, 22].Value.ToString();
                        samxpanel.EW1 = range2[j + sheet2_num, 23].Value == null ? string.Empty : range2[j + sheet2_num, 23].Value.ToString();
                        samxpanel.EW2 = range2[j + sheet2_num, 24].Value == null ? string.Empty : range2[j + sheet2_num, 24].Value.ToString();
                        Sheet2_Panel.Add(samxpanel);
                    }
                    int k = 9;
                    for (int sheet3_num = 0; sheet3_num <= iRows3 - k; sheet3_num++)
                    {
                        samxpanel = new SAMXPanel();
                        samxpanel.Index = (++PanelNum).ToString();
                        samxpanel.ColorSku = range3[k + sheet3_num, 1].Value == null ? string.Empty : range3[k + sheet3_num, 1].Value.ToString();
                        samxpanel.Name = range3[k + sheet3_num, 2].Value == null ? string.Empty : range3[k + sheet3_num, 2].Value.ToString();
                        samxpanel.CuttingLength = range3[k + sheet3_num, 3].Value == null ? string.Empty : range3[k + sheet3_num, 3].Value.ToString();
                        samxpanel.CuttingWidth = range3[k + sheet3_num, 4].Value == null ? string.Empty : range3[k + sheet3_num, 4].Value.ToString();
                        samxpanel.CuttingThickness = range3[k + sheet3_num, 5].Value == null ? string.Empty : range3[k + sheet3_num, 5].Value.ToString();
                        samxpanel.CuttingNum = range3[k + sheet3_num, 6].Value == null ? string.Empty : range3[k + sheet3_num, 6].Value.ToString();
                        samxpanel.EageNum = range3[k + sheet3_num, 7].Value == null ? string.Empty : range3[k + sheet3_num, 7].Value.ToString();
                        samxpanel.Face6FileName = range3[k + sheet3_num, 8].Value == null ? string.Empty : range3[k + sheet3_num, 8].Value.ToString();
                        samxpanel.Face5FileName = range3[k + sheet3_num, 9].Value == null ? string.Empty : range3[k + sheet3_num, 9].Value.ToString();
                        samxpanel.YiXing = range3[k + sheet3_num, 10].Value == null ? string.Empty : range3[k + sheet3_num, 10].Value.ToString();
                        samxpanel.Length = range3[k + sheet3_num, 11].Value == null ? string.Empty : range3[k + sheet3_num, 11].Value.ToString();
                        samxpanel.Width = range3[k + sheet3_num, 12].Value == null ? string.Empty : range3[k + sheet3_num, 12].Value.ToString();
                        samxpanel.Thickness = range3[k + sheet3_num, 13].Value == null ? string.Empty : range3[k + sheet3_num, 13].Value.ToString();
                        samxpanel.Num = range3[k + sheet3_num, 14].Value == null ? string.Empty : range3[k + sheet3_num, 14].Value.ToString();
                        samxpanel.Area = range3[k + sheet3_num, 15].Value == null ? string.Empty : range3[k + sheet3_num, 15].Value.ToString();
                        samxpanel.PackNo = range3[k + sheet3_num, 16].Value == null ? string.Empty : range3[k + sheet3_num, 16].Value.ToString();
                        samxpanel.CoderNo = range3[k + sheet3_num, 17].Value == null ? string.Empty : range3[k + sheet3_num, 17].Value.ToString();
                        samxpanel.CabinetNo = (num + 1).ToString();
                        samxpanel.HoleNum = range3[k + sheet3_num, 19].Value == null ? string.Empty : range3[k + sheet3_num, 19].Value.ToString();
                        samxpanel.DrawerNo = range3[k + sheet3_num,20].Value == null ? string.Empty : range3[k + sheet3_num, 20].Value.ToString();
                        samxpanel.Material = range3[k + sheet3_num, 21].Value == null ? string.Empty : range3[k + sheet3_num, 21].Value.ToString();
                        samxpanel.EL1 = range3[k + sheet3_num, 22].Value == null ? string.Empty : range3[k + sheet3_num, 22].Value.ToString();
                        samxpanel.EL2 = range3[k + sheet3_num, 23].Value == null ? string.Empty : range3[k + sheet3_num, 23].Value.ToString();
                        samxpanel.EW1 = range3[k + sheet3_num, 24].Value == null ? string.Empty : range3[k + sheet3_num, 24].Value.ToString();
                        samxpanel.EW2 = range3[k + sheet3_num, 25].Value == null ? string.Empty : range3[k + sheet3_num, 25].Value.ToString();
                        Sheet3_Panel.Add(samxpanel);
                    }
                }
            }
            else
            {
                return;
            }
            this.button_PlanOutReport_Click(this, e);  //因为有好多重复，故将其他的部份放另外一个按钮里
        }

        private void button_PlanOutReport_Click(object sender, EventArgs e)
        {
            NestlistEntity nest = new NestlistEntity();
            string Part_template = @"\\192.168.1.20\数据源\模板忽删\板件清单.xls";
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))
            {
                Part_template = @"D:\模板忽删\板件清单.xls";
            }
            IWorkbook bookproduct = null;
            bookproduct = Factory.GetWorkbook(Part_template);

            #region EXCEL表格的连接
            DataTable ExcelTable1;
            DataTable ExcelTable2;
            DataTable ExcelTable3;
            DataSet ds1 = new DataSet();
            DataSet ds2 = new DataSet();
            DataSet ds3 = new DataSet();

            //Excel的连接
            OleDbConnection objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Part_template + ";" + "Extended Properties=Excel 8.0;");  //需要在选择的文件里循环
            objConn.Open();
            DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

            string tableName1 = schemaTable.Rows[0][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 
            string tableName2 = schemaTable.Rows[2][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 
            string tableName3 = schemaTable.Rows[4][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 

            string strSql1 = "select * from [" + tableName1 + "]";
            OleDbCommand objCmd1 = new OleDbCommand(strSql1, objConn);
            OleDbDataAdapter myData1 = new OleDbDataAdapter(strSql1, objConn);
            myData1.Fill(ds1, tableName1);//填充数据

            string strSql2 = "select * from [" + tableName2 + "]";
            OleDbCommand objCmd2 = new OleDbCommand(strSql2, objConn);
            OleDbDataAdapter myData2 = new OleDbDataAdapter(strSql2, objConn);
            myData2.Fill(ds2, tableName2);//填充数据

            string strSql3 = "select * from [" + tableName3 + "]";
            OleDbCommand objCmd3 = new OleDbCommand(strSql3, objConn);
            OleDbDataAdapter myData3 = new OleDbDataAdapter(strSql3, objConn);
            myData3.Fill(ds3, tableName3);//填充数据

            objConn.Close();

            ExcelTable1 = ds1.Tables[tableName1];
            int iColums1 = ExcelTable1.Columns.Count;//列数
            int iRows1 = ExcelTable1.Rows.Count;//行数

            ExcelTable2 = ds2.Tables[tableName2];
            int iColums2 = ExcelTable2.Columns.Count;//列数
            int iRows2 = ExcelTable2.Rows.Count;//行数

            ExcelTable3 = ds3.Tables[tableName3];
            int iColums3 = ExcelTable3.Columns.Count;//列数
            int iRows3 = ExcelTable3.Rows.Count;//行数
            #endregion

            IRange range1 = bookproduct.Worksheets[0].Cells;//第一页
            IRange range2 = bookproduct.Worksheets[1].Cells;//第二页
            IRange range3 = bookproduct.Worksheets[2].Cells;//第三页
            int EndNum = 0; //用于计划最后板号的序号输出
            int i = 0;
            var After_Sheet1_Panel = Sheet1_Panel.OrderBy(it => it.ColorSku).ThenByDescending(it => Convert.ToDouble(it.Thickness)).ThenByDescending(it => Convert.ToDouble(it.Width)).ThenByDescending(it => Convert.ToDouble(it.Length)).ToList();
            foreach (var str in After_Sheet1_Panel)
            {
                i++;
                range1[iRows1 + i, 0].Value = ++EndNum;
                range1[iRows1 + i, 1].Value = str.ColorSku;
                range1[iRows1 + i, 2].Value = str.Name;
                range1[iRows1 + i, 3].Value = str.CuttingLength;
                range1[iRows1 + i, 4].Value = str.CuttingWidth;
                range1[iRows1 + i, 5].Value = str.CuttingThickness;
                range1[iRows1 + i, 6].Value = str.CuttingNum;
                range1[iRows1 + i, 7].Value = str.EageNum;
                range1[iRows1 + i, 8].Value = str.Face6FileName;
                range1[iRows1 + i, 9].Value = str.Face5FileName;
                range1[iRows1 + i, 10].Value = str.YiXing;
                range1[iRows1 + i, 11].Value = str.Length;
                range1[iRows1 + i, 12].Value = str.Width;
                range1[iRows1 + i, 13].Value = str.Thickness;
                range1[iRows1 + i, 14].Value = str.Num;
                range1[iRows1 + i, 15].Value = str.Area;
                range1[iRows1 + i, 16].Value = str.PackNo;
                range1[iRows1 + i, 17].Value = str.CoderNo;
                range1[iRows1 + i, 18].Value = str.CabinetNo;
                range1[iRows1 + i, 19].Value = str.HoleNum;
                range1[iRows1 + i, 20].Value = str.Material;
                range1[iRows1 + i, 21].Value = str.EL1;
                range1[iRows1 + i, 22].Value = str.EL2;
                range1[iRows1 + i, 23].Value = str.EW1;
                range1[iRows1 + i, 24].Value = str.EW2;

                if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))//老系统自己做的排程单数据的生成
                {
                    nest = new NestlistEntity();
                    nest.Index = range1[iRows1 + i, 0].Value.ToString();
                    nest.Mat = str.Material;
                    nest.EW1 = str.EW1;
                    nest.EW2 = str.EW2;
                    nest.EL1 = str.EL1;
                    nest.EL2 = str.EL2;
                    nest.PartName = str.Name;
                    nest.Length = str.Length;
                    nest.Width = str.Width;
                    nest.Num = str.Num;
                    nest.Filename = str.Face5FileName;
                    nest.Filename6 = str.Face6FileName;
                    nest.Batch = string.Empty;
                    nest.Tasknum = str.CabinetNo;
                    nest.ID = string.Empty;
                    nest.Split = "22";
                    nest.Common5 = string.Empty;
                    nest.Common6 = string.Empty;
                    nest.Common7 = string.Empty;
                    nest.Common8 = string.Empty;
                    Nest_Sheet1.Add(nest);
                }
            }
            int j = 0;
            var After_Sheet2_Panel = Sheet2_Panel.OrderBy(it => it.ColorSku).ThenByDescending(it => Convert.ToDouble(it.Thickness)).ThenByDescending(it => Convert.ToDouble(it.Width)).ThenByDescending(it => Convert.ToDouble(it.Length)).ToList();
            foreach (var str in After_Sheet2_Panel)
            {
                j++;
                range2[iRows2 + j, 0].Value = ++EndNum;
                range2[iRows2 + j, 1].Value = str.ColorSku;
                range2[iRows2 + j, 2].Value = str.Name;
                range2[iRows2 + j, 3].Value = str.CuttingLength;
                range2[iRows2 + j, 4].Value = str.CuttingWidth;
                range2[iRows2 + j, 5].Value = str.CuttingThickness;
                range2[iRows2 + j, 6].Value = str.CuttingNum;
                range2[iRows2 + j, 7].Value = str.EageNum;
                range2[iRows2 + j, 8].Value = str.Face6FileName;
                range2[iRows2 + j, 9].Value = str.Face5FileName;
                range2[iRows2 + j, 10].Value = str.YiXing;
                range2[iRows2 + j, 11].Value = str.Length;
                range2[iRows2 + j, 12].Value = str.Width;
                range2[iRows2 + j, 13].Value = str.Thickness;
                range2[iRows2 + j, 14].Value = str.Num;
                range2[iRows2 + j, 15].Value = str.Area;
                range2[iRows2 + j, 16].Value = str.PackNo;
                range2[iRows2 + j, 17].Value = str.CoderNo;
                range2[iRows2 + j, 18].Value = str.CabinetNo;
                range2[iRows2 + j, 19].Value = str.HoleNum;
                range2[iRows2 + j, 20].Value = str.Material;
                range2[iRows2 + j, 21].Value = str.EL1;
                range2[iRows2 + j, 22].Value = str.EL2;
                range2[iRows2 + j, 23].Value = str.EW1;
                range2[iRows2 + j, 24].Value = str.EW2;

                if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))//老系统自己做的排程单数据的生成
                {
                    nest = new NestlistEntity();
                    nest.Index = range2[iRows2 + j, 0].Value.ToString();
                    nest.Mat = str.Material;
                    nest.EW1 = str.EW1;
                    nest.EW2 = str.EW2;
                    nest.EL1 = str.EL1;
                    nest.EL2 = str.EL2;
                    nest.PartName = str.Name;
                    nest.Length = str.Length;
                    nest.Width = str.Width;
                    nest.Num = str.Num;
                    nest.Filename = str.Face5FileName;
                    nest.Filename6 = str.Face6FileName;
                    nest.Batch = string.Empty;
                    nest.Tasknum = str.CabinetNo;
                    nest.ID = string.Empty;
                    nest.Split = "22";
                    nest.Common5 = string.Empty;
                    nest.Common6 = string.Empty;
                    nest.Common7 = string.Empty;
                    nest.Common8 = string.Empty;
                    Nest_Sheet2.Add(nest);
                }
            }
            int k = 0;
            string taskno = string.Empty;
            string plandrawerno = string.Empty;
            int drawerNumNo = 0;
            foreach (var str in Sheet3_Panel)
            {
                k++;
                range3[iRows3 + k, 0].Value = ++EndNum;
                range3[iRows3 + k, 1].Value = str.ColorSku;
                range3[iRows3 + k, 2].Value = str.Name;
                range3[iRows3 + k, 3].Value = str.CuttingLength;
                range3[iRows3 + k, 4].Value = str.CuttingWidth;
                range3[iRows3 + k, 5].Value = str.CuttingThickness;
                range3[iRows3 + k, 6].Value = str.CuttingNum;
                range3[iRows3 + k, 7].Value = str.EageNum;
                range3[iRows3 + k, 8].Value = str.Face6FileName;
                range3[iRows3 + k, 9].Value = str.Face5FileName;
                range3[iRows3 + k, 10].Value = str.YiXing;
                range3[iRows3 + k, 11].Value = str.Length;
                range3[iRows3 + k, 12].Value = str.Width;
                range3[iRows3 + k, 13].Value = str.Thickness;
                range3[iRows3 + k, 14].Value = str.Num;
                range3[iRows3 + k, 15].Value = str.Area;
                range3[iRows3 + k, 16].Value = str.PackNo;
                range3[iRows3 + k, 17].Value = str.CoderNo;
                range3[iRows3 + k, 18].Value = str.CabinetNo;
                range3[iRows3 + k, 19].Value = str.HoleNum;

                //将合并的抽屉按从1开始排
                if (plandrawerno.Equals(str.DrawerNo) && taskno.Equals(str.CabinetNo))
                    range3[iRows3 + k, 20].Value = drawerNumNo.ToString();
                else
                {
                    range3[iRows3 + k, 20].Value = (++drawerNumNo).ToString();
                    plandrawerno = str.DrawerNo;
                    taskno = str.CabinetNo;
                }

                range3[iRows3 + k, 21].Value = str.Material;
                range3[iRows3 + k, 22].Value = str.EL1;
                range3[iRows3 + k, 23].Value = str.EL2;
                range3[iRows3 + k, 24].Value = str.EW1;
                range3[iRows3 + k, 25].Value = str.EW2;
            }

            range1 = bookproduct.Worksheets[0].Range[0, 0, iRows1 + i, 24];
            range1.Borders.Color = Color.Black;
            bookproduct.Worksheets[0].PageSetup.PrintArea = "$A$1:$T$"+(iRows1 + i + 1).ToString();//将打印区域设置为整个工作区域
            range2 = bookproduct.Worksheets[1].Range[0, 0, iRows2 + j, 24];
            range2.Borders.Color = Color.Black;
            bookproduct.Worksheets[1].PageSetup.PrintArea = "$A$1:$T$" + (iRows2 + j + 1).ToString();
            range3 = bookproduct.Worksheets[2].Range[0, 0, iRows3 + k, 25];
            range3.Borders.Color = Color.Black;
            bookproduct.Worksheets[2].PageSetup.PrintArea = "$A$1:$U$" + (iRows3 + k + 1).ToString();

            saveFileDialog1.Filter = "SAMEX板件清单|*";
            DateTime dt = DateTime.Now;
            string nowdate = dt.Year.ToString().Substring(2, 2) + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0');
            saveFileDialog1.FileName = nowdate;
            saveFileDialog1.AddExtension = true;
   
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                int totalpanel = Sheet1_Panel.Count + Sheet2_Panel.Count + Sheet3_Panel.Count;
                SavePlanName = saveFileDialog1.FileName;
                range1[2, 1].Value = Path.GetFileName(SavePlanName);
                range2[2, 1].Value = Path.GetFileName(SavePlanName);
                range3[2, 1].Value = Path.GetFileName(SavePlanName);
                bookproduct.SaveAs(saveFileDialog1.FileName + "批次板件生产清单.xls", FileFormat.Excel8);
                this.button_PlanTask_Click(this, e);

                if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))//老系统自己做的排程单数据的生成
                {
                    this.button_OutSamexNest_Click(this, e);
                }
                    MessageBox.Show("总共的板件数目 " + totalpanel + "\n第一页的板件数目 " +
                        Sheet1_Panel.Count + "\n第二页的板件数目 " + Sheet2_Panel.Count +
                        "\n第三页的板件数目" + Sheet3_Panel.Count, "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button_PlanTask_Click(object sender, EventArgs e)
        {
            string Task_template = @"\\192.168.1.20\数据源\模板忽删\板件批次任务单.xls";
            string Third_Csv_Path = @"\\192.168.1.20\ptp160";
            if (Environment.MachineName.Equals("WIN-M2KRCJPECH2") || Environment.MachineName.Equals("SXG035") || Environment.MachineName.Equals("SXG035.000"))
            {
                Third_Csv_Path = @"C:\Users\sxg035\Desktop\MV\PTP160";
                Task_template = @"D:\模板忽删\板件批次任务单.xls";
                Fisrt_Csv_Path = @"C: \Users\sxg035\Desktop\MV\三维家订单正反面加工码";
            }

            IWorkbook bookproduct = null;
            bookproduct = Factory.GetWorkbook(Task_template);

            #region EXCEL表格的连接
            DataTable ExcelTable1;
            DataTable ExcelTable2;

            DataSet ds1 = new DataSet();
            DataSet ds2 = new DataSet();

            //Excel的连接
            OleDbConnection objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Task_template + ";" + "Extended Properties=Excel 8.0;");  //需要在选择的文件里循环
            objConn.Open();
            DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

            string tableName1 = schemaTable.Rows[0][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 
            string tableName2 = schemaTable.Rows[2][2].ToString().Trim();//获取 Excel 的表名，[0][2]是sheet1 [2][2]是sheet2 [6][2]是sheet3 

            string strSql1 = "select * from [" + tableName1 + "]";
            OleDbCommand objCmd1 = new OleDbCommand(strSql1, objConn);
            OleDbDataAdapter myData1 = new OleDbDataAdapter(strSql1, objConn);
            myData1.Fill(ds1, tableName1);//填充数据

            string strSql2 = "select * from [" + tableName2 + "]";
            OleDbCommand objCmd2 = new OleDbCommand(strSql2, objConn);
            OleDbDataAdapter myData2 = new OleDbDataAdapter(strSql2, objConn);
            myData2.Fill(ds2, tableName2);//填充数据

            objConn.Close();

            ExcelTable1 = ds1.Tables[tableName1];
            int iColums1 = ExcelTable1.Columns.Count;//列数
            int iRows1 = ExcelTable1.Rows.Count;//行数

            ExcelTable2 = ds2.Tables[tableName2];
            int iColums2 = ExcelTable2.Columns.Count;//列数
            int iRows2 = ExcelTable2.Rows.Count;//行数
            #endregion

            IRange range1 = bookproduct.Worksheets[0].Cells;//第一页
            IRange range2 = bookproduct.Worksheets[1].Cells;//第二页

            int j = 0;
            foreach (var str in mangerordercabinetno)
            {
                j++;
                range2[iRows2 + j, 0].Value = str.CabinetNo;
                range2[iRows2 + j, 1].Value = str.OrderNo;
                range2[iRows2 + j, 7].Value = "1";
                range2[iRows2 + j, 9].Value = str.OrderNo;
            }

            range2 = bookproduct.Worksheets[1].Range[0, 0, iRows2 + j, 9];
            range2.Borders.Color = Color.Black;
            range2.RowHeight = 21.25;
            bookproduct.Worksheets[1].PageSetup.PrintArea = "";
            bookproduct.SaveAs(SavePlanName + ".xls", FileFormat.Excel8);

            #region 将组好批次的正反面加工码拷贝至PTP160目录下
            string PTP160BachNum = Path.Combine(Third_Csv_Path, Path.GetFileNameWithoutExtension(SavePlanName));
            if (Directory.Exists(PTP160BachNum))
            {
                Directory.Delete(PTP160BachNum, true);
                Directory.CreateDirectory(PTP160BachNum);
            }
            else
            {
                Directory.CreateDirectory(PTP160BachNum);
            }

            foreach (var str in mangerordercabinetno)
            {
                string HaveOutCsvPath = Path.Combine(Path.Combine(Fisrt_Csv_Path, str.OrderNo));

                DirectoryInfo direinfo_third = new DirectoryInfo(HaveOutCsvPath);

                try
                {
                    foreach (FileSystemInfo fileinfo in direinfo_third.GetFiles())
                    {
                        if (fileinfo is DirectoryInfo)
                        {
                            MessageBox.Show(HaveOutCsvPath + " \n这文件夹下不应该再有文件夹啊!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        else
                        {
                            FileInfo file = new FileInfo(fileinfo.FullName);
                            file.Attributes = FileAttributes.Normal;
                            file.CopyTo(Path.Combine(PTP160BachNum, file.Name));
                        }
                    }
                }
                catch
                {
                    MessageBox.Show(HaveOutCsvPath + "\n\n服务器上面没有此目录,请与拆单人员联系.请他们重新处理此表单!", "出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            #endregion
        }

        private void ComboBox_3VJ_SMAX_SelectedIndexChanged(object sender, EventArgs e)
        {
            string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
            IniFiles inifile = new IniFiles(inipath);

            if (inifile.ExistINIFile())
            {
                if (ComboBox_3VJ_SMAX.SelectedIndex - 0 < 0.1)
                    inifile.IniWriteValue("SAMEX", "ps", "0");
                else
                    inifile.IniWriteValue("SAMEX", "ps", "1");
            }
            else
            {
                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                return;
            }
        }
        /// <summary>
        /// 用于产生老系统的排程单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_OutSamexNest_Click(object sender, EventArgs e)
        {
            string nestpathwriter = SavePlanName + ".csv";
            StreamWriter sw = new StreamWriter(nestpathwriter, false,Encoding.Default);
            sw.WriteLine("Nesting优化板件清单,,,,,,,,,,,,,,,,,,,");
            sw.WriteLine("板件信息,,,,,,,成品尺寸,,,加工信息,,,,,,,,,,,,,");
            sw.WriteLine("序号,材料,封边宽1,封边宽2,封边长1,封边长2,名称,高,宽,数量,正面加工码,反面加工码,备注1（批次号）,备注2（任务编码）,备注3（序列号）,备注4（分流）,备注5,备注6,备注7,备注8");
            foreach(var str in Nest_Sheet1)
            {
                str.Batch = Path.GetFileNameWithoutExtension(SavePlanName);
                str.ID = str.Batch + str.Index.PadLeft(3,'0') + str.Num.PadLeft(2,'0');
                sw.WriteLine(str.OutPutCsvString());
            }
            sw.Flush();
            sw.Close();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}