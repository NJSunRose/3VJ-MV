using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace _3VJ_MV 
{
    public class T3VProduct
    {
        public List<Cabinet> Cabinetlist = new List<Cabinet>();

        public void LoadFromXML(string filepath)
        {
            Cabinetlist = new List<Cabinet>();

            XmlDocument doc = new XmlDocument();

            doc.Load(filepath);

            XmlElement root = doc.DocumentElement;

            foreach (XmlNode xnl in root.ChildNodes)
            {
                //if (xnl.Name == "GroupCabinet")//临时增加20181108
                //{
                //    MessageBox.Show("出错了!出现了之前没有解析过的XML节点!\n\n转换终止!", "出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    return;
                //}

                if (xnl.Name == "GroupCabinet" || xnl.Name == "Cabinet") //20180912 删除之前很多的判断
                {
                    #region 20180912 组合柜里的板件或组件提取 还有就是单独的柜子的提取
                    Cabinet cabinet = new Cabinet();

                    cabinet.CabinetNo = xnl.Attributes["CabinetNo"] == null ? "" : xnl.Attributes["CabinetNo"].Value;
                    cabinet.CabinetPanelNo = xnl.Attributes["CabinetPanelNo"] == null ? "" : xnl.Attributes["CabinetPanelNo"].Value;
                    cabinet.PositionNumber = xnl.Attributes["PositionNumber"] == null ? "" : xnl.Attributes["PositionNumber"].Value;
                    cabinet.Name = xnl.Attributes["Name"] == null ? "" : xnl.Attributes["Name"].Value;
                    cabinet.Id = xnl.Attributes["Id"] == null ? "" : xnl.Attributes["Id"].Value;
                    cabinet.Series = xnl.Attributes["Series"] == null ? "" : xnl.Attributes["Series"].Value;  //20180416
                    cabinet.Length = xnl.Attributes["Length"] == null ? "" : xnl.Attributes["Length"].Value;
                    cabinet.Width = xnl.Attributes["Width"] == null ? "" : xnl.Attributes["Width"].Value;
                    cabinet.Height = xnl.Attributes["Height"] == null ? "" : xnl.Attributes["Height"].Value;
                    cabinet.SubType = xnl.Attributes["SubType"] == null ? "" : xnl.Attributes["SubType"].Value;  //20180416
                    cabinet.Material = xnl.Attributes["Material"] == null ? "" : xnl.Attributes["Material"].Value; //20180416
                    cabinet.BasicMaterial = xnl.Attributes["BasicMaterial"] == null ? "" : xnl.Attributes["BasicMaterial"].Value; //20180416
                    cabinet.Model = xnl.Attributes["Model"] == null ? "" : xnl.Attributes["Model"].Value; //20180416
                    cabinet.CraftMark = xnl.Attributes["CraftMark"] == null ? "" : xnl.Attributes["CraftMark"].Value;  //20180416
                    cabinet.PartNumber = xnl.Attributes["PartNumber"] == null ? "" : xnl.Attributes["PartNumber"].Value;  //20180416
                    cabinet.GroupName = xnl.Attributes["GroupName"] == null ? "" : xnl.Attributes["GroupName"].Value;  //20180416
                    cabinet.RoomName = xnl.Attributes["RoomName"] == null ? "" : xnl.Attributes["RoomName"].Value;
                    cabinet.OrderDate = xnl.Attributes["OrderDate"] == null ? "" : xnl.Attributes["OrderDate"].Value;
                    cabinet.Designer = xnl.Attributes["Designer"] == null ? "" : xnl.Attributes["Designer"].Value;
                    cabinet.CustomId = xnl.Attributes["CustomId"] == null ? "" : xnl.Attributes["CustomId"].Value;
                    cabinet.BatchNo = xnl.Attributes["BatchNo"] == null ? "" : xnl.Attributes["BatchNo"].Value;
                    cabinet.ThinEdgeValue = xnl.Attributes["ThinEdgeValue"] == null ? "" : xnl.Attributes["ThinEdgeValue"].Value;
                    cabinet.ThickEdgeValue = xnl.Attributes["ThickEdgeValue"] == null ? "" : xnl.Attributes["ThickEdgeValue"].Value;
                    cabinet.OrderNo = xnl.Attributes["OrderNo"] == null ? "" : xnl.Attributes["OrderNo"].Value;
                    cabinet.AlongSys = xnl.Attributes["AlongSys"] == null ? "" : xnl.Attributes["AlongSys"].Value;

                    for (int xnpanelnum = 0; xnpanelnum < xnl.ChildNodes.Count; xnpanelnum++) //20180627
                    {

                        #region 20180914 若有抽屉，抽屉数量的写入
                        if (xnl.Attributes["SubType"] != null && xnl.Attributes["SubType"].Value == "drawer")//20180914
                        {
                            string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                            IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                            if (inifile.ExistINIFile())
                            {
                                int lastdrawernum = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                inifile.IniWriteValue("DrawerNum", "draw", (lastdrawernum + 1).ToString());
                            }
                            else
                            {
                                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                return;
                            }
                        }
                        #endregion

                        foreach (XmlNode xnpanel in xnl.ChildNodes[xnpanelnum].ChildNodes)
                        {
                            if (xnpanel.Name == "Panel")
                            {
                                Panel panel = new Panel();
                                panel.LoadFromXmlNode(xnpanel, cabinet);

                                if ((Math.Abs(double.Parse(panel.Thickness) - 5) < 0.01) ||
                                    (Math.Abs(double.Parse(panel.Thickness) - 8) < 0.01) ||
                                    (Math.Abs(double.Parse(panel.Thickness) - 15) < 0.01) ||
                                    (Math.Abs(double.Parse(panel.Thickness) - 18) < 0.01) ||
                                    (Math.Abs(double.Parse(panel.Thickness) - 22) < 0.01) ||  // 说有22厚的门板
                                    (Math.Abs(double.Parse(panel.Thickness) - 25) < 0.01) ||
                                    (Math.Abs(double.Parse(panel.Thickness) - 35) < 0.01) ||
                                    (Math.Abs(double.Parse(panel.Thickness) - 50) < 0.01) ||
                                    panel.CabinetType.Contains("door") || panel.SubType.Contains("door"))
                                    panel.thicknessflag = true;

                                if (!panel.SubType.Contains("hardware") && !panel.CabinetType.Contains("hardware")) //20180714
                                {
                                    #region 20180914 若有抽屉,则抽屉板件的读取标识
                                    if (xnl.Attributes["SubType"] != null && xnl.Attributes["SubType"].Value == "drawer")
                                    {
                                        string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                        IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                        if (inifile.ExistINIFile())
                                        {
                                            panel.drawer = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));

                                        }
                                        else
                                        {
                                            MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                            return;
                                        }
                                    }
                                    #endregion

                                    if (!panel.MulPanels && panel.thicknessflag) //20180714
                                    {
                                        string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                        IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                        if (inifile.ExistINIFile())
                                        {
                                            panel.SMAEX = inifile.IniReadValue("SAMEX", "ps");

                                        }
                                        else
                                        {
                                            MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                            return;
                                        }

                                        if (panel.SMAEX.Equals("1") && panel.cabinet.CabinetNo != string.Empty) //20181228
                                            panel.Name = panel.cabinet.CabinetNo + "-" + panel.Name;

                                        cabinet.Panellist.Add(panel);
                                        //if (panel.CabinetNo == "" && double.Parse(xnpanel.Attributes["Thickness"].Value) < 50.1)  //过滤功能件节点里的数据  20180329    //20180418
                                        //    cabinet.Panellist.Add(panel);
                                        //else if (panel.Name.Contains("抽面") || panel.Category.Contains("Door"))   //  但如果节点里有门或抽面则不需要过滤 20180409  //发现门 应该从属性里判断 宋新刚 20180622
                                        //    cabinet.Panellist.Add(panel);
                                    }
                                }


                            }

                        }

                        foreach (XmlNode xnmental in xnl.ChildNodes[xnpanelnum].ChildNodes)
                        {
                            if (xnmental.Name == "Metal")
                            {
                                Metal metal = new Metal();
                                metal.LoadFromXmlNode(xnmental);
                                cabinet.Metallist.Add(metal);
                            }
                        }
                    }

                    Cabinetlist.Add(cabinet);

                    #endregion

                    #region 20180912 组合柜里面的柜子将节点再提到上一层 重新对cabinet赋值提取
                    foreach (XmlNode Group_down_xnl in xnl) //20180627
                    {
                        if (Group_down_xnl.Name == "Cabinet")
                        {
                            Cabinet Group_down_cabinet = new Cabinet();

                            Group_down_cabinet.CabinetNo = Group_down_xnl.Attributes["CabinetNo"] == null ? "" : Group_down_xnl.Attributes["CabinetNo"].Value;
                            Group_down_cabinet.CabinetPanelNo = Group_down_xnl.Attributes["CabinetPanelNo"] == null ? "" : Group_down_xnl.Attributes["CabinetPanelNo"].Value;
                            Group_down_cabinet.PositionNumber = Group_down_xnl.Attributes["PositionNumber"] == null ? "" : Group_down_xnl.Attributes["PositionNumber"].Value;
                            Group_down_cabinet.Name = Group_down_xnl.Attributes["Name"] == null ? "" : Group_down_xnl.Attributes["Name"].Value;
                            Group_down_cabinet.Id = Group_down_xnl.Attributes["Id"] == null ? "" : Group_down_xnl.Attributes["Id"].Value;
                            Group_down_cabinet.Series = Group_down_xnl.Attributes["Series"] == null ? "" : Group_down_xnl.Attributes["Series"].Value;  //20180416
                            Group_down_cabinet.Length = Group_down_xnl.Attributes["Length"] == null ? "" : Group_down_xnl.Attributes["Length"].Value;
                            Group_down_cabinet.Width = Group_down_xnl.Attributes["Width"] == null ? "" : Group_down_xnl.Attributes["Width"].Value;
                            Group_down_cabinet.Height = Group_down_xnl.Attributes["Height"] == null ? "" : Group_down_xnl.Attributes["Height"].Value;
                            Group_down_cabinet.SubType = Group_down_xnl.Attributes["SubType"] == null ? "" : Group_down_xnl.Attributes["SubType"].Value;  //20180416
                            Group_down_cabinet.Material = Group_down_xnl.Attributes["Material"] == null ? "" : Group_down_xnl.Attributes["Material"].Value; //20180416
                            Group_down_cabinet.BasicMaterial = Group_down_xnl.Attributes["BasicMaterial"] == null ? "" : Group_down_xnl.Attributes["BasicMaterial"].Value; //20180416
                            Group_down_cabinet.Model = Group_down_xnl.Attributes["Model"] == null ? "" : Group_down_xnl.Attributes["Model"].Value; //20180416
                            Group_down_cabinet.CraftMark = Group_down_xnl.Attributes["CraftMark"] == null ? "" : Group_down_xnl.Attributes["CraftMark"].Value;  //20180416
                            Group_down_cabinet.PartNumber = Group_down_xnl.Attributes["PartNumber"] == null ? "" : Group_down_xnl.Attributes["PartNumber"].Value;  //20180416
                            Group_down_cabinet.GroupName = Group_down_xnl.Attributes["GroupName"] == null ? "" : Group_down_xnl.Attributes["GroupName"].Value;  //20180416
                            Group_down_cabinet.RoomName = Group_down_xnl.Attributes["RoomName"] == null ? "" : Group_down_xnl.Attributes["RoomName"].Value;
                            Group_down_cabinet.OrderDate = Group_down_xnl.Attributes["OrderDate"] == null ? "" : Group_down_xnl.Attributes["OrderDate"].Value;
                            Group_down_cabinet.Designer = Group_down_xnl.Attributes["Designer"] == null ? "" : Group_down_xnl.Attributes["Designer"].Value;
                            Group_down_cabinet.CustomId = Group_down_xnl.Attributes["CustomId"] == null ? "" : Group_down_xnl.Attributes["CustomId"].Value;
                            Group_down_cabinet.BatchNo = Group_down_xnl.Attributes["BatchNo"] == null ? "" : Group_down_xnl.Attributes["BatchNo"].Value;
                            Group_down_cabinet.ThinEdgeValue = Group_down_xnl.Attributes["ThinEdgeValue"] == null ? "" : Group_down_xnl.Attributes["ThinEdgeValue"].Value;
                            Group_down_cabinet.ThickEdgeValue = Group_down_xnl.Attributes["ThickEdgeValue"] == null ? "" : Group_down_xnl.Attributes["ThickEdgeValue"].Value;
                            Group_down_cabinet.OrderNo = Group_down_xnl.Attributes["OrderNo"] == null ? "" : Group_down_xnl.Attributes["OrderNo"].Value;
                            Group_down_cabinet.AlongSys = Group_down_xnl.Attributes["AlongSys"] == null ? "" : Group_down_xnl.Attributes["AlongSys"].Value;

                            for (int Grpup_down_xnpanelnum = 0; Grpup_down_xnpanelnum < Group_down_xnl.ChildNodes.Count; Grpup_down_xnpanelnum++)
                            {
                                #region 20180914 若有抽屉，抽屉数量的写入
                                if (Group_down_xnl.Attributes["SubType"] != null && Group_down_xnl.Attributes["SubType"].Value == "drawer")//20180914
                                {
                                    string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                    IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                    if (inifile.ExistINIFile())
                                    {
                                        int lastdrawernum = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                        inifile.IniWriteValue("DrawerNum", "draw", (lastdrawernum + 1).ToString());
                                    }
                                    else
                                    {
                                        MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                        return;
                                    }
                                }
                                #endregion

                                foreach (XmlNode xnpanel in Group_down_xnl.ChildNodes[Grpup_down_xnpanelnum].ChildNodes)
                                {
                                    if (xnpanel.Name == "Panel")
                                    {
                                        Panel panel = new Panel();
                                        panel.LoadFromXmlNode(xnpanel, Group_down_cabinet);

                                        if ((Math.Abs(double.Parse(panel.Thickness) - 5) < 0.01) ||
                                            (Math.Abs(double.Parse(panel.Thickness) - 8) < 0.01) ||
                                            (Math.Abs(double.Parse(panel.Thickness) - 15) < 0.01) ||
                                            (Math.Abs(double.Parse(panel.Thickness) - 18) < 0.01) ||
                                            (Math.Abs(double.Parse(panel.Thickness) - 22) < 0.01) ||  // 说有22厚的门板
                                            (Math.Abs(double.Parse(panel.Thickness) - 25) < 0.01) ||
                                            (Math.Abs(double.Parse(panel.Thickness) - 35) < 0.01) ||
                                            (Math.Abs(double.Parse(panel.Thickness) - 50) < 0.01) ||
                                            panel.CabinetType.Contains("door") || panel.SubType.Contains("door"))
                                            panel.thicknessflag = true;

                                        if (!panel.SubType.Contains("hardware") && !panel.CabinetType.Contains("hardware")) //20180714
                                        {
                                            #region 20180914 若有抽屉,则抽屉板件的读取标识
                                            if (Group_down_xnl.Attributes["SubType"] != null && Group_down_xnl.Attributes["SubType"].Value == "drawer")
                                            {
                                                string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                                IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                                if (inifile.ExistINIFile())
                                                {
                                                    panel.drawer = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                                }
                                                else
                                                {
                                                    MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                                    return;
                                                }
                                            }
                                            #endregion

                                            if (!panel.MulPanels && panel.thicknessflag) //20180714
                                            {
                                                string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                                IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                                if (inifile.ExistINIFile())
                                                {
                                                    panel.SMAEX = inifile.IniReadValue("SAMEX", "ps");

                                                }
                                                else
                                                {
                                                    MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                                    return;
                                                }

                                                if (panel.SMAEX.Equals("1") && panel.cabinet.CabinetNo != string.Empty) //20181228
                                                    panel.Name = panel.cabinet.CabinetNo + "-" + panel.Name;

                                                Group_down_cabinet.Panellist.Add(panel);
                                                //if (panel.CabinetNo == "" && double.Parse(xnpanel.Attributes["Thickness"].Value) < 50.1)  //过滤功能件节点里的数据  20180329    //20180418
                                                //    cabinet.Panellist.Add(panel);
                                                //else if (panel.Name.Contains("抽面") || panel.Category.Contains("Door"))   //  但如果节点里有门或抽面则不需要过滤 20180409  //发现门 应该从属性里判断 宋新刚 20180622
                                                //    cabinet.Panellist.Add(panel);
                                            }
                                        }


                                    }

                                }

                                foreach (XmlNode xnmental in Group_down_xnl.ChildNodes[Grpup_down_xnpanelnum].ChildNodes)//20190104
                                {
                                    if (xnmental.Name == "Metal")
                                    {
                                        Metal metal = new Metal();
                                        metal.LoadFromXmlNode(xnmental);
                                        Group_down_cabinet.Metallist.Add(metal);
                                    }
                                }
                            }
                            Cabinetlist.Add(Group_down_cabinet);
                        }
                    }
                    #endregion
                }

            }


        }
        public bool OutputCSV(string csvpath)
        {
            return true;
        }
        public bool OutputCSV(string xmlpath, string csvpath)
        {
            LoadFromXML(xmlpath);
            OutputCSV(csvpath);
            return true;
        }


    }


    public class Cabinet : Panel
    {
        //public string CabinetNo { get; set; }
        //public string CabinetPanelNo { get; set; }
        //public string PositionNumber { get; set; }
        //public string Name { get; set; }
        public string Id { get; set; }
        //public string Series { get; set; }
        //public string Length { get; set; }
        //public string Width { get; set; }
        public string Height { get; set; }
        //public string SubType { get; set; }
        //public string Material { get; set; }
        //public string BasicMaterial { get; set; }
        //public string Model { get; set; }
        //public string CraftMark { get; set; }
        //public string PartNumber { get; set; }
        public string GroupName { get; set; }
        public string RoomName { get; set; }
        public string OrderDate { get; set; }
        public string Designer { get; set; }
        public string CustomId { get; set; }
        public string BatchNo { get; set; }
        public string ThinEdgeValue { get; set; }
        public string ThickEdgeValue { get; set; }
        public string OrderNo { get; set; }
        public string AlongSys { get; set; }
        public string DrawNum { get; set; } //20180914

        public List<Panel> Panellist = new List<Panel>();
        public List<Metal> Metallist = new List<Metal>();

    }

    public class Panel
    {
        public string PositionNumber { get; set; }
        public string IsProduce { get; set; }
        public string Thickness { get; set; }
        public string CraftMark { get; set; }
        public string SubType { get; set; }
        public string Length { get; set; }
        public string Width { get; set; }
        public string ID { get; set; }
        public string Name { get; set; }
        public string Material { get; set; }
        public string MaterialId { get; set; }
        public string BaseMaterialCategoryId { get; set; }
        public string MaterialCategoryId { get; set; }
        public string Model { get; set; }
        public string CabinetType { get; set; } //20180714 
        public string Type { get; set; }
        public string edgeMaterial { get; set; }
        public string StandardCategory { get; set; }
        public string IsAccurate { get; set; }
        public string MachiningPoint { get; set; }
        public string Grain { get; set; }
        public string ProdutionNo { get; set; }
        public string ProductionName { get; set; }
        public string Face5ID { get; set; }
        public string Face6ID { get; set; }
        public string clerk { get; set; }
        public string PkgNo { get; set; }
        public string BasicMaterial { get; set; }
        public string PartNumber { get; set; }
        public string DoorDirection { get; set; }
        public string Category { get; set; }
        public string thickLength { get; set; }
        public string thinLength { get; set; }
        public string customLength { get; set; }
        public string slotDis { get; set; }
        public string slotFace { get; set; }
        public string HasHorizontalHole { get; set; }
        public string ActualLength { get; set; }
        public string ActualWidth { get; set; }
        public string Series { get; set; }
        public bool MulPanels { get; set; } //20180714
        public bool thicknessflag { get; set; } //20180718
        public int drawer { get; set; } //20180914
        public string SMAEX { get; set; } //20180915

        public string CabinetNo { get; set; }
        public string CabinetPanelNo { get; set; }

        public Cabinet cabinet { get; set; }

        public List<Edge> Edgelist = new List<Edge>();
        public List<Machining> Machininglist = new List<Machining>();


        internal void LoadFromXmlNode(XmlNode xnpanel, Cabinet parentcabinet)
        {
            try
            {
                cabinet = parentcabinet;  //20180714
                MulPanels = false;

                for (int xmlsubnum = 0; xmlsubnum < xnpanel.ChildNodes.Count; xmlsubnum++)//20180627
                {
                    if (xnpanel.ChildNodes[xmlsubnum].Name == "Panels")
                    {
                        #region 20180914 若有抽屉，抽屉数量的写入
                        if (xnpanel.Attributes["SubType"] != null && xnpanel.Attributes["SubType"].Value == "drawer")//20180914
                        {
                            string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                            IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                            if (inifile.ExistINIFile())
                            {
                                int lastdrawernum = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                inifile.IniWriteValue("DrawerNum", "draw", (lastdrawernum + 1).ToString());
                            }
                            else
                            {
                                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                return;
                            }
                        }
                        #endregion

                        foreach (XmlNode xmlsub in xnpanel.ChildNodes[xmlsubnum].ChildNodes)
                        {
                            if (xmlsub.Name == "Panel")
                            {
                                Panel childpanel = new Panel();
                                childpanel.LoadFromXmlChildNode(xmlsub, parentcabinet); //20180914

                                if ((Math.Abs(double.Parse(childpanel.Thickness) - 5) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 8) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 15) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 18) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 22) < 0.01) ||  // 说有22厚的门板
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 25) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 35) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 50) < 0.01) ||
                                    childpanel.CabinetType.Contains("door") || childpanel.SubType.Contains("door"))
                                    childpanel.thicknessflag = true;

                                if (!childpanel.SubType.Contains("hardware") && !childpanel.CabinetType.Contains("hardware")) //20180714
                                {
                                    #region 20180914 若有抽屉,则抽屉板件的读取标识
                                    if (xnpanel.Attributes["SubType"] != null && xnpanel.Attributes["SubType"].Value == "drawer")
                                    {
                                        string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                        IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                        if (inifile.ExistINIFile())
                                        {
                                            childpanel.drawer = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                        }
                                        else
                                        {
                                            MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                            return;
                                        }
                                    }
                                    #endregion

                                    if (!childpanel.MulPanels && childpanel.thicknessflag)
                                    {
                                        string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                        IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                        if (inifile.ExistINIFile())
                                        {
                                            childpanel.SMAEX = inifile.IniReadValue("SAMEX", "ps");

                                        }
                                        else
                                        {
                                            MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                            return;
                                        }

                                        if (childpanel.SMAEX.Equals("1") && childpanel.cabinet.CabinetNo != string.Empty) //20181228
                                            childpanel.Name = childpanel.cabinet.CabinetNo + "-" + childpanel.Name;

                                        cabinet.Panellist.Add(childpanel);
                                    }
                                }
                            }
                        }
                        MulPanels = true;
                    }
                }
                //    }

                PositionNumber = xnpanel.Attributes["PositionNumber"] == null ? "" : xnpanel.Attributes["PositionNumber"].Value;
                IsProduce = xnpanel.Attributes["IsProduce"] == null ? "" : xnpanel.Attributes["IsProduce"].Value;
                Thickness = xnpanel.Attributes["Thickness"] == null ? "" : xnpanel.Attributes["Thickness"].Value;
                CraftMark = xnpanel.Attributes["CraftMark"] == null ? "" : xnpanel.Attributes["CraftMark"].Value;
                SubType = xnpanel.Attributes["SubType"] == null ? "" : xnpanel.Attributes["SubType"].Value;
                Length = xnpanel.Attributes["Length"] == null ? "" : xnpanel.Attributes["Length"].Value;
                Width = xnpanel.Attributes["Width"] == null ? "" : xnpanel.Attributes["Width"].Value;
                ID = xnpanel.Attributes["ID"] == null ? "" : xnpanel.Attributes["ID"].Value;

                if (parentcabinet.CabinetNo != string.Empty) //20181228
                Name = parentcabinet.CabinetNo + "-" + xnpanel.Attributes["Name"] == null ? "" : xnpanel.Attributes["Name"].Value;
                else
                Name = xnpanel.Attributes["Name"] == null ? "" : xnpanel.Attributes["Name"].Value;

                Material = xnpanel.Attributes["Material"] == null ? "" : xnpanel.Attributes["Material"].Value;
                MaterialId = xnpanel.Attributes["MaterialId"] == null ? "" : xnpanel.Attributes["MaterialId"].Value;
                BaseMaterialCategoryId = xnpanel.Attributes["BaseMaterialCategoryId"] == null ? "" : xnpanel.Attributes["BaseMaterialCategoryId"].Value;
                MaterialCategoryId = xnpanel.Attributes["MaterialCategoryId"] == null ? "" : xnpanel.Attributes["MaterialCategoryId"].Value;
                Model = xnpanel.Attributes["Model"] == null ? "" : xnpanel.Attributes["Model"].Value;
                CabinetType = xnpanel.Attributes["CabinetType"] == null ? "" : xnpanel.Attributes["CabinetType"].Value; //20180714
                Type = xnpanel.Attributes["Type"] == null ? "" : xnpanel.Attributes["Type"].Value;
                edgeMaterial = xnpanel.Attributes["edgeMaterial"] == null ? "" : xnpanel.Attributes["edgeMaterial"].Value;
                StandardCategory = xnpanel.Attributes["StandardCategory"] == null ? "" : xnpanel.Attributes["StandardCategory"].Value;
                IsAccurate = xnpanel.Attributes["IsAccurate"] == null ? "" : xnpanel.Attributes["IsAccurate"].Value;
                MachiningPoint = xnpanel.Attributes["MachiningPoint"] == null ? "" : xnpanel.Attributes["MachiningPoint"].Value;
                Grain = xnpanel.Attributes["Grain"] == null ? "" : xnpanel.Attributes["Grain"].Value;
                ProdutionNo = xnpanel.Attributes["ProdutionNo"] == null ? "" : xnpanel.Attributes["ProdutionNo"].Value;
                ProductionName = xnpanel.Attributes["ProductionName"] == null ? "" : xnpanel.Attributes["ProductionName"].Value;
                
                Face5ID = "P" + ID.Substring((ID.Length - 3), 3) + "X";  //20180409
                Face6ID = "P" + ID.Substring((ID.Length - 3), 3) + "Y";  //20180409

                clerk = xnpanel.Attributes["clerk"] == null ? "" : xnpanel.Attributes["clerk"].Value;
                PkgNo = xnpanel.Attributes["PkgNo"] == null ? "" : xnpanel.Attributes["PkgNo"].Value;
                BasicMaterial = xnpanel.Attributes["BasicMaterial"] == null ? "" : xnpanel.Attributes["BasicMaterial"].Value;
                PartNumber = xnpanel.Attributes["PartNumber"] == null ? "" : xnpanel.Attributes["PartNumber"].Value;

                if (PartNumber.ToUpper().Contains("Y"))   //20180417
                {
                    Name = Name + "-Y";
                }

                DoorDirection = xnpanel.Attributes["DoorDirection"] == null ? "" : xnpanel.Attributes["DoorDirection"].Value;
                Category = xnpanel.Attributes["Category"] == null ? "" : xnpanel.Attributes["Category"].Value;
                thickLength = xnpanel.Attributes["thickLength"] == null ? "" : xnpanel.Attributes["thickLength"].Value;  //20180329
                thinLength = xnpanel.Attributes["thinLength"] == null ? "" : xnpanel.Attributes["thinLength"].Value;    //20180329
                customLength = xnpanel.Attributes["customLength"] == null ? "" : xnpanel.Attributes["customLength"].Value;  //20180329
                slotDis = xnpanel.Attributes["slotDis"] == null ? "" : xnpanel.Attributes["slotDis"].Value;
                slotFace = xnpanel.Attributes["slotFace"] == null ? "" : xnpanel.Attributes["slotFace"].Value;
                HasHorizontalHole = xnpanel.Attributes["HasHorizontalHole"] == null ? "" : xnpanel.Attributes["HasHorizontalHole"].Value;
                ActualLength = xnpanel.Attributes["ActualLength"] == null ? "" : xnpanel.Attributes["ActualLength"].Value;
                ActualWidth = xnpanel.Attributes["ActualWidth"] == null ? "" : xnpanel.Attributes["ActualWidth"].Value;
                Series = xnpanel.Attributes["Series"] == null ? "" : xnpanel.Attributes["Series"].Value;

                CabinetNo = xnpanel.Attributes["CabinetNo"] == null ? "" : xnpanel.Attributes["CabinetNo"].Value;
                CabinetPanelNo = xnpanel.Attributes["CabinetPanelNo"] == null ? "" : xnpanel.Attributes["CabinetPanelNo"].Value;

                for (int xmlnum = 0; xmlnum < xnpanel.ChildNodes.Count; xmlnum++)   //20180525
                {
                    foreach (XmlNode xnedge in xnpanel.ChildNodes[xmlnum].ChildNodes)
                    {
                        if (xnedge.Name == "Edge")
                        {
                            Edge edge = new Edge();
                            edge.LoadFromXmlNode(xnedge);
                            Edgelist.Add(edge);
                        }
                    }

                    foreach (XmlNode xnmachine in xnpanel.ChildNodes[xmlnum].ChildNodes)
                    {
                        if (xnmachine.Name == "Machining")
                        {
                            Machining machining = new Machining();
                            machining.LoadFromXmlNode(xnmachine);
                            if (machining.Type == "1" || machining.Type == "2" || machining.Type == "3")  //20180703  发现有除了1、2、3类型外的数字 则过滤
                                Machininglist.Add(machining);
                        }
                    }
                }

            }
            catch
            {
                throw new NotImplementedException(ID);
            }
        }

        private void LoadFromXmlChildNode(XmlNode xmlchildnode, Cabinet parentcabinet)  //20180329
        {
            try
            {
                cabinet = parentcabinet;
                MulPanels = false;
                //if (xmlchildnode.Attributes["CabinetNo"] != null)
                //{
                for (int xmlsubnum = 0; xmlsubnum < xmlchildnode.ChildNodes.Count; xmlsubnum++)//20180627
                {
                    if (xmlchildnode.ChildNodes[xmlsubnum].Name == "Panels")
                    {
                        #region 20180914 若有抽屉，抽屉数量的写入
                        if (xmlchildnode.Attributes["SubType"] != null && xmlchildnode.Attributes["SubType"].Value == "drawer")
                        {
                            string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                            IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                            if (inifile.ExistINIFile())
                            {
                                int lastdrawernum = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                inifile.IniWriteValue("DrawerNum", "draw", (lastdrawernum + 1).ToString());
                            }
                            else
                            {
                                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                return;
                            }
                        }
                        #endregion

                        foreach (XmlNode xmlsub in xmlchildnode.ChildNodes[xmlsubnum].ChildNodes)
                        {
                            if (xmlsub.Name == "Panel")
                            {
                                Panel childpanel = new Panel();
                                childpanel.LoadFromXmlsubChildNode(xmlsub, parentcabinet);

                                if ((Math.Abs(double.Parse(childpanel.Thickness) - 5) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 8) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 15) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 18) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 22) < 0.01) ||  // 说有22厚的门板
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 25) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 35) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 50) < 0.01) ||
                                    childpanel.CabinetType.Contains("door") || childpanel.SubType.Contains("door"))
                                    childpanel.thicknessflag = true;

                                if (!childpanel.SubType.Contains("hardware") && !childpanel.CabinetType.Contains("hardware")) //20180714
                                {
                                    #region 20180914 若有抽屉,则抽屉板件的读取标识
                                    if (xmlchildnode.Attributes["SubType"] != null && xmlchildnode.Attributes["SubType"].Value == "drawer")
                                    {
                                        string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                        IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                        if (inifile.ExistINIFile())
                                        {
                                            childpanel.drawer = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                        }
                                        else
                                        {
                                            MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                            return;
                                        }
                                    }

                                    #endregion

                                    if (!childpanel.MulPanels && childpanel.thicknessflag)
                                    {
                                        string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                        IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                        if (inifile.ExistINIFile())
                                        {
                                            childpanel.SMAEX = inifile.IniReadValue("SAMEX", "ps");

                                        }
                                        else
                                        {
                                            MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                            return;
                                        }

                                        if (childpanel.SMAEX.Equals("1") && childpanel.cabinet.CabinetNo != string.Empty) //20181228
                                            childpanel.Name = childpanel.cabinet.CabinetNo + "-" + childpanel.Name;


                                        cabinet.Panellist.Add(childpanel);
                                        //    if (childpanel.CabinetNo == "" && double.Parse(xmlsub.Attributes["Thickness"].Value) < 50.1)      //  子节点中 也是需要判断是否有组件20180410   //20180418
                                        //    cabinet.Panellist.Add(childpanel);
                                        //else if (childpanel.Category.Contains("Door") || childpanel.Name.Contains("抽面"))  //发现门 应该从属性里判断 宋新刚 20180622
                                        //    cabinet.Panellist.Add(childpanel);
                                    }
                                }
                            }
                        }
                        MulPanels = true;
                    }
                }
                // }
                PositionNumber = xmlchildnode.Attributes["PositionNumber"] == null ? "" : xmlchildnode.Attributes["PositionNumber"].Value;
                IsProduce = xmlchildnode.Attributes["IsProduce"] == null ? "" : xmlchildnode.Attributes["IsProduce"].Value;
                Thickness = xmlchildnode.Attributes["Thickness"] == null ? "" : xmlchildnode.Attributes["Thickness"].Value;
                CraftMark = xmlchildnode.Attributes["CraftMark"] == null ? "" : xmlchildnode.Attributes["CraftMark"].Value;
                SubType = xmlchildnode.Attributes["SubType"] == null ? "" : xmlchildnode.Attributes["SubType"].Value;
                Length = xmlchildnode.Attributes["Length"] == null ? "" : xmlchildnode.Attributes["Length"].Value;
                Width = xmlchildnode.Attributes["Width"] == null ? "" : xmlchildnode.Attributes["Width"].Value;
                ID = xmlchildnode.Attributes["ID"] == null ? "" : xmlchildnode.Attributes["ID"].Value;
                Name = xmlchildnode.Attributes["Name"] == null ? "" : xmlchildnode.Attributes["Name"].Value;
                Material = xmlchildnode.Attributes["Material"] == null ? "" : xmlchildnode.Attributes["Material"].Value;
                MaterialId = xmlchildnode.Attributes["MaterialId"] == null ? "" : xmlchildnode.Attributes["MaterialId"].Value;
                BaseMaterialCategoryId = xmlchildnode.Attributes["BaseMaterialCategoryId"] == null ? "" : xmlchildnode.Attributes["BaseMaterialCategoryId"].Value;
                MaterialCategoryId = xmlchildnode.Attributes["MaterialCategoryId"] == null ? "" : xmlchildnode.Attributes["MaterialCategoryId"].Value;
                Model = xmlchildnode.Attributes["Model"] == null ? "" : xmlchildnode.Attributes["Model"].Value;
                CabinetType = xmlchildnode.Attributes["CabinetType"] == null ? "" : xmlchildnode.Attributes["CabinetType"].Value; //20180714
                Type = xmlchildnode.Attributes["Type"] == null ? "" : xmlchildnode.Attributes["Type"].Value;
                edgeMaterial = xmlchildnode.Attributes["edgeMaterial"] == null ? "" : xmlchildnode.Attributes["edgeMaterial"].Value;
                StandardCategory = xmlchildnode.Attributes["StandardCategory"] == null ? "" : xmlchildnode.Attributes["StandardCategory"].Value;
                IsAccurate = xmlchildnode.Attributes["IsAccurate"] == null ? "" : xmlchildnode.Attributes["IsAccurate"].Value;
                MachiningPoint = xmlchildnode.Attributes["MachiningPoint"] == null ? "" : xmlchildnode.Attributes["MachiningPoint"].Value;
                Grain = xmlchildnode.Attributes["Grain"] == null ? "" : xmlchildnode.Attributes["Grain"].Value;
                ProdutionNo = xmlchildnode.Attributes["ProdutionNo"] == null ? "" : xmlchildnode.Attributes["ProdutionNo"].Value;
                ProductionName = xmlchildnode.Attributes["ProductionName"] == null ? "" : xmlchildnode.Attributes["ProductionName"].Value;

                Face5ID = "P" + ID.Substring((ID.Length - 3), 3) + "X";  //20180409
                Face6ID = "P" + ID.Substring((ID.Length - 3), 3) + "Y";  //20180409

                clerk = xmlchildnode.Attributes["clerk"] == null ? "" : xmlchildnode.Attributes["clerk"].Value;
                PkgNo = xmlchildnode.Attributes["PkgNo"] == null ? "" : xmlchildnode.Attributes["PkgNo"].Value;
                BasicMaterial = xmlchildnode.Attributes["BasicMaterial"] == null ? "" : xmlchildnode.Attributes["BasicMaterial"].Value;
                PartNumber = xmlchildnode.Attributes["PartNumber"] == null ? "" : xmlchildnode.Attributes["PartNumber"].Value;

                if (PartNumber.ToUpper().Contains("Y"))   //20180417
                {
                    Name = Name + "-Y";
                }

                DoorDirection = xmlchildnode.Attributes["DoorDirection"] == null ? "" : xmlchildnode.Attributes["DoorDirection"].Value;
                Category = xmlchildnode.Attributes["Category"] == null ? "" : xmlchildnode.Attributes["Category"].Value;
                thickLength = xmlchildnode.Attributes["thickLength"] == null ? "" : xmlchildnode.Attributes["thickLength"].Value;  // 20180410
                thinLength = xmlchildnode.Attributes["thinLength"] == null ? "" : xmlchildnode.Attributes["thinLength"].Value;  // 20180410
                customLength = xmlchildnode.Attributes["customLength"] == null ? "" : xmlchildnode.Attributes["customLength"].Value; // 20180410
                slotDis = xmlchildnode.Attributes["slotDis"] == null ? "" : xmlchildnode.Attributes["slotDis"].Value;
                slotFace = xmlchildnode.Attributes["slotFace"] == null ? "" : xmlchildnode.Attributes["slotFace"].Value;
                HasHorizontalHole = xmlchildnode.Attributes["HasHorizontalHole"] == null ? "" : xmlchildnode.Attributes["HasHorizontalHole"].Value;
                ActualLength = xmlchildnode.Attributes["ActualLength"] == null ? "" : xmlchildnode.Attributes["ActualLength"].Value;
                ActualWidth = xmlchildnode.Attributes["ActualWidth"] == null ? "" : xmlchildnode.Attributes["ActualWidth"].Value;
                Series = xmlchildnode.Attributes["Series"] == null ? "" : xmlchildnode.Attributes["Series"].Value;

                CabinetNo = xmlchildnode.Attributes["CabinetNo"] == null ? "" : xmlchildnode.Attributes["CabinetNo"].Value;
                CabinetPanelNo = xmlchildnode.Attributes["CabinetPanelNo"] == null ? "" : xmlchildnode.Attributes["CabinetPanelNo"].Value;

                for (int xmlnum = 0; xmlnum < xmlchildnode.ChildNodes.Count; xmlnum++)   //20180525
                {
                    foreach (XmlNode xnedge in xmlchildnode.ChildNodes[xmlnum].ChildNodes)
                    {
                        if (xnedge.Name == "Edge")
                        {
                            Edge edge = new Edge();
                            edge.LoadFromXmlNode(xnedge);
                            Edgelist.Add(edge);
                        }
                    }

                    foreach (XmlNode xnmachine in xmlchildnode.ChildNodes[xmlnum].ChildNodes)
                    {
                        if (xnmachine.Name == "Machining")
                        {
                            Machining machining = new Machining();
                            machining.LoadFromXmlNode(xnmachine);
                            if (machining.Type == "1" || machining.Type == "2" || machining.Type == "3")  //20180703  发现有除了1、2、3类型外的数字 则过滤
                                Machininglist.Add(machining);
                        }
                    }
                }
            }
            catch
            {
                throw new NotImplementedException();
            }
        }

        private void LoadFromXmlsubChildNode(XmlNode xmlsubsub, Cabinet parentcabinet)  //组件里含组件  20180329
        {
            try
            {
                cabinet = parentcabinet;

                MulPanels = false;
                //if (xmlsubsub.Attributes["CabinetNo"] != null)
                //{
                for (int xmlsubnum = 0; xmlsubnum < xmlsubsub.ChildNodes.Count; xmlsubnum++)//20180627
                {
                    if (xmlsubsub.ChildNodes[xmlsubnum].Name == "Panels")
                    {
                        #region 20180914 若有抽屉，抽屉数量的写入
                        if (xmlsubsub.Attributes["SubType"] != null && xmlsubsub.Attributes["SubType"].Value == "drawer")
                        {
                            string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                            IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                            if (inifile.ExistINIFile())
                            {
                                int lastdrawernum = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                inifile.IniWriteValue("DrawerNum", "draw", (lastdrawernum + 1).ToString());
                            }
                            else
                            {
                                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                return;
                            }
                        }
                        #endregion

                        foreach (XmlNode xmlsub in xmlsubsub.ChildNodes[xmlsubnum].ChildNodes)
                        {
                            if (xmlsub.Name == "Panel")
                            {
                                Panel childpanel = new Panel();
                                childpanel.LoadFromXmlsubsubChildNode(xmlsub, parentcabinet);

                                if ((Math.Abs(double.Parse(childpanel.Thickness) - 5) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 8) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 15) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 18) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 22) < 0.01) ||  // 说有22厚的门板
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 25) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 35) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 50) < 0.01) ||
                                    childpanel.CabinetType.Contains("door") || childpanel.SubType.Contains("door"))
                                    childpanel.thicknessflag = true;

                                if (!childpanel.SubType.Contains("hardware") && !childpanel.CabinetType.Contains("hardware")) //20180714
                                {
                                    #region 20180914 若有抽屉,则抽屉板件的读取标识
                                    if (xmlsubsub.Attributes["SubType"] != null && xmlsubsub.Attributes["SubType"].Value == "drawer")
                                    {
                                        string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                        IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                        if (inifile.ExistINIFile())
                                        {
                                            childpanel.drawer = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                        }
                                        else
                                        {
                                            MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                            return;
                                        }
                                    }
                                    #endregion

                                    if (!childpanel.MulPanels && childpanel.thicknessflag)
                                    {
                                        string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                        IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                        if (inifile.ExistINIFile())
                                        {
                                            childpanel.SMAEX = inifile.IniReadValue("SAMEX", "ps");

                                        }
                                        else
                                        {
                                            MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                            return;
                                        }

                                        if (childpanel.SMAEX.Equals("1") && childpanel.cabinet.CabinetNo != string.Empty) //20181228
                                            childpanel.Name = childpanel.cabinet.CabinetNo + "-" + childpanel.Name;

                                        cabinet.Panellist.Add(childpanel);
                                        //    if (childpanel.CabinetNo == "" && double.Parse(xmlsub.Attributes["Thickness"].Value) < 50.1)      //  子节点中 也是需要判断是否有组件20180410   //20180418
                                        //    cabinet.Panellist.Add(childpanel);
                                        //else if (childpanel.Category.Contains("Door") || childpanel.Name.Contains("抽面"))  //发现门 应该从属性里判断 宋新刚 20180622
                                        //    cabinet.Panellist.Add(childpanel);
                                    }
                                }
                            }
                        }
                        MulPanels = true;
                    }
                }
                //  }
                PositionNumber = xmlsubsub.Attributes["PositionNumber"] == null ? "" : xmlsubsub.Attributes["PositionNumber"].Value;
                IsProduce = xmlsubsub.Attributes["IsProduce"] == null ? "" : xmlsubsub.Attributes["IsProduce"].Value;
                Thickness = xmlsubsub.Attributes["Thickness"] == null ? "" : xmlsubsub.Attributes["Thickness"].Value;
                CraftMark = xmlsubsub.Attributes["CraftMark"] == null ? "" : xmlsubsub.Attributes["CraftMark"].Value;
                SubType = xmlsubsub.Attributes["SubType"] == null ? "" : xmlsubsub.Attributes["SubType"].Value;
                Length = xmlsubsub.Attributes["Length"] == null ? "" : xmlsubsub.Attributes["Length"].Value;
                Width = xmlsubsub.Attributes["Width"] == null ? "" : xmlsubsub.Attributes["Width"].Value;
                ID = xmlsubsub.Attributes["ID"] == null ? "" : xmlsubsub.Attributes["ID"].Value;
                Name = xmlsubsub.Attributes["Name"] == null ? "" : xmlsubsub.Attributes["Name"].Value;
                Material = xmlsubsub.Attributes["Material"] == null ? "" : xmlsubsub.Attributes["Material"].Value;
                MaterialId = xmlsubsub.Attributes["MaterialId"] == null ? "" : xmlsubsub.Attributes["MaterialId"].Value;
                BaseMaterialCategoryId = xmlsubsub.Attributes["BaseMaterialCategoryId"] == null ? "" : xmlsubsub.Attributes["BaseMaterialCategoryId"].Value;
                MaterialCategoryId = xmlsubsub.Attributes["MaterialCategoryId"] == null ? "" : xmlsubsub.Attributes["MaterialCategoryId"].Value;
                Model = xmlsubsub.Attributes["Model"] == null ? "" : xmlsubsub.Attributes["Model"].Value;
                CabinetType = xmlsubsub.Attributes["CabinetType"] == null ? "" : xmlsubsub.Attributes["CabinetType"].Value; //20180714
                Type = xmlsubsub.Attributes["Type"] == null ? "" : xmlsubsub.Attributes["Type"].Value;
                edgeMaterial = xmlsubsub.Attributes["edgeMaterial"] == null ? "" : xmlsubsub.Attributes["edgeMaterial"].Value;
                StandardCategory = xmlsubsub.Attributes["StandardCategory"] == null ? "" : xmlsubsub.Attributes["StandardCategory"].Value;
                IsAccurate = xmlsubsub.Attributes["IsAccurate"] == null ? "" : xmlsubsub.Attributes["IsAccurate"].Value;
                MachiningPoint = xmlsubsub.Attributes["MachiningPoint"] == null ? "" : xmlsubsub.Attributes["MachiningPoint"].Value;
                Grain = xmlsubsub.Attributes["Grain"] == null ? "" : xmlsubsub.Attributes["Grain"].Value;
                ProdutionNo = xmlsubsub.Attributes["ProdutionNo"] == null ? "" : xmlsubsub.Attributes["ProdutionNo"].Value;
                ProductionName = xmlsubsub.Attributes["ProductionName"] == null ? "" : xmlsubsub.Attributes["ProductionName"].Value;

                Face5ID = "P" + ID.Substring((ID.Length - 3), 3) + "X";  //20180409
                Face6ID = "P" + ID.Substring((ID.Length - 3), 3) + "Y";  //20180409

                clerk = xmlsubsub.Attributes["clerk"] == null ? "" : xmlsubsub.Attributes["clerk"].Value;
                PkgNo = xmlsubsub.Attributes["PkgNo"] == null ? "" : xmlsubsub.Attributes["PkgNo"].Value;
                BasicMaterial = xmlsubsub.Attributes["BasicMaterial"] == null ? "" : xmlsubsub.Attributes["BasicMaterial"].Value;
                PartNumber = xmlsubsub.Attributes["PartNumber"] == null ? "" : xmlsubsub.Attributes["PartNumber"].Value;

                if (PartNumber.ToUpper().Contains("Y"))   //20180417
                {
                    Name = Name + "-Y";
                }

                DoorDirection = xmlsubsub.Attributes["DoorDirection"] == null ? "" : xmlsubsub.Attributes["DoorDirection"].Value;
                Category = xmlsubsub.Attributes["Category"] == null ? "" : xmlsubsub.Attributes["Category"].Value;
                thickLength = xmlsubsub.Attributes["thickLength"] == null ? "" : xmlsubsub.Attributes["thickLength"].Value;
                thinLength = xmlsubsub.Attributes["thinLength"] == null ? "" : xmlsubsub.Attributes["thinLength"].Value;
                customLength = xmlsubsub.Attributes["customLength"] == null ? "" : xmlsubsub.Attributes["customLength"].Value;
                slotDis = xmlsubsub.Attributes["slotDis"] == null ? "" : xmlsubsub.Attributes["slotDis"].Value;
                slotFace = xmlsubsub.Attributes["slotFace"] == null ? "" : xmlsubsub.Attributes["slotFace"].Value;
                HasHorizontalHole = xmlsubsub.Attributes["HasHorizontalHole"] == null ? "" : xmlsubsub.Attributes["HasHorizontalHole"].Value;
                ActualLength = xmlsubsub.Attributes["ActualLength"] == null ? "" : xmlsubsub.Attributes["ActualLength"].Value;
                ActualWidth = xmlsubsub.Attributes["ActualWidth"] == null ? "" : xmlsubsub.Attributes["ActualWidth"].Value;
                Series = xmlsubsub.Attributes["Series"] == null ? "" : xmlsubsub.Attributes["Series"].Value;

                CabinetNo = xmlsubsub.Attributes["CabinetNo"] == null ? "" : xmlsubsub.Attributes["CabinetNo"].Value;
                CabinetPanelNo = xmlsubsub.Attributes["CabinetPanelNo"] == null ? "" : xmlsubsub.Attributes["CabinetPanelNo"].Value;

                for (int xmlnum = 0; xmlnum < xmlsubsub.ChildNodes.Count; xmlnum++)   //20180525
                {
                    foreach (XmlNode xnedge in xmlsubsub.ChildNodes[xmlnum].ChildNodes)
                    {
                        if (xnedge.Name == "Edge")
                        {
                            Edge edge = new Edge();
                            edge.LoadFromXmlNode(xnedge);
                            Edgelist.Add(edge);
                        }
                    }

                    foreach (XmlNode xnmachine in xmlsubsub.ChildNodes[xmlnum].ChildNodes)
                    {
                        if (xnmachine.Name == "Machining")
                        {
                            Machining machining = new Machining();
                            machining.LoadFromXmlNode(xnmachine);
                            if (machining.Type == "1" || machining.Type == "2" || machining.Type == "3")  //20180703  发现有除了1、2、3类型外的数字 则过滤
                                Machininglist.Add(machining);
                        }
                    }
                }

            }
            catch
            {
                throw new NotImplementedException();
            }

        }

        private void LoadFromXmlsubsubChildNode(XmlNode xmlsubsubsub, Cabinet parentcabinet) // 20180712
        {
            try
            {
                cabinet = parentcabinet;
                MulPanels = false;
                //if (xmlsubsubsub.Attributes["CabinetNo"] != null)
                //{
                for (int xmlsubnum = 0; xmlsubnum < xmlsubsubsub.ChildNodes.Count; xmlsubnum++)//20180627
                {
                    if (xmlsubsubsub.ChildNodes[xmlsubnum].Name == "Panels")
                    {
                        #region 20180914 若有抽屉，抽屉数量的写入
                        if (xmlsubsubsub.Attributes["SubType"] != null && xmlsubsubsub.Attributes["SubType"].Value == "drawer")
                        {
                            string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                            IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                            if (inifile.ExistINIFile())
                            {
                                int lastdrawernum = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                inifile.IniWriteValue("DrawerNum", "draw", (lastdrawernum + 1).ToString());
                            }
                            else
                            {
                                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                return;
                            }
                        }
                        #endregion

                        foreach (XmlNode xmlsub in xmlsubsubsub.ChildNodes[xmlsubnum].ChildNodes)
                        {
                            if (xmlsub.Name == "Panel")
                            {
                                Panel childpanel = new Panel();
                                childpanel.LoadFromXmlsubsubsubChildNode(xmlsub, parentcabinet);

                                if ((Math.Abs(double.Parse(childpanel.Thickness) - 5) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 8) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 15) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 18) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 22) < 0.01) ||  // 说有22厚的门板
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 25) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 35) < 0.01) ||
                                    (Math.Abs(double.Parse(childpanel.Thickness) - 50) < 0.01) ||
                                    childpanel.CabinetType.Contains("door") || childpanel.SubType.Contains("door"))
                                    childpanel.thicknessflag = true;

                                if (!childpanel.SubType.Contains("hardware") && !childpanel.CabinetType.Contains("hardware")) //20180714
                                {
                                    #region 20180914 若有抽屉,则抽屉板件的读取标识
                                    if (xmlsubsubsub.Attributes["SubType"] != null && xmlsubsubsub.Attributes["SubType"].Value == "drawer")
                                    {
                                        string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                        IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                        if (inifile.ExistINIFile())
                                        {
                                            childpanel.drawer = Convert.ToInt32(inifile.IniReadValue("DrawerNum", "draw"));
                                        }
                                        else
                                        {
                                            MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                            return;
                                        }
                                    }
                                    #endregion

                                    if (!childpanel.MulPanels && childpanel.thicknessflag)
                                    {
                                        string inipath = Path.Combine(Environment.CurrentDirectory, "OrderNo.ini");
                                        IniFiles inifile = new IniFiles(Path.Combine(Environment.CurrentDirectory, "OrderNo.ini"));
                                        if (inifile.ExistINIFile())
                                        {
                                            childpanel.SMAEX = inifile.IniReadValue("SAMEX", "ps");

                                        }
                                        else
                                        {
                                            MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                                            return;
                                        }

                                        if (childpanel.SMAEX.Equals("1") && childpanel.cabinet.CabinetNo != string.Empty) //20181228
                                            childpanel.Name = childpanel.cabinet.CabinetNo + "-" + childpanel.Name;

                                        cabinet.Panellist.Add(childpanel);
                                        //    if (childpanel.CabinetNo == "" && double.Parse(xmlsub.Attributes["Thickness"].Value) < 50.1)      //  子节点中 也是需要判断是否有组件20180410   //20180418
                                        //    cabinet.Panellist.Add(childpanel);
                                        //else if (childpanel.Category.Contains("Door") || childpanel.Name.Contains("抽面"))  //发现门 应该从属性里判断 宋新刚 20180622
                                        //    cabinet.Panellist.Add(childpanel);
                                    }
                                }
                            }
                        }
                        MulPanels = true;
                    }
                }
                // }

                PositionNumber = xmlsubsubsub.Attributes["PositionNumber"] == null ? "" : xmlsubsubsub.Attributes["PositionNumber"].Value;
                IsProduce = xmlsubsubsub.Attributes["IsProduce"] == null ? "" : xmlsubsubsub.Attributes["IsProduce"].Value;
                Thickness = xmlsubsubsub.Attributes["Thickness"] == null ? "" : xmlsubsubsub.Attributes["Thickness"].Value;
                CraftMark = xmlsubsubsub.Attributes["CraftMark"] == null ? "" : xmlsubsubsub.Attributes["CraftMark"].Value;
                SubType = xmlsubsubsub.Attributes["SubType"] == null ? "" : xmlsubsubsub.Attributes["SubType"].Value;
                Length = xmlsubsubsub.Attributes["Length"] == null ? "" : xmlsubsubsub.Attributes["Length"].Value;
                Width = xmlsubsubsub.Attributes["Width"] == null ? "" : xmlsubsubsub.Attributes["Width"].Value;
                ID = xmlsubsubsub.Attributes["ID"] == null ? "" : xmlsubsubsub.Attributes["ID"].Value;
                Name = xmlsubsubsub.Attributes["Name"] == null ? "" : xmlsubsubsub.Attributes["Name"].Value;
                Material = xmlsubsubsub.Attributes["Material"] == null ? "" : xmlsubsubsub.Attributes["Material"].Value;
                MaterialId = xmlsubsubsub.Attributes["MaterialId"] == null ? "" : xmlsubsubsub.Attributes["MaterialId"].Value;
                BaseMaterialCategoryId = xmlsubsubsub.Attributes["BaseMaterialCategoryId"] == null ? "" : xmlsubsubsub.Attributes["BaseMaterialCategoryId"].Value;
                MaterialCategoryId = xmlsubsubsub.Attributes["MaterialCategoryId"] == null ? "" : xmlsubsubsub.Attributes["MaterialCategoryId"].Value;
                Model = xmlsubsubsub.Attributes["Model"] == null ? "" : xmlsubsubsub.Attributes["Model"].Value;
                CabinetType = xmlsubsubsub.Attributes["CabinetType"] == null ? "" : xmlsubsubsub.Attributes["CabinetType"].Value; //20180714
                Type = xmlsubsubsub.Attributes["Type"] == null ? "" : xmlsubsubsub.Attributes["Type"].Value;
                edgeMaterial = xmlsubsubsub.Attributes["edgeMaterial"] == null ? "" : xmlsubsubsub.Attributes["edgeMaterial"].Value;
                StandardCategory = xmlsubsubsub.Attributes["StandardCategory"] == null ? "" : xmlsubsubsub.Attributes["StandardCategory"].Value;
                IsAccurate = xmlsubsubsub.Attributes["IsAccurate"] == null ? "" : xmlsubsubsub.Attributes["IsAccurate"].Value;
                MachiningPoint = xmlsubsubsub.Attributes["MachiningPoint"] == null ? "" : xmlsubsubsub.Attributes["MachiningPoint"].Value;
                Grain = xmlsubsubsub.Attributes["Grain"] == null ? "" : xmlsubsubsub.Attributes["Grain"].Value;
                ProdutionNo = xmlsubsubsub.Attributes["ProdutionNo"] == null ? "" : xmlsubsubsub.Attributes["ProdutionNo"].Value;
                ProductionName = xmlsubsubsub.Attributes["ProductionName"] == null ? "" : xmlsubsubsub.Attributes["ProductionName"].Value;

                Face5ID = "P" + ID.Substring((ID.Length - 3), 3) + "X";  //20180409
                Face6ID = "P" + ID.Substring((ID.Length - 3), 3) + "Y";  //20180409

                clerk = xmlsubsubsub.Attributes["clerk"] == null ? "" : xmlsubsubsub.Attributes["clerk"].Value;
                PkgNo = xmlsubsubsub.Attributes["PkgNo"] == null ? "" : xmlsubsubsub.Attributes["PkgNo"].Value;
                BasicMaterial = xmlsubsubsub.Attributes["BasicMaterial"] == null ? "" : xmlsubsubsub.Attributes["BasicMaterial"].Value;
                PartNumber = xmlsubsubsub.Attributes["PartNumber"] == null ? "" : xmlsubsubsub.Attributes["PartNumber"].Value;

                if (PartNumber.ToUpper().Contains("Y"))   //20180417
                {
                    Name = Name + "-Y";
                }

                DoorDirection = xmlsubsubsub.Attributes["DoorDirection"] == null ? "" : xmlsubsubsub.Attributes["DoorDirection"].Value;
                Category = xmlsubsubsub.Attributes["Category"] == null ? "" : xmlsubsubsub.Attributes["Category"].Value;
                thickLength = xmlsubsubsub.Attributes["thickLength"] == null ? "" : xmlsubsubsub.Attributes["thickLength"].Value;
                thinLength = xmlsubsubsub.Attributes["thinLength"] == null ? "" : xmlsubsubsub.Attributes["thinLength"].Value;
                customLength = xmlsubsubsub.Attributes["customLength"] == null ? "" : xmlsubsubsub.Attributes["customLength"].Value;
                slotDis = xmlsubsubsub.Attributes["slotDis"] == null ? "" : xmlsubsubsub.Attributes["slotDis"].Value;
                slotFace = xmlsubsubsub.Attributes["slotFace"] == null ? "" : xmlsubsubsub.Attributes["slotFace"].Value;
                HasHorizontalHole = xmlsubsubsub.Attributes["HasHorizontalHole"] == null ? "" : xmlsubsubsub.Attributes["HasHorizontalHole"].Value;
                ActualLength = xmlsubsubsub.Attributes["ActualLength"] == null ? "" : xmlsubsubsub.Attributes["ActualLength"].Value;
                ActualWidth = xmlsubsubsub.Attributes["ActualWidth"] == null ? "" : xmlsubsubsub.Attributes["ActualWidth"].Value;
                Series = xmlsubsubsub.Attributes["Series"] == null ? "" : xmlsubsubsub.Attributes["Series"].Value;

                CabinetNo = xmlsubsubsub.Attributes["CabinetNo"] == null ? "" : xmlsubsubsub.Attributes["CabinetNo"].Value;
                CabinetPanelNo = xmlsubsubsub.Attributes["CabinetPanelNo"] == null ? "" : xmlsubsubsub.Attributes["CabinetPanelNo"].Value;

                for (int xmlnum = 0; xmlnum < xmlsubsubsub.ChildNodes.Count; xmlnum++)   //20180525
                {
                    foreach (XmlNode xnedge in xmlsubsubsub.ChildNodes[xmlnum].ChildNodes)
                    {
                        if (xnedge.Name == "Edge")
                        {
                            Edge edge = new Edge();
                            edge.LoadFromXmlNode(xnedge);
                            Edgelist.Add(edge);
                        }
                    }

                    foreach (XmlNode xnmachine in xmlsubsubsub.ChildNodes[xmlnum].ChildNodes)
                    {
                        if (xnmachine.Name == "Machining")
                        {
                            Machining machining = new Machining();
                            machining.LoadFromXmlNode(xnmachine);
                            if (machining.Type == "1" || machining.Type == "2" || machining.Type == "3")  //20180703  发现有除了1、2、3类型外的数字 则过滤
                                Machininglist.Add(machining);
                        }
                    }
                }

            }
            catch
            {
                throw new NotImplementedException();
            }
        }

        private void LoadFromXmlsubsubsubChildNode(XmlNode xmlsubsubsubsub, Cabinet parentcabinet) //20180712
        {
            try
            {
                cabinet = parentcabinet;
                PositionNumber = xmlsubsubsubsub.Attributes["PositionNumber"] == null ? "" : xmlsubsubsubsub.Attributes["PositionNumber"].Value;
                IsProduce = xmlsubsubsubsub.Attributes["IsProduce"] == null ? "" : xmlsubsubsubsub.Attributes["IsProduce"].Value;
                Thickness = xmlsubsubsubsub.Attributes["Thickness"] == null ? "" : xmlsubsubsubsub.Attributes["Thickness"].Value;
                CraftMark = xmlsubsubsubsub.Attributes["CraftMark"] == null ? "" : xmlsubsubsubsub.Attributes["CraftMark"].Value;
                SubType = xmlsubsubsubsub.Attributes["SubType"] == null ? "" : xmlsubsubsubsub.Attributes["SubType"].Value;
                Length = xmlsubsubsubsub.Attributes["Length"] == null ? "" : xmlsubsubsubsub.Attributes["Length"].Value;
                Width = xmlsubsubsubsub.Attributes["Width"] == null ? "" : xmlsubsubsubsub.Attributes["Width"].Value;
                ID = xmlsubsubsubsub.Attributes["ID"] == null ? "" : xmlsubsubsubsub.Attributes["ID"].Value;
                Name = xmlsubsubsubsub.Attributes["Name"] == null ? "" : xmlsubsubsubsub.Attributes["Name"].Value;
                Material = xmlsubsubsubsub.Attributes["Material"] == null ? "" : xmlsubsubsubsub.Attributes["Material"].Value;
                MaterialId = xmlsubsubsubsub.Attributes["MaterialId"] == null ? "" : xmlsubsubsubsub.Attributes["MaterialId"].Value;
                BaseMaterialCategoryId = xmlsubsubsubsub.Attributes["BaseMaterialCategoryId"] == null ? "" : xmlsubsubsubsub.Attributes["BaseMaterialCategoryId"].Value;
                MaterialCategoryId = xmlsubsubsubsub.Attributes["MaterialCategoryId"] == null ? "" : xmlsubsubsubsub.Attributes["MaterialCategoryId"].Value;
                Model = xmlsubsubsubsub.Attributes["Model"] == null ? "" : xmlsubsubsubsub.Attributes["Model"].Value;
                CabinetType = xmlsubsubsubsub.Attributes["CabinetType"] == null ? "" : xmlsubsubsubsub.Attributes["CabinetType"].Value; //20180714
                Type = xmlsubsubsubsub.Attributes["Type"] == null ? "" : xmlsubsubsubsub.Attributes["Type"].Value;
                edgeMaterial = xmlsubsubsubsub.Attributes["edgeMaterial"] == null ? "" : xmlsubsubsubsub.Attributes["edgeMaterial"].Value;
                StandardCategory = xmlsubsubsubsub.Attributes["StandardCategory"] == null ? "" : xmlsubsubsubsub.Attributes["StandardCategory"].Value;
                IsAccurate = xmlsubsubsubsub.Attributes["IsAccurate"] == null ? "" : xmlsubsubsubsub.Attributes["IsAccurate"].Value;
                MachiningPoint = xmlsubsubsubsub.Attributes["MachiningPoint"] == null ? "" : xmlsubsubsubsub.Attributes["MachiningPoint"].Value;
                Grain = xmlsubsubsubsub.Attributes["Grain"] == null ? "" : xmlsubsubsubsub.Attributes["Grain"].Value;
                ProdutionNo = xmlsubsubsubsub.Attributes["ProdutionNo"] == null ? "" : xmlsubsubsubsub.Attributes["ProdutionNo"].Value;
                ProductionName = xmlsubsubsubsub.Attributes["ProductionName"] == null ? "" : xmlsubsubsubsub.Attributes["ProductionName"].Value;

                Face5ID = "P" + ID.Substring((ID.Length - 3), 3) + "X";  //20180409
                Face6ID = "P" + ID.Substring((ID.Length - 3), 3) + "Y";  //20180409

                clerk = xmlsubsubsubsub.Attributes["clerk"] == null ? "" : xmlsubsubsubsub.Attributes["clerk"].Value;
                PkgNo = xmlsubsubsubsub.Attributes["PkgNo"] == null ? "" : xmlsubsubsubsub.Attributes["PkgNo"].Value;
                BasicMaterial = xmlsubsubsubsub.Attributes["BasicMaterial"] == null ? "" : xmlsubsubsubsub.Attributes["BasicMaterial"].Value;
                PartNumber = xmlsubsubsubsub.Attributes["PartNumber"] == null ? "" : xmlsubsubsubsub.Attributes["PartNumber"].Value;

                if (PartNumber.ToUpper().Contains("Y"))   //20180417
                {
                    Name = Name + "-Y";
                }

                DoorDirection = xmlsubsubsubsub.Attributes["DoorDirection"] == null ? "" : xmlsubsubsubsub.Attributes["DoorDirection"].Value;
                Category = xmlsubsubsubsub.Attributes["Category"] == null ? "" : xmlsubsubsubsub.Attributes["Category"].Value;
                thickLength = xmlsubsubsubsub.Attributes["thickLength"] == null ? "" : xmlsubsubsubsub.Attributes["thickLength"].Value;
                thinLength = xmlsubsubsubsub.Attributes["thinLength"] == null ? "" : xmlsubsubsubsub.Attributes["thinLength"].Value;
                customLength = xmlsubsubsubsub.Attributes["customLength"] == null ? "" : xmlsubsubsubsub.Attributes["customLength"].Value;
                slotDis = xmlsubsubsubsub.Attributes["slotDis"] == null ? "" : xmlsubsubsubsub.Attributes["slotDis"].Value;
                slotFace = xmlsubsubsubsub.Attributes["slotFace"] == null ? "" : xmlsubsubsubsub.Attributes["slotFace"].Value;
                HasHorizontalHole = xmlsubsubsubsub.Attributes["HasHorizontalHole"] == null ? "" : xmlsubsubsubsub.Attributes["HasHorizontalHole"].Value;
                ActualLength = xmlsubsubsubsub.Attributes["ActualLength"] == null ? "" : xmlsubsubsubsub.Attributes["ActualLength"].Value;
                ActualWidth = xmlsubsubsubsub.Attributes["ActualWidth"] == null ? "" : xmlsubsubsubsub.Attributes["ActualWidth"].Value;
                Series = xmlsubsubsubsub.Attributes["Series"] == null ? "" : xmlsubsubsubsub.Attributes["Series"].Value;

                CabinetNo = xmlsubsubsubsub.Attributes["CabinetNo"] == null ? "" : xmlsubsubsubsub.Attributes["CabinetNo"].Value;
                CabinetPanelNo = xmlsubsubsubsub.Attributes["CabinetPanelNo"] == null ? "" : xmlsubsubsubsub.Attributes["CabinetPanelNo"].Value;

                for (int xmlnum = 0; xmlnum < xmlsubsubsubsub.ChildNodes.Count; xmlnum++)   //20180525
                {
                    foreach (XmlNode xnedge in xmlsubsubsubsub.ChildNodes[xmlnum].ChildNodes)
                    {
                        if (xnedge.Name == "Edge")
                        {
                            Edge edge = new Edge();
                            edge.LoadFromXmlNode(xnedge);
                            Edgelist.Add(edge);
                        }
                    }

                    foreach (XmlNode xnmachine in xmlsubsubsubsub.ChildNodes[xmlnum].ChildNodes)
                    {
                        if (xnmachine.Name == "Machining")
                        {
                            Machining machining = new Machining();
                            machining.LoadFromXmlNode(xnmachine);
                            if (machining.Type == "1" || machining.Type == "2" || machining.Type == "3")  //20180703  发现有除了1、2、3类型外的数字 则过滤
                                Machininglist.Add(machining);
                        }
                    }
                }
            }
            catch
            {
                throw new NotImplementedException();
            }

        }
    }

    public class Edge
    {
        public string Face { get; set; }
        public string LindID { get; set; }
        public string Thickness { get; set; }
        public string EdgeType { get; set; }
        public string Length { get; set; }
        public string Pre_Milling { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
        public string CentralAngle { get; set; }

        internal void LoadFromXmlNode(XmlNode xnedge)
        {
            try
            {
                Face = xnedge.Attributes["Face"].Value;

                if (xnedge.Attributes["LindID"] != null)   //20180328
                    LindID = xnedge.Attributes["LindID"].Value;
                else if (xnedge.Attributes["LineID"] != null)
                    LindID = xnedge.Attributes["LineID"].Value;

                Thickness = xnedge.Attributes["Thickness"] == null ? "" : xnedge.Attributes["Thickness"].Value;
                EdgeType = xnedge.Attributes["EdgeType"] == null ? "" : xnedge.Attributes["EdgeType"].Value;
                Length = xnedge.Attributes["Length"] == null ? "" : xnedge.Attributes["Length"].Value;   //20180329
                Pre_Milling = xnedge.Attributes["Pre_Milling"] == null ? "" : xnedge.Attributes["Pre_Milling"].Value;
                X = xnedge.Attributes["X"] == null ? "" : xnedge.Attributes["X"].Value;
                Y = xnedge.Attributes["Y"] == null ? "" : xnedge.Attributes["Y"].Value;
                CentralAngle = xnedge.Attributes["CentralAngle"] == null ? "" : xnedge.Attributes["CentralAngle"].Value;
            }
            catch
            {
                throw new NotImplementedException();
            }

        }
    }

    public class Machining
    {
        public string ID { get; set; }
        public string IsGenCode { get; set; }
        public string Type { get; set; }
        public string Face { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
        public string Z { get; set; }
        public string HoleType { get; set; }
        public string Diameter { get; set; }
        public string Depth { get; set; }
        public string Width { get; set; }
        public string ToolName { get; set; }
        public string ToolOffset { get; set; }
        public string EdgeABThickness { get; set; }
        public string GrooveType { get; set; }

        public List<Line> Linelist = new List<Line>();


        internal void LoadFromXmlNode(XmlNode xnedge)
        {
            try
            {
                ID = xnedge.Attributes["ID"] == null ? "" : xnedge.Attributes["ID"].Value;
                IsGenCode = xnedge.Attributes["IsGenCode"] == null ? "" : xnedge.Attributes["IsGenCode"].Value;
                Type = xnedge.Attributes["Type"] == null ? "" : xnedge.Attributes["Type"].Value;
                Face = xnedge.Attributes["Face"] == null ? "" : xnedge.Attributes["Face"].Value;
                X = xnedge.Attributes["X"] == null ? "" : xnedge.Attributes["X"].Value;
                Y = xnedge.Attributes["Y"] == null ? "" : xnedge.Attributes["Y"].Value;

                Depth = xnedge.Attributes["Depth"] == null ? "" : xnedge.Attributes["Depth"].Value;

                if (Type == "3")
                {
                    if (xnedge.Attributes["Width"] != null)
                    {
                        Width = xnedge.Attributes["Width"].Value;
                    }

                    if (xnedge.Attributes["EdgeABThickness"] != null || xnedge.Attributes["GrooveType"] != null)
                    {
                        EdgeABThickness = xnedge.Attributes["EdgeABThickness"].Value;
                        GrooveType = xnedge.Attributes["GrooveType"].Value;
                    }

                    ToolName = xnedge.Attributes["ToolName"] == null ? "" : xnedge.Attributes["ToolName"].Value;  // 20180703
                    ToolOffset = xnedge.Attributes["ToolOffset"] == null ? "" : xnedge.Attributes["ToolOffset"].Value; // 20180703
                    Diameter = "";
                    Z = "";
                    foreach (XmlNode xnline in xnedge.ChildNodes[0].ChildNodes)
                    {
                        if (xnline.Name == "Line")
                        {
                            Line line = new Line();
                            line.LoadFromXmlNode(xnline);
                            Linelist.Add(line);
                        }
                    }
                }
                else
                {
                    ToolName = "";
                    ToolOffset = "";
                    Diameter = xnedge.Attributes["Diameter"] == null ? "" : xnedge.Attributes["Diameter"].Value;//20180703
                    Z = xnedge.Attributes["Z"].Value == null ? "" : xnedge.Attributes["Z"].Value;//20180703
                    HoleType = xnedge.Attributes["HoleType"] == null ? "" : xnedge.Attributes["HoleType"].Value;//20180703
                }

            }
            catch
            {

                throw new NotImplementedException();
            }


        }
    }

    public class Metal
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string PartNumber { get; set; }
        public string length { get; set; }
        public string width { get; set; }
        public string height { get; set; }
        public string Num { get; set; }
        public string Style { get; set; }
        public string Unit { get; set; }


        internal void LoadFromXmlNode(XmlNode xnpanel)
        {
            try
            {
                Id = xnpanel.Attributes["Id"] == null ? "" : xnpanel.Attributes["Id"].Value;
                Name = xnpanel.Attributes["Name"] == null ? "" : xnpanel.Attributes["Name"].Value;
                PartNumber = xnpanel.Attributes["PartNumber"] == null ? "" : xnpanel.Attributes["PartNumber"].Value;
                length = xnpanel.Attributes["length"] == null ? "" : xnpanel.Attributes["length"].Value;
                width = xnpanel.Attributes["width"] == null ? "" : xnpanel.Attributes["width"].Value;
                height = xnpanel.Attributes["height"] == null ? "" : xnpanel.Attributes["height"].Value;
                Num = xnpanel.Attributes["Num"] == null ? "" : xnpanel.Attributes["Num"].Value;
                Style = xnpanel.Attributes["Style"] == null ? "" : xnpanel.Attributes["Style"].Value;
                Unit = xnpanel.Attributes["Unit"] == null ? "" : xnpanel.Attributes["Unit"].Value;
            }
            catch
            {
                throw new NotImplementedException();
            }


        }
    }

    public class Line
    {
        public string LineID { get; set; }
        public string EndX { get; set; }
        public string EndY { get; set; }
        public string Angle { get; set; }
        //public string Thickness { get; set; }   20180323


        internal void LoadFromXmlNode(XmlNode xnline)
        {
            try
            {
                LineID = xnline.Attributes["LineID"] == null ? "" : xnline.Attributes["LineID"].Value;
                EndX = xnline.Attributes["EndX"] == null ? "" : xnline.Attributes["EndX"].Value;
                EndY = xnline.Attributes["EndY"] == null ? "" : xnline.Attributes["EndY"].Value;
                Angle = xnline.Attributes["Angle"] == null ? "" : xnline.Attributes["Angle"].Value;
                //Thickness = xnline.Attributes["Thickness"].Value;   20180323
            }
            catch
            {

                throw new NotImplementedException();
            }


        }
    }
}