using SWJXMLToCSV.GeneratePictures;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SWJXMLToCSV
{
    public class ConvertCore
    {
        private static string csvpath = string.Empty;

        /// <summary>
        /// 删除文件夹及子文件内文件
        /// </summary>
        /// <param name="str"></param>
        public static void DeleteFiles(string str)
        {
            DirectoryInfo fatherFolder = new DirectoryInfo(str);
            //删除当前文件夹内文件
            FileInfo[] files = fatherFolder.GetFiles();
            foreach (FileInfo file in files)
            {
                //string fileName = file.FullName.Substring((file.FullName.LastIndexOf("\\") + 1), file.FullName.Length - file.FullName.LastIndexOf("\\") - 1);
                string fileName = file.Name;
                try
                {
                    if (!fileName.Equals("index.dat"))
                    {
                        File.Delete(file.FullName);
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            //递归删除子文件夹内文件
            foreach (DirectoryInfo childFolder in fatherFolder.GetDirectories())
            {
                DeleteFiles(childFolder.FullName);
            }


        }

        public static List<string> Output(List<Cabinet> list, string orderNo=null) {
            //if (targetPath == null)
            //{
            //    string path = "/xml/" + string.Format("{0}/{1}/", (DateTime.Now.Year - 2000), DateTime.Now.ToString("MMdd"));
            //    if(System.Web.HttpContext.Current == null)
            //        csvpath = AppDomain.CurrentDomain.BaseDirectory+ Guid.NewGuid().ToString("N");
            //    else
            //        csvpath = System.Web.HttpContext.Current.Server.MapPath("~" + path) + Guid.NewGuid().ToString("N");
            //}
            //else
            //    csvpath = targetPath;
            //if (!Directory.Exists(csvpath))
            //    Directory.CreateDirectory(csvpath);
            List<string> fileNameList = new List<string>();
            ArrayList face5list = new ArrayList();
            ArrayList face6list = new ArrayList();
            string basePath = string.Empty;
            //以 /cdwj/OrderNo/CabinetNo 作为CSV存放路径
            if (!string.IsNullOrEmpty(orderNo))
            {
                if (System.Web.HttpContext.Current == null)
                    basePath = AppDomain.CurrentDomain.BaseDirectory + "cdwj\\" + orderNo + "\\";
                else
                    basePath = System.Web.HttpContext.Current.Server.MapPath("~/cdwj/" + orderNo + "/");
            }
            if (Directory.Exists(basePath)) {
                try
                {
                   Directory.Delete(basePath, true); //20180829 作用C#删除文件夹的方法
                   //DeleteFiles(basePath);
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }
            foreach (var cabinet in list)
            {
                csvpath = basePath + cabinet.CabinetNo;
                try
                {
                    Directory.CreateDirectory(csvpath);
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message); ;
                }
                
                foreach (var panel in cabinet.Panellist)
                {
                    string Cao_5 = "0";  //20180807 初始化没有槽
                    string Cao_6 = "0";

                    List<fourpoint> point4 = new List<fourpoint>();
                    face5list = new ArrayList();
                    face6list = new ArrayList();

                    #region borderseq
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

                    borderseq.Description = panel.Name;
                    borderseq.PartQty = "1";
                    borderseq.CutPartWidth = panel.ActualWidth;
                    borderseq.CutPartLength = panel.ActualLength;
                    borderseq.MaterialName = panel.Thickness + "mm" + panel.Material + panel.BasicMaterial;
                    borderseq.MaterialCode = panel.MaterialId;
                    if (panel.Edgelist.Count == 4)  // 20180525
                    {
                        borderseq.Edgeband1 = panel.Edgelist[0].Thickness + "mm" + panel.Material + "封边条";               //应该要
                        borderseq.Edgeband2 = panel.Edgelist[1].Thickness + "mm" + panel.Material + "封边条";                                   //应该要
                        borderseq.Edgeband3 = panel.Edgelist[2].Thickness + "mm" + panel.Material + "封边条";                                     //应该要
                        borderseq.Edgeband4 = panel.Edgelist[3].Thickness + "mm" + panel.Material + "封边条";                                       //应该要

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
                    borderseq.SawsCount = "0"; //记录面5 槽类型等  20180807
                    borderseq.NestRoutesCount = "";

                    point4.Add(new fourpoint(0, 0));
                    point4.Add(new fourpoint(double.Parse(borderseq.PanelLength), 0));
                    point4.Add(new fourpoint(double.Parse(borderseq.PanelLength), double.Parse(borderseq.PanelWidth)));
                    point4.Add(new fourpoint(0, double.Parse(borderseq.PanelWidth)));
                    #endregion

                    #region Machininglist
                    for (int i = 0; i < panel.Machininglist.Count; i++)
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

                            if (panel.Machininglist[i].Face != "0")  //20180816 发现水平孔有面为0的情况，将之面为0的情况过滤
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

                            if (vdrillseq.VDrillDiameter != "")//20180828 发现垂直孔的直径出现为空的情况，将此情况过滤
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
                                                        routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y - 3 / 2);

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

                                                        routesecond = RouteProcess(panel.Machininglist[i], height, min_x, min_y + 3 / 2);
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
                                                                routesecond = RouteProcess(panel.Machininglist[i], height, leadin_x - 3 / 2, max_y);   //20180420

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
                                            //MessageBox.Show("槽宽为:  " + height.ToString() + "不在有其 3 - 8.5mm的范围之内");
                                            //return;
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
                                                    routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y - 8.5 / 2);
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
                                                    routesecond = RouteProcess(panel.Machininglist[i], height, min_x, min_y + 8.5 / 2);
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
                                                            routesecond = RouteProcess(panel.Machininglist[i], height, leadin_x - 3 / 2, max_y);   //20180420

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
                                                routefirst = RouteProcess(panel.Machininglist[i], height, min_x, max_y + 8.5 / 2);

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
                                            {//这边的情况有如下几点：
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
                                                        routesecond = RouteProcess(panel.Machininglist[i], height, leadin_x - 10 / 2, max_y);   //20180420  //20180818 将这里的8.5改成10

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
                                        //MessageBox.Show("槽宽为:  " + height.ToString() + "没有合适的刀具进行加工!");
                                        //return;
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
                                        useD3after129 = false;  //20180620 意大利铝框灰玻门铰 不需要D3刀绕圈 51改到55是因为要多切出去4mm

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
                                                        routefirst = RouteProcess(panel.Machininglist[i], width, min_x + 3 / 2, min_y);

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

                                                        routesecond = RouteProcess(panel.Machininglist[i], width, max_x - 3 / 2, min_y);
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
                                                                routesecond = RouteProcess(panel.Machininglist[i], width, min_x, leadin_y - 3 / 2);

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
                                            //MessageBox.Show("槽宽为:  " + width.ToString() + "不在有其 3 - 8.5mm的范围之内");
                                            //return;
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
                                                    routesecond = RouteProcess(panel.Machininglist[i], width, max_x - 8.5 / 2, min_y);
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
                                                            routesecond = RouteProcess(panel.Machininglist[i], width, min_x, leadin_y - 3 / 2);

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
                                                routefirst = RouteProcess(panel.Machininglist[i], width, min_x - 8.5 / 2, min_y);

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
                                            {//这边的情况有如下几点：
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
                                                        routesecond = RouteProcess(panel.Machininglist[i], width, min_x, (leadin_y - 10 / 2)); //20180818 将这里的8.5改成10

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
                                        //MessageBox.Show("槽宽为:  " + width.ToString() + "没有合适的刀具进行加工!");
                                        //return;
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
                                else if (panel.Machininglist[i].ToolName == "开槽刀")// 厨柜因为背板只有5mm，增加了6.35 20180419
                                {                                                                       
                                    if (Math.Abs(double.Parse(panel.Machininglist[i].Width) - 6) < 0.1)
                                    {
                                        routesetmillseq.RouteDiameter = "6.35";
                                        routesetmillseq.RouteToolName = "131";
                                    }
                                    else
                                    {
                                        routesetmillseq.RouteDiameter = "8.5";
                                        routesetmillseq.RouteToolName = "129";
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
                                routesetmillseq.RouteEndTangentY = "";

                                #region 判断第一个点是不是工件的四个顶点。如果是四个顶点的其中一个，则删点。如果不是，则暂时跳过步骤

                                List<fourpoint> pointxy1 = new List<fourpoint>();

                                List<fourpoint> partnotfourpoints = new List<fourpoint>();  //建个容器用来保存不是顶点的点.如果第一个点就不是顶点中的点，则不需要增加，因为一段多段线的最后一个点肯定也是这个点

                                bool firstcomeout = true;
                                bool firstcomeout1 = true;
                                bool yesornofourpoint = false;
                                bool comeoutallprofile = false;
                                bool needallprofile = false;
                                if (Isforpoint(point4, (new fourpoint(double.Parse(routesetmillseq.RouteX), double.Parse(routesetmillseq.RouteY)))))
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

                                for (int j = 0; j < panel.Machininglist[i].Linelist.Count; j++)
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
                                    else if (panel.Machininglist[i].ToolName == "开槽刀")// 厨柜因为背板只有5mm，增加了6.35 20180419
                                    {
                                        
                                        if (Math.Abs(double.Parse(panel.Machininglist[i].Width) - 6) < 0.1)
                                        {
                                            routeseq.RouteDiameter = "6.35";
                                            routeseq.RouteToolName = "131";
                                        }
                                        else
                                        {
                                            routeseq.RouteDiameter = "8.5";
                                            routeseq.RouteToolName = "129";
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
                                        routeseq.RouteBulge = ((1 - Math.Cos(numflag2)) / Math.Sin(numflag2) * -1).ToString("F5");
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
                                                    throw new Exception("开始点超出统计范围(一)，请与宋新刚 18913812043联系！");
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
                                                    throw new Exception("开始点超出统计范围(二)，请与宋新刚 18913812043联系！");
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
                    #endregion

                    // 因为有些是需要在水平孔、垂直孔判读，故在最后再增加板件信息
                    face5list.Add(borderseq.OutPutCsvString());

                    borderseq.FileName = "";   //至反面加工码的时候，不需要正面加工码的名字
                    face6list.Add(borderseq.OutPutCsvString());

                    // 因为有些是需要在水平孔、垂直孔判读，故在最后再增加板件信息

                    //csvname = csvname = panel.ID.Substring(0, 6) + panel.ID.Substring(7, 3) + "X";
                    //string csvname = borderseq.FileName;
                    string csvname = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "X";
                    OutFace5Face6Csv(face5list, csvname, fileNameList, double.Parse(borderseq.PanelThickness));   // 20180418
                    //csvname = csvname = panel.ID.Substring(0, 6) + panel.ID.Substring(7, 3) + "Y";
                    //csvname = borderseq.Face6FileName;
                    csvname = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "Y";
                    OutFace5Face6Csv(face6list, csvname, fileNameList, double.Parse(borderseq.PanelThickness));   // 20180418
                }
            }

            foreach (var cabinet in list)
            {
                foreach (var panel in cabinet.Panellist)
                {
                    #region         
                    string F5FileName = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "X";
                    string F6FileName = "P" + panel.ID.Substring((panel.ID.Length - 3), 3) + "Y";

                    string F5csv = Path.Combine(Path.Combine(Path.GetDirectoryName(csvpath),cabinet.CabinetNo), F5FileName + ".csv");
                    string F6csv = Path.Combine(Path.Combine(Path.GetDirectoryName(csvpath), cabinet.CabinetNo), F6FileName + ".csv");
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
                            fileNameList.Remove(F6csv);
                            File.Delete(F6csv);

                            StreamWriter swF6_F5 = new StreamWriter(F5csv, false, Encoding.Default);
                            fileNameList.Add(F5csv);
                            foreach (string str in face6_face5)
                            {
                                swF6_F5.WriteLine(str);
                            }
                            swF6_F5.WriteLine("EndSequence,,,,");
                            swF6_F5.Flush();
                            swF6_F5.Close();

                        }
                    }
                    else
                    {
                        if (!File.Exists(F6csv))
                        {
                            //fileNameList.Remove(F6csv);
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
                                fileNameList.Remove(F6csv);
                                File.Delete(F6csv);
                                fileNameList.Remove(F5csv);  //20180821 删除码的时候，这里也要删掉
                                File.Delete(F5csv);

                                StreamWriter swF6_F5 = new StreamWriter(F5csv, false, Encoding.Default);
                                fileNameList.Add(F5csv);
                                foreach (string str in face6_face5)
                                {
                                    swF6_F5.WriteLine(str);
                                }
                                swF6_F5.WriteLine("EndSequence,,,,");
                                swF6_F5.Flush();
                                swF6_F5.Close();

                                //nest.F6FileName = "";
                            }

                        }

                    }
                    #endregion
                    if (panel.Category == "Door" && fileNameList.Contains(F5csv) && !panel.Name.Contains("榻榻米手动台面板") && !panel.Name.Contains("假门")) //20180713 不是门 但如果画图有问题
                    {

                        MvDrawPart.DrawPartBitmap(F5csv);
                        var imageUri = F5csv.Replace(".csv", ".jpg");
                        fileNameList.Add(imageUri);
                    }
                }

            }
            return fileNameList;
        }

        public static RouteSetMillSequenceEntity RouteProcess(Machining machining, double width, double pointx, double pointy)  //20180327
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
            routesetmillseq.RouteStartOffsetY = (pointy + float.Parse(routesetmillseq.RouteDiameter) / 2).ToString();
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
            routesetmillseq.RouteEndTangentY = "";

            return (routesetmillseq);
        }

        private static bool Isforpoint(List<fourpoint> point4, fourpoint fourpoint)
        {
            try
            {
                foreach (fourpoint FPT in point4)
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

        private static bool twopointisHorVline(List<fourpoint> pointxy1, List<fourpoint> pointxy2)
        {
            try
            {
                double px1 = 0;
                double py1 = 0;
                double px2 = 0;
                double py2 = 0;

                foreach (fourpoint fp in pointxy1)
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

        public static void OutFace5Face6Csv(ArrayList list, string csvname,List<string> fileNameList,double Thickness)
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
            for (int m = 0; m < orderRouteSequencemodify.Count; m++)
            {
                string line = orderRouteSequencemodify[m].ToString();
                RouteSetMillSequenceEntity subvalue = new RouteSetMillSequenceEntity(line);

                if (subvalue.RouteDiameter == "10" && Thickness - 8 > 0.1 && subvalue.RouteEndTangentY != "1")   //20180418 //20180726 增加减不减封边的识别
                {
                    if (line.StartsWith("RouteSetMillSequence"))
                    {
                        if (m != 0)
                        {
                            string line2 = orderRouteSequencemodify[m - 1].ToString();
                            RouteSetMillSequenceEntity subvalue2 = new RouteSetMillSequenceEntity(line2);
                            if (subvalue2.RouteDiameter == "10")  //上一段铣型也必须要是用10mm的刀具  宋新刚 20180326
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
                        //MessageBox.Show("异形减封边不在上述的范围之内.请将XML文件准备好，与宋新刚18913812043联系!");
                    }

                    string line1 = orderRouteSequencemodify[m - 1].ToString();
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

            fileNameList.Add(path);
        }


    }
}
