using WorkList;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CamFeatureDataHelper
{
    public class MVDataHelper
    {
        public MVDataHelper()
        { }

        public bool ReadDataFromCsv(string csvFilename, bool bface6)
        {
            try
            {
                StreamReader sr = new StreamReader(csvFilename, Encoding.Default);
                string strLine = sr.ReadLine();
                //strLine = sr.ReadToEnd();
                int RouteSetIndex = 0;
                int routeswqindex = 0;
                ArrayList Seqlist = new ArrayList();
                Seqlist.Clear();

                ArrayList outstrlist = new ArrayList();
                BorderSequenceEntity borderSeq = new BorderSequenceEntity(strLine);
                while (strLine != null)
                {
                    strLine = strLine.Replace("D8.5", "129").Replace("D10", "130");
                    ////处理每行的字符串
                    if (strLine.IndexOf("BorderSequence,") == 0)
                    {
                        //BorderSequenceEntity borderSeq = new BorderSequenceEntity(strLine);
                        borderSeq.NestRoutesCount = "0";
                        borderSeq.FileName = Path.GetFileNameWithoutExtension(csvFilename);
                        borderSeq.Edgeband1 = borderSeq.Edgeband1.Replace("MM", "mm");
                        borderSeq.Edgeband2 = borderSeq.Edgeband1.Replace("MM", "mm");
                        borderSeq.Edgeband3 = borderSeq.Edgeband1.Replace("MM", "mm");
                        borderSeq.Edgeband4 = borderSeq.Edgeband1.Replace("MM", "mm");
                        string output = borderSeq.OutPutCsvString();
                        outstrlist.Add(output);
                    }
                    else if (strLine.IndexOf("HDrillSequence,") == 0)
                    {
                        ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                        HDrillSequenceEntity hdrillSeq = new HDrillSequenceEntity(strLine);
                        hdrillSeq.HDrillFeedSpeed = "3000";
                        hdrillSeq.HDrillEntrySpeed = "2700";
                        string output = hdrillSeq.OutPutCsvString();
                        outstrlist.Add(output);
                    }
                    else if (strLine.IndexOf("VdrillSequence,") == 0)
                    {
                        ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                        VdrillSequenceEntity vdrillSeq = new VdrillSequenceEntity(strLine);
                        vdrillSeq.VDrillFeedSpeed = "3000";
                        vdrillSeq.VDrillEntrySpeed = "8000";
                        string output = vdrillSeq.OutPutCsvString();
                        outstrlist.Add(output);
                    }
                    else if (strLine.IndexOf("RouteSetMillSequence,") == 0)
                    {
                        ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                        RouteSetIndex++;
                        RouteSetMillSequenceEntity routeSeq = new RouteSetMillSequenceEntity(strLine);
                        routeSeq.RouteSetMillCounter = RouteSetIndex.ToString();
                        routeSeq.RouteVectorCounter = "";
                        routeSeq.RouteVectorCount = "";
                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118
                        if (Convert.ToDouble(routeSeq.RouteToolComp) == 1)
                            routeSeq.RouteToolComp = "2";

                        else if (Convert.ToDouble(routeSeq.RouteToolComp) == 2)
                            routeSeq.RouteToolComp = "1";

                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118

                        Seqlist.Add(routeSeq);
                    }
                    else if (strLine.IndexOf("RouteSequence,") == 0)
                    {
                        routeswqindex++;
                        RouteSetMillSequenceEntity routeSeq = new RouteSetMillSequenceEntity(strLine);
                        routeSeq.RouteSetMillCounter = RouteSetIndex.ToString();
                        routeSeq.RouteVectorCounter = routeswqindex.ToString();
                        routeSeq.RouteVectorCount = "";
                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118
                        if (Convert.ToDouble(routeSeq.RouteToolComp) == 1)
                            routeSeq.RouteToolComp = "2";

                        else if (Convert.ToDouble(routeSeq.RouteToolComp) == 2)
                            routeSeq.RouteToolComp = "1";

                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118
                        Seqlist.Add(routeSeq);
                    }
                    strLine = sr.ReadLine();
                }
                sr.Close();
                ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                string TempFile = Path.ChangeExtension(csvFilename, "temp");
                StreamWriter sw = new StreamWriter(TempFile, false, Encoding.Default);
                foreach (string str in outstrlist)
                {
                    if (string.IsNullOrEmpty(str))
                        continue;
                    sw.WriteLine(str);
                }
                sw.Flush();
                sw.Close();
                return true;
            }
            catch (SystemException ex)
            {
                return false;
            }

        }
        public bool ModifyDataFromCsv(string csvFilename, bool bface6)
        {
            int thickness = 0; //发现水平孔位的Z值有超过板厚的数据。将这数据过滤掉 宋新刚 20171101
            try
            {
                StreamReader sr = new StreamReader(csvFilename, Encoding.Default);
                string strLine = sr.ReadLine();
                //strLine = sr.ReadToEnd();
                int RouteSetIndex = 0;
                int routeswqindex = 0;
                ArrayList Seqlist = new ArrayList();
                Seqlist.Clear();
                bool hasend = false;
                ArrayList outstrlist = new ArrayList();
                BorderSequenceEntity borderSeq = new BorderSequenceEntity(strLine);

                while (strLine != null)
                {
                    strLine = strLine.Replace("D8.5", "129").Replace("D10", "130");
                    if (strLine.IndexOf("EndSequence,") == 0)
                    {
                        hasend = true;
                        outstrlist.Add(strLine);
                    }
                    ////处理每行的字符串
                    else if (strLine.IndexOf("BorderSequence,") == 0)
                    {
                        borderSeq.NestRoutesCount = "0";
                        borderSeq.Edgeband1 = borderSeq.Edgeband1.Replace("MM", "mm");
                        borderSeq.Edgeband2 = borderSeq.Edgeband1.Replace("MM", "mm");
                        borderSeq.Edgeband3 = borderSeq.Edgeband1.Replace("MM", "mm");
                        borderSeq.Edgeband4 = borderSeq.Edgeband1.Replace("MM", "mm");

                        if (borderSeq.CurrentZoneName == "AD")
                        {
                            borderSeq.CurrentZoneName = "M";
                            borderSeq.RunField = "3";
                        }

                        if (borderSeq.CurrentZoneName == "DA")
                        {
                            borderSeq.CurrentZoneName = "N";
                            borderSeq.RunField = "4";
                        }

                        thickness = Convert.ToInt32(borderSeq.PanelThickness);  //发现水平孔位的Z值有超过板厚的数据。将这数据过滤掉 宋新刚 20171101

                        if (bface6)
                            borderSeq.Face6FileName = Path.GetFileNameWithoutExtension(csvFilename);
                        else
                            borderSeq.FileName = Path.GetFileNameWithoutExtension(csvFilename);

                        string output = borderSeq.OutPutCsvString();
                        outstrlist.Add(output);
                    }
                    else if (strLine.IndexOf("HDrillSequence,") == 0)
                    {
                        ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                        HDrillSequenceEntity hdrillSeq = new HDrillSequenceEntity(strLine);

                        hdrillSeq.HDrillFeedSpeed = "3000";
                        hdrillSeq.HDrillEntrySpeed = "2700";
                        /***  用于解决水平孔加工面产生小数点的问题   宋新刚20161115 * **/
                        if (hdrillSeq.CurrentFace == "1.0000") hdrillSeq.CurrentFace = "1";
                        else if (hdrillSeq.CurrentFace == "2.0000") hdrillSeq.CurrentFace = "2";
                        else if (hdrillSeq.CurrentFace == "3.0000") hdrillSeq.CurrentFace = "3";
                        else if (hdrillSeq.CurrentFace == "4.0000") hdrillSeq.CurrentFace = "4";

                        // 以下是解决工程部 双向拉杆有水平孔为0的情况
                        if (hdrillSeq.HDrillDiameter == "0" || hdrillSeq.HDrillDiameter == "0.0000")
                        {
                            strLine = sr.ReadLine();
                            continue;
                        }

                        // 以下是解决工程部 水平孔的坐标出现负数的情况
                        if (hdrillSeq.HDrillZ.IndexOf("-") == 0 || hdrillSeq.HDrillX.IndexOf("-") == 0 || hdrillSeq.HDrillY.IndexOf("-") == 0)
                        {
                            // MessageBox.Show(" 以下是解决工程部 水平孔的坐标出现负数的情况");
                            strLine = sr.ReadLine();
                            continue;
                        }
                        // 宋新刚 20161114

                        // 以下是解决舟山翻床厚背板镜像加工 在面1和面2上多出现10的水平孔 通过式PTP无法加工   
                        if (hdrillSeq.HDrillDiameter == "10.0000" || hdrillSeq.HDrillDiameter == "10")
                        {
                            strLine = sr.ReadLine();
                            continue;

                        }  //证明水平孔不可能有10的情况。遇到有10的水平孔全部删除

                        // 以下是发现水平孔位的Z值有超过板厚的数据。将这数据过滤掉 宋新刚 20171101
                        if (Convert.ToDouble(hdrillSeq.HDrillY) > thickness)
                        {
                            strLine = sr.ReadLine();
                            continue;
                        }

                        string output = hdrillSeq.OutPutCsvString();
                        outstrlist.Add(output);
                    }
                    else if (strLine.IndexOf("VdrillSequence,") == 0)
                    {
                        ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                        VdrillSequenceEntity vdrillSeq = new VdrillSequenceEntity(strLine);
                        vdrillSeq.VDrillFeedSpeed = "3000";
                        vdrillSeq.VDrillEntrySpeed = "8000";
                        // 以下是解决工程部 垂直孔的坐标出现负数的情况
                        if (vdrillSeq.VDrillX.IndexOf("-") == 0 || vdrillSeq.VDrillY.IndexOf("-") == 0 || vdrillSeq.VDrillZ.IndexOf("-") == 0)
                        {
                            strLine = sr.ReadLine();
                            continue;
                        }
                        //以下是解决工程部模型 孔位超出板材范围的情况
                        if (Convert.ToDouble(vdrillSeq.VDrillY) - Convert.ToDouble(borderSeq.PanelWidth) > 0)
                        {
                            strLine = sr.ReadLine();
                            continue;
                        }

                        if (Convert.ToDouble(vdrillSeq.VDrillX) - Convert.ToDouble(borderSeq.PanelLength) > 0)
                        {
                            strLine = sr.ReadLine();
                            continue;
                        }
                        // 宋新刚 20161114
                        string output = vdrillSeq.OutPutCsvString();
                        outstrlist.Add(output);
                    }
                    else if (strLine.IndexOf("RouteSetMillSequence,") == 0)
                    {
                        ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                        RouteSetIndex++;
                        RouteSetMillSequenceEntity routeSeq = new RouteSetMillSequenceEntity(strLine);
                        routeSeq.RouteSetMillCounter = RouteSetIndex.ToString();
                        routeSeq.RouteVectorCounter = "";
                        routeSeq.RouteVectorCount = "";
                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118
                        if (Convert.ToDouble(routeSeq.RouteToolComp) == 1)
                            routeSeq.RouteToolComp = "2";

                        else if (Convert.ToDouble(routeSeq.RouteToolComp) == 2)
                            routeSeq.RouteToolComp = "1";

                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118

                        if (Convert.ToDouble(routeSeq.RouteSetMillZ) == 9.25 && Convert.ToDouble(routeSeq.RouteDiameter) == 8.5 && Convert.ToDouble(routeSeq.RouteZ) == 9.25)  //发现AD过来的背板槽 有固定的9.25深  修复为5mm深  宋新刚 2018年1月26日
                        {
                            routeSeq.RouteSetMillZ = "5";
                            routeSeq.RouteZ = "5";
                        }

                        Seqlist.Add(routeSeq);
                    }
                    else if (strLine.IndexOf("RouteSequence,") == 0)
                    {
                        routeswqindex++;
                        RouteSetMillSequenceEntity routeSeq = new RouteSetMillSequenceEntity(strLine);
                        routeSeq.RouteSetMillCounter = RouteSetIndex.ToString();
                        routeSeq.RouteVectorCounter = routeswqindex.ToString();
                        routeSeq.RouteVectorCount = "";
                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118
                        if (Convert.ToDouble(routeSeq.RouteToolComp) == 1)
                            routeSeq.RouteToolComp = "2";

                        else if (Convert.ToDouble(routeSeq.RouteToolComp) == 2)
                            routeSeq.RouteToolComp = "1";

                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118

                        if (Convert.ToDouble(routeSeq.RouteSetMillZ) == 9.25 && Convert.ToDouble(routeSeq.RouteDiameter) == 8.5 && Convert.ToDouble(routeSeq.RouteZ) == 9.25)  //发现AD过来的背板槽 有固定的9.25深  修复为5mm深  宋新刚 2018年1月26日
                        {
                            routeSeq.RouteSetMillZ = "5";
                            routeSeq.RouteZ = "5";
                        }

                        Seqlist.Add(routeSeq);
                    }
                    strLine = sr.ReadLine();
                }
                sr.Close();
                ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                //string TempFile = Path.ChangeExtension(csvFilename, "temp");
                StreamWriter sw = new StreamWriter(csvFilename, false, Encoding.Default);
                foreach (string str in outstrlist)
                {
                    if (string.IsNullOrEmpty(str))
                        continue;
                    sw.WriteLine(str);
                }
                if (!hasend)
                    sw.WriteLine("EndSequence,,,,");
                sw.Flush();
                sw.Close();
                return true;
            }
            catch (SystemException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }

        }

        public bool ModifyDataFromCsv(string csvFilename, bool bface6, ref double sheetWidth, ref double sheetLength)
        {

            int thickness = 0; //发现水平孔位的Z值有超过板厚的数据。将这数据过滤掉 宋新刚 20171101
            try
            {
                StreamReader sr = new StreamReader(csvFilename, Encoding.Default);
                string strLine = sr.ReadLine();
                //strLine = sr.ReadToEnd();
                int RouteSetIndex = 0;
                int routeswqindex = 0;
                ArrayList Seqlist = new ArrayList();
                Seqlist.Clear();
                bool hasend = false;
                ArrayList outstrlist = new ArrayList();
                BorderSequenceEntity borderSeq = new BorderSequenceEntity(strLine);
                sheetWidth = 0;
                double.TryParse(borderSeq.PanelWidth, out sheetWidth);
                sheetLength = 0;
                double.TryParse(borderSeq.PanelLength, out sheetLength);

                while (strLine != null)
                {
                    strLine = strLine.Replace("D8.5", "129").Replace("D10", "130");
                    if (strLine.IndexOf("EndSequence,") == 0)
                    {
                        hasend = true;
                        outstrlist.Add(strLine);
                    }
                    ////处理每行的字符串
                    else if (strLine.IndexOf("BorderSequence,") == 0)
                    {
                        borderSeq.NestRoutesCount = "0";
                        borderSeq.Edgeband1 = borderSeq.Edgeband1.Replace("MM", "mm");
                        borderSeq.Edgeband2 = borderSeq.Edgeband1.Replace("MM", "mm");
                        borderSeq.Edgeband3 = borderSeq.Edgeband1.Replace("MM", "mm");
                        borderSeq.Edgeband4 = borderSeq.Edgeband1.Replace("MM", "mm");

                        if (borderSeq.CurrentZoneName == "AD")
                        {
                            borderSeq.CurrentZoneName = "M";
                            borderSeq.RunField = "3";
                        }

                        if (borderSeq.CurrentZoneName == "DA")
                        {
                            borderSeq.CurrentZoneName = "N";
                            borderSeq.RunField = "4";
                        }

                        thickness = Convert.ToInt32(borderSeq.PanelThickness);  //发现水平孔位的Z值有超过板厚的数据。将这数据过滤掉 宋新刚 20171101

                        if (bface6)
                            borderSeq.Face6FileName = Path.GetFileNameWithoutExtension(csvFilename);
                        else
                            borderSeq.FileName = Path.GetFileNameWithoutExtension(csvFilename);

                        string output = borderSeq.OutPutCsvString();
                        outstrlist.Add(output);
                    }
                    else if (strLine.IndexOf("HDrillSequence,") == 0)
                    {
                        ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                        HDrillSequenceEntity hdrillSeq = new HDrillSequenceEntity(strLine);

                        hdrillSeq.HDrillFeedSpeed = "3000";
                        hdrillSeq.HDrillEntrySpeed = "2700";
                        /***  用于解决水平孔加工面产生小数点的问题   宋新刚20161115 * **/
                        if (hdrillSeq.CurrentFace == "1.0000") hdrillSeq.CurrentFace = "1";
                        else if (hdrillSeq.CurrentFace == "2.0000") hdrillSeq.CurrentFace = "2";
                        else if (hdrillSeq.CurrentFace == "3.0000") hdrillSeq.CurrentFace = "3";
                        else if (hdrillSeq.CurrentFace == "4.0000") hdrillSeq.CurrentFace = "4";

                        // 以下是解决工程部 双向拉杆有水平孔为0的情况
                        if (hdrillSeq.HDrillDiameter == "0" || hdrillSeq.HDrillDiameter == "0.0000")
                        {
                            strLine = sr.ReadLine();
                            continue;
                        }
                        // 以下是发现水平孔位的Z值有超过板厚的数据。将这数据过滤掉 宋新刚 20171101
                        if (Convert.ToDouble(hdrillSeq.HDrillY) > thickness) //发现其高度值有小数。所以这里必须得是Double类型   宋新刚20171106
                        {
                            strLine = sr.ReadLine();
                            continue;
                        }
                        // 以下是解决工程部 水平孔的坐标出现负数的情况
                        if (hdrillSeq.HDrillZ.IndexOf("-") == 0 || hdrillSeq.HDrillX.IndexOf("-") == 0 || hdrillSeq.HDrillY.IndexOf("-") == 0)
                        {
                           // MessageBox.Show(" 以下是解决工程部 水平孔的坐标出现负数的情况");
                            strLine = sr.ReadLine();
                            continue;
                        }
                        // 宋新刚 20161114

                        // 以下是解决舟山翻床厚背板镜像加工 在面1和面2上多出现10的水平孔 通过式PTP无法加工   
                        if (hdrillSeq.HDrillDiameter == "10.0000" || hdrillSeq.HDrillDiameter == "10")
                        {
                            strLine = sr.ReadLine();
                            continue;

                        }  //证明水平孔不可能有10的情况。遇到有10的水平孔全部删除

                        string output = hdrillSeq.OutPutCsvString();
                        outstrlist.Add(output);
                    }
                    else if (strLine.IndexOf("VdrillSequence,") == 0)
                    {
                        ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                        VdrillSequenceEntity vdrillSeq = new VdrillSequenceEntity(strLine);
                        vdrillSeq.VDrillFeedSpeed = "3000";
                        vdrillSeq.VDrillEntrySpeed = "8000";
                        // 以下是解决工程部 垂直孔的坐标出现负数的情况
                        if (vdrillSeq.VDrillX.IndexOf("-") == 0 || vdrillSeq.VDrillY.IndexOf("-") == 0 || vdrillSeq.VDrillZ.IndexOf("-") == 0)
                        {
                            strLine = sr.ReadLine();
                            continue;
                        }
                        //以下是解决工程部模型 孔位超出板材范围的情况
                        if (Convert.ToDouble(vdrillSeq.VDrillY) - Convert.ToDouble(borderSeq.PanelWidth) > 0)
                        {
                            strLine = sr.ReadLine();
                            continue;
                        }

                        if (Convert.ToDouble(vdrillSeq.VDrillX) - Convert.ToDouble(borderSeq.PanelLength) > 0)
                        {
                            strLine = sr.ReadLine();
                            continue;
                        }
                        // 宋新刚 20161114
                        string output = vdrillSeq.OutPutCsvString();
                        outstrlist.Add(output);
                    }
                    else if (strLine.IndexOf("RouteSetMillSequence,") == 0)
                    {
                        ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                        RouteSetIndex++;
                        RouteSetMillSequenceEntity routeSeq = new RouteSetMillSequenceEntity(strLine);
                        routeSeq.RouteSetMillCounter = RouteSetIndex.ToString();
                        routeSeq.RouteVectorCounter = "";
                        routeSeq.RouteVectorCount = "";
                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118
                        if (Convert.ToDouble(routeSeq.RouteToolComp) == 1)
                            routeSeq.RouteToolComp = "2";

                        else if (Convert.ToDouble(routeSeq.RouteToolComp) == 2)
                            routeSeq.RouteToolComp = "1";

                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118

                        if (Convert.ToDouble(routeSeq.RouteSetMillZ) == 9.25 && Convert.ToDouble(routeSeq.RouteDiameter) == 8.5 && Convert.ToDouble(routeSeq.RouteZ) == 9.25)  //发现AD过来的背板槽 有固定的9.25深  修复为5mm深  宋新刚 2018年1月26日
                        {
                            routeSeq.RouteSetMillZ = "5";
                            routeSeq.RouteZ = "5";
                        }

                        Seqlist.Add(routeSeq);
                    }
                    else if (strLine.IndexOf("RouteSequence,") == 0)
                    {
                        routeswqindex++;
                        RouteSetMillSequenceEntity routeSeq = new RouteSetMillSequenceEntity(strLine);
                        routeSeq.RouteSetMillCounter = RouteSetIndex.ToString();
                        routeSeq.RouteVectorCounter = routeswqindex.ToString();
                        routeSeq.RouteVectorCount = "";
                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118
                        if (Convert.ToDouble(routeSeq.RouteToolComp) == 1)
                            routeSeq.RouteToolComp = "2";

                        else if (Convert.ToDouble(routeSeq.RouteToolComp) == 2)
                            routeSeq.RouteToolComp = "1";

                        // 发现AD出来的只要有偏置的地方都需要对调一下   宋新刚  20170118

                        if (Convert.ToDouble(routeSeq.RouteSetMillZ) == 9.25 && Convert.ToDouble(routeSeq.RouteDiameter) == 8.5 && Convert.ToDouble(routeSeq.RouteZ) == 9.25)  //发现AD过来的背板槽 有固定的9.25深  修复为5mm深  宋新刚 2018年1月26日
                        {
                            routeSeq.RouteSetMillZ = "5";
                            routeSeq.RouteZ = "5";
                        }


                        Seqlist.Add(routeSeq);
                    }
                    strLine = sr.ReadLine();
                }
                sr.Close();
                ModifyRouteSeqV2(bface6, borderSeq, Seqlist, ref routeswqindex, ref outstrlist);

                //string TempFile = Path.ChangeExtension(csvFilename, "temp");
                StreamWriter sw = new StreamWriter(csvFilename, false, Encoding.Default);
                foreach (string str in outstrlist)
                {
                    if (string.IsNullOrEmpty(str))
                        continue;
                    sw.WriteLine(str);
                }
                if (!hasend)
                    sw.WriteLine("EndSequence,,,,");
                sw.Flush();
                sw.Close();
                return true;
            }
            catch (SystemException ex)
            {
                MessageBox.Show("代码转换出现异常   " + ex.Message + "\n" + csvFilename);
                return false;
            }

        }

        void ModifyRouteSeqV(ArrayList Seqlist, ref int routeswqindex, ref ArrayList outstrlist)
        {
            if (Seqlist.Count > 0)
            {
                for (int i = 0; i < Seqlist.Count; i++)
                {

                    RouteSetMillSequenceEntity route = Seqlist[i] as RouteSetMillSequenceEntity;
                    route.RouteVectorCount = (Seqlist.Count - 1).ToString();
                    if (!string.IsNullOrEmpty(route.RouteBulge) && route.RouteBulge != "0.0000" && route.RouteBulge != "0")
                    {

                        if (i > 0)
                        {
                            RouteSetMillSequenceEntity route1 = Seqlist[i - 1] as RouteSetMillSequenceEntity;
                            double x1 = 0, x2 = 0, y1 = 0, y2 = 0, l = 0, u = 0, radius = 0, ang = 0, cenx = 0, ceny = 0;

                            double.TryParse(route1.RouteX, out x1);
                            double.TryParse(route1.RouteY, out y1);

                            double.TryParse(route.RouteX, out x2);
                            double.TryParse(route.RouteY, out y2);

                            double.TryParse(route.RouteBulge, out u);
                            l = Math.Sqrt(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2));
                            radius = 0.25 * l * (u + 1 / u);
                            route.RouteRadius = string.Format("{0:f4}", radius);
                            if (x1 == x2)
                            {
                                ang = Math.PI / 2;
                                if (u > 0)
                                {
                                    cenx = (x1 + x2) / 2 + (radius - l * u / 2);
                                    ceny = (y1 + y2) / 2;
                                    if ((x2 - x1) * (ceny - y2) - (y2 - y1) * (cenx - x2) < 0)
                                    {
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                    else
                                    {
                                        cenx = (x1 + x2) / 2 - (radius - l * u / 2);
                                        ceny = (y1 + y2) / 2;
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                }
                                else
                                {
                                    cenx = (x1 + x2) / 2 - (radius - l * u / 2);
                                    ceny = (y1 + y2) / 2;
                                    if ((x2 - x1) * (ceny - y2) - (y2 - y1) * (cenx - x2) < 0)
                                    {
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                    else
                                    {
                                        cenx = (x1 + x2) / 2 + (radius - l * u / 2);
                                        ceny = (y1 + y2) / 2;
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
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
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                    else
                                    {
                                        cenx = (x1 + x2) / 2 - (radius - l * u / 2) * Math.Sin(ang);
                                        ceny = (y1 + y2) / 2 + (radius - l * u / 2) * Math.Cos(ang);
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                }
                                else
                                {
                                    cenx = (x1 + x2) / 2 - (radius - l * u / 2) * Math.Sin(ang);
                                    ceny = (y1 + y2) / 2 + (radius - l * u / 2) * Math.Cos(ang);
                                    if ((x2 - x1) * (ceny - y2) - (y2 - y1) * (cenx - x2) < 0)
                                    {
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                    else
                                    {
                                        cenx = (x1 + x2) / 2 + (radius - l * u / 2) * Math.Sin(ang);
                                        ceny = (y1 + y2) / 2 - (radius - l * u / 2) * Math.Cos(ang);
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                }
                            }
                        }

                    }
                    outstrlist.Add(route.OutPutCsvString());
                }
                routeswqindex = 0;
                Seqlist.Clear();
            }
        }

        void ModifyRouteSeqV2(bool bface6, BorderSequenceEntity borderSeq, ArrayList Seqlist, ref int routeswqindex, ref ArrayList outstrlist)
        {
            if (Seqlist.Count > 0)
            {
                if (bface6)
                {

                    for (int i = 0; i < Seqlist.Count; i++)
                    {
                        double d1 = Convert.ToDouble(borderSeq.PanelThickness);
                        double c1 = Convert.ToDouble(borderSeq.PanelLength);

                        RouteSetMillSequenceEntity route = Seqlist[i] as RouteSetMillSequenceEntity;
                        double di = Convert.ToDouble(route.RouteSetMillZ);
                        double pi = Convert.ToDouble(route.RouteX);

                        if (di < d1 && (pi <= 0 || pi >= c1) && route.RouteToolName != "D3") //开槽发现有D3刀开的槽是需要过滤的，不需要开槽功能  宋新刚20170823
                        {
                            VdrillSequenceEntity vdrill = new VdrillSequenceEntity();
                            double drillx = Convert.ToDouble(route.RouteX);
                            if (drillx <= 0)
                                vdrill.VDrillX = "0.1";
                            else if (drillx >= c1)
                                vdrill.VDrillX = (c1 - 0.1).ToString();

                            if (Convert.ToDouble(route.RouteToolComp) == 0)
                                vdrill.VDrillY = route.RouteSetMillY;
                            else if (Convert.ToDouble(route.RouteToolComp) == 1)
                                vdrill.VDrillY = (Convert.ToDouble(route.RouteSetMillY) + 0.5 * Convert.ToDouble(route.RouteDiameter)).ToString();
                            else if (Convert.ToDouble(route.RouteToolComp) == 2)
                                vdrill.VDrillY = (Convert.ToDouble(route.RouteSetMillY) - 0.5 * Convert.ToDouble(route.RouteDiameter)).ToString();

                            vdrill.VdrillSequence = "VdrillSequence";
                            vdrill.VDrillZ = route.RouteSetMillZ;
                            vdrill.VDrillXOffset = "0";
                            vdrill.VDrillYOffset = "0";
                            vdrill.VDrillDiameter = "12.2";
                            vdrill.VDrillToolName = "12.2mm";
                            vdrill.VDrillFeedSpeed = "3000";
                            vdrill.VDrillEntrySpeed = "8000";

                            if (borderSeq.Description.Trim() == "翻门板")
                                continue;

                            outstrlist.Add(vdrill.OutPutCsvString());
                        }
                    }
                }
                for (int i = 0; i < Seqlist.Count; i++)
                {
                    RouteSetMillSequenceEntity route = Seqlist[i] as RouteSetMillSequenceEntity;
                    route.RouteVectorCount = (Seqlist.Count - 1).ToString();

                    // 解决翻板门板矩形孔出现波浪型线条问题  宋新刚 20170114
                    if (route.RouteVectorCount == "19" && borderSeq.Description.IndexOf("门板")>-1)
                        continue;

                    if (Convert.ToDouble(route.RouteVectorCount) > 20)  //发现AD有切大的矩形  在这里如果拐点大于20就删   宋新刚 20170309
                        continue;

                    if (route.RouteVectorCount == "3" && (borderSeq.Description.Equals("MB-左开门板(横纹)") || borderSeq.Description.Equals("MB-右开门板(横纹)")))
                        continue;   //发现AD过来的时候，对于门板横纹 做拉手孔的时候也出现波浪线的问题。做了修复  宋新刚20170322

                    if (route.RouteVectorCount == "3" && (borderSeq.Description.IndexOf("MB-榻榻米门板") > -1 && borderSeq.Description.IndexOf("(横纹)") > -1))
                        continue;   //发现AD过来的时候，对新增加的名字做了修复  宋新刚20170829

                    if (route.RouteVectorCount == "3" && (borderSeq.Description.IndexOf("门板") > -1))
                        continue;   //发现AD过来的时候，对新增加的名字做了修复  宋新刚20171024

                    //if (route.RouteVectorCount == "1" && borderSeq.Description.Trim() == "翻门板")
                    //    continue;   //发现多出来的部分是切除多余的板数据的  所以取消  宋新刚 20170118

                    // 解决翻板门板矩形孔出现波浪型线条问题  宋新刚 20170114

                    if (!string.IsNullOrEmpty(route.RouteBulge) && route.RouteBulge != "0.0000" && route.RouteBulge != "0")
                    {
                        if (i > 0)
                        {
                            RouteSetMillSequenceEntity route1 = Seqlist[i - 1] as RouteSetMillSequenceEntity;
                            double x1 = 0, x2 = 0, y1 = 0, y2 = 0, l = 0, u = 0, radius = 0, ang = 0, cenx = 0, ceny = 0;

                            double.TryParse(route1.RouteX, out x1);
                            double.TryParse(route1.RouteY, out y1);

                            double.TryParse(route.RouteX, out x2);
                            double.TryParse(route.RouteY, out y2);

                            double.TryParse(route.RouteBulge, out u);
                            l = Math.Sqrt(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2));
                            radius = 0.25 * l * (u + 1 / u);
                            route.RouteRadius = string.Format("{0:f4}", radius);
                            if (x1 == x2)
                            {
                                ang = Math.PI / 2;
                                if (u > 0)
                                {
                                    cenx = (x1 + x2) / 2 + (radius - l * u / 2);
                                    ceny = (y1 + y2) / 2;
                                    if ((x2 - x1) * (ceny - y2) - (y2 - y1) * (cenx - x2) < 0)
                                    {
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                    else
                                    {
                                        cenx = (x1 + x2) / 2 - (radius - l * u / 2);
                                        ceny = (y1 + y2) / 2;
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                }
                                else
                                {
                                    cenx = (x1 + x2) / 2 - (radius - l * u / 2);
                                    ceny = (y1 + y2) / 2;
                                    if ((x2 - x1) * (ceny - y2) - (y2 - y1) * (cenx - x2) < 0)
                                    {
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                    else
                                    {
                                        cenx = (x1 + x2) / 2 + (radius - l * u / 2);
                                        ceny = (y1 + y2) / 2;
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
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
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                    else
                                    {
                                        cenx = (x1 + x2) / 2 - (radius - l * u / 2) * Math.Sin(ang);
                                        ceny = (y1 + y2) / 2 + (radius - l * u / 2) * Math.Cos(ang);
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                }
                                else
                                {
                                    cenx = (x1 + x2) / 2 - (radius - l * u / 2) * Math.Sin(ang);
                                    ceny = (y1 + y2) / 2 + (radius - l * u / 2) * Math.Cos(ang);
                                    if ((x2 - x1) * (ceny - y2) - (y2 - y1) * (cenx - x2) < 0)
                                    {
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                    else
                                    {
                                        cenx = (x1 + x2) / 2 + (radius - l * u / 2) * Math.Sin(ang);
                                        ceny = (y1 + y2) / 2 - (radius - l * u / 2) * Math.Cos(ang);
                                        route.RouteCenterX = string.Format("{0:f4}", cenx); ;
                                        route.RouteCenterY = string.Format("{0:f4}", ceny); ;
                                    }
                                }
                            }
                        }

                    }
                    outstrlist.Add(route.OutPutCsvString());
                }
                routeswqindex = 0;
                Seqlist.Clear();
            }
        }

        public bool SplitBoardCsvBySize(string csvFilename, decimal headerlineCount, double minlength, double maxlength)
        {
            StreamReader sr = new StreamReader(csvFilename, Encoding.Default);
            List<nesttingcsvEntity> nestlistA = new List<nesttingcsvEntity>();
            List<nesttingcsvEntity> nestlistB = new List<nesttingcsvEntity>();
            List<nesttingcsvEntity> nestlistC = new List<nesttingcsvEntity>();
            string strLine = sr.ReadLine();
            int index = 0;
            string header = "";
            while (strLine != null)
            {
                index++;
                if (index <= headerlineCount)
                    header += strLine + "\r\n";
                else
                {
                    nesttingcsvEntity nestinginfo = new nesttingcsvEntity(strLine);
                    double dh = 0; double.TryParse(nestinginfo.heitht, out dh);
                    double dw = 0; double.TryParse(nestinginfo.width, out dw);
                    if (Math.Min(dh, dw) <= minlength)
                        nestlistC.Add(nestinginfo);
                    else if (dw >= maxlength || dh >= maxlength)
                        nestlistA.Add(nestinginfo);
                    else
                        nestlistB.Add(nestinginfo);
                }

                strLine = sr.ReadLine();
            }
            header = header.Substring(0, header.Length - 4);
            sr.Close();

            string TempFile = Path.Combine(Path.GetDirectoryName(csvFilename), "A" + Path.GetFileNameWithoutExtension(csvFilename) + ".csv");
            StreamWriter sw1 = new StreamWriter(TempFile, false, Encoding.Default);
            sw1.WriteLine(header);
            foreach (nesttingcsvEntity nestinginfo in nestlistA)
            {
                string str = nestinginfo.OutPutCsvString();
                sw1.WriteLine(str);
            }
            sw1.Flush();
            sw1.Close();

            TempFile = Path.Combine(Path.GetDirectoryName(csvFilename), "B" + Path.GetFileNameWithoutExtension(csvFilename) + ".csv");
            StreamWriter sw2 = new StreamWriter(TempFile, false, Encoding.Default);
            sw2.WriteLine(header);
            foreach (nesttingcsvEntity nestinginfo in nestlistB)
            {
                string str = nestinginfo.OutPutCsvString();
                sw2.WriteLine(str);
            }
            sw2.Flush();
            sw2.Close();

            TempFile = Path.Combine(Path.GetDirectoryName(csvFilename), "C" + Path.GetFileNameWithoutExtension(csvFilename) + ".csv");
            StreamWriter sw3 = new StreamWriter(TempFile, false, Encoding.Default);
            sw3.WriteLine(header);
            foreach (nesttingcsvEntity nestinginfo in nestlistC)
            {
                string str = nestinginfo.OutPutCsvString();
                sw3.WriteLine(str);
            }
            sw3.Flush();
            sw3.Close();

            return true;
        }
    }
}
