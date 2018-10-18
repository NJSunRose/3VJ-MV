using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.IO.Compression;
using System.Web;

namespace SWJXMLToCSV
{
    public class Invoke
    {
        public static List<string> OutputCsvByXml(string xmlString,string orderNo=null) {
            List<Cabinet> list = T3VProduct.LoadFromXML(xmlString);
            List<string> files= ConvertCore.Output(list, orderNo);
            var newFiles = files.Where(p => !p.Contains(".jpg")).ToList();
            GZip(newFiles);
            return files;
        }

        /// <summary>
        /// 合并订单行解析xml，不打压缩包
        /// </summary>
        /// <param name="xmlString"></param>
        /// <param name="orderNo"></param>
        /// <returns></returns>
        public static List<string> OutputCsvByXml1(string xmlString, string orderNo = null)
        {
            List<Cabinet> list = T3VProduct.LoadFromXML(xmlString);
            List<string> files = ConvertCore.Output(list, orderNo);
            var newFiles = files.Where(p => !p.Contains(".jpg")).ToList();
            return files;
        }
        public static List<string> OutputCsvByXml(List<Cabinet> list, string orderNo = null) {
            List<string> files = ConvertCore.Output(list, orderNo);
            GZip(files);
            return files;
        }


       
        public static List<string> CSVGzip(string orderNo, string mapping)
        {
            if (string.IsNullOrEmpty(orderNo)) {
                throw new Exception("d订单号为空");
            }
            if (string.IsNullOrEmpty(mapping))
            {
                throw new Exception("csv目录与压缩文件映射关系为空");
            }
            List<string> imgList = new List<string>();
            //获取当前系统path路径
            var path = HttpContext.Current.Server.MapPath("~") + "\\cdwj\\"+ orderNo;
            //检查是否有project目录，没有的话创建
            if (!Directory.Exists(path + "\\Project\\"))
            {
                Directory.CreateDirectory(path+ "\\Project\\");
            }
            #region  处理传过来的mapping
            List<CSVGzipEntity> list = new List<CSVGzipEntity>();
            var str1 = mapping.Split(',');
            if (str1 != null && str1.Length > 0) {
                foreach (var i in str1) {
                    var str2=i.Split('=');
                    if (str2 != null && str2.Length > 0) {
                        CSVGzipEntity model = new CSVGzipEntity();
                        model.hanghao = str2[0];
                        model.wenjianjia=str2[1].Split('+');
                        list.Add(model);
                    }
                }
            }
            #endregion
            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            #region 创建压缩文件夹路径
            if (list.Count() > 0) {
                foreach (var i in list)
                {
                    //压缩文件名
                    string zipName = orderNo + "_" + i.hanghao + "_DMSCSV.zip";
                    string zipFullPath = path + "\\Project\\" + zipName;
                    if (!dic.ContainsKey(zipFullPath)) {
                        dic.Add(zipFullPath, new List<string>());
                    }
                    //读取文件夹下的文件，循环添加到压缩包
                    foreach (var j in i.wenjianjia) {
                        var dePath = path + "\\" + j;
                        DirectoryInfo dir = new DirectoryInfo(dePath);
                        //读取图片的路径
                        List<FileInfo> img = new List<FileInfo>();
                        List<FileInfo> csvlist = new List<FileInfo>();
                        FileInfo[] files = null;
                        try {
                            files = dir.GetFiles();
                        } catch (Exception e) {

                        }
                        //try {
                        //   img = dir.GetFiles("*.jpg");
                        //}
                        //catch (Exception e) {
                        //}
                        if (files != null && files.Count() > 0) {
                            foreach (var m in files)
                            {
                                if (m.Extension.Contains(".jpg"))
                                {
                                    img.Add(m);
                                }
                                if (m.Extension.Contains(".csv"))
                                {
                                    csvlist.Add(m);
                                }
                            }
                        }
                      
                        if (img != null && img.Count() > 0) {
                            foreach (var m in img)
                            {
                                imgList.Add(m.FullName);
                            }
                        }
                      
                        if (csvlist != null && csvlist.Count() > 0) {
                            foreach (var m in csvlist)
                            {
                                dic[zipFullPath].Add(m.FullName);
                            }
                        }
                    } 
                }
            }
            #endregion
            
            //打压缩包
            foreach (var item in dic)
            {
                ZipHelper.MultiZip(item.Key, item.Value.ToArray());
            }

            return imgList;
           
        }





        private static void GZip(List<string> files)
        {
            string number = string.Empty;
            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            foreach (string file in files)
            {
                //获取文件路径
                string fileDir = Path.GetDirectoryName(file);

                string[] dirArr = fileDir.Split('\\');
                string num = dirArr.Last();
                dirArr[dirArr.Length - 1] = string.Empty;
                string zipFullPath = string.Empty;

                string zipPath = string.Join("\\",dirArr) + "\\Project\\";
                if (!Directory.Exists(zipPath))
                    Directory.CreateDirectory(zipPath);
                string zipName = dirArr[dirArr.Length - 2] + "_" + num + "_DMSCSV.zip";
                //ZipHelper.ZipFile(file, zipPath + zipName);
                zipFullPath = zipPath + zipName;
                if(!dic.ContainsKey(zipFullPath))
                    dic.Add(zipFullPath, new List<string>());
                //}
                dic[zipFullPath].Add(file);
                number = num;
            }

            //打压缩包
            foreach (var item in dic)
            {
                ZipHelper.MultiZip(item.Key, item.Value.ToArray());
            }
        }
    }
    public class CSVGzipEntity {
        public string hanghao { get; set; }
        public string[] wenjianjia { get; set; }

    }
}
