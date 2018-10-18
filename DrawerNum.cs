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
using Ionic.Zip;
using Dimeng.FTP;
using SpreadsheetGear;


namespace _3VJ_MV
{
    class DrawerNum
    {
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
                inifile.IniWriteValue("CsvNum", "Num", oldcsvnum.Substring(0, 1) + (int.Parse(oldcsvnum.Substring(1, 4)) + 1).ToString().PadLeft(4, '0'));
            }
            else
            {
                MessageBox.Show("记录订单号的配置文件不存在，请手动在 " + inipath + " 目录下创建！");
                return;
            }
    }
}