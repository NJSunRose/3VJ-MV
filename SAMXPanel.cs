using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _3VJ_MV
{
    public class SAMXPanel
    {
        public string Index { get; set; }
        public string ColorSku { get; set; }
        public string Name { get; set; }
        public string CuttingLength { get; set; }
        public string CuttingWidth { get; set; }
        public string CuttingThickness { get; set; }
        public string CuttingNum { get; set; }
        public string EageNum { get; set; }
        public string Face6FileName { get; set; }
        public string Face5FileName { get; set; }
        public string YiXing { get; set; }
        public string Length { get; set; }
        public string Width { get; set; }
        public string Thickness { get; set; }
        public string Num { get; set; }
        public string Area { get; set; }
        public string PackNo { get; set; }
        public string CoderNo { get; set; }
        public string CabinetNo { get; set; }
        public string HoleNum { get; set; }
        public string Material { get; set; }
        public string EW1{ get; set; }
        public string EW2 { get; set; }
        public string EL1 { get; set; }
        public string EL2 { get; set; }
        public string DrawerNo { get; set; }

        public int GetEdgeNum(string edge)
        {
            if (edge.Trim() != string.Empty)
                return 1;
            else
                return 0;
        }

        public double GetArea(string length,string width)
        {
            return (Convert.ToDouble(length) * Convert.ToDouble(width)/1000000);
        }

    }
}
