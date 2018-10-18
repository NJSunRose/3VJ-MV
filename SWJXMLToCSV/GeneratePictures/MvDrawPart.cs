using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWJXMLToCSV.GeneratePictures
{
    public class MvDrawPart
    {
        private static string filename;
        private static Part part;
        private static Bitmap buffer;
        private static string outfile;

        public static void DrawPartBitmap(string csvFilePath)
        {
            outfile = csvFilePath;
            if (!File.Exists(csvFilePath))
            {
                return;
            }

            FileInfo fi = new FileInfo(csvFilePath);
            filename = fi.Name.Substring(0, fi.Name.LastIndexOf("."));

            part = new Part();
            using (FileStream fs = new FileStream(csvFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (StreamReader sr = new StreamReader(fs, Encoding.Default))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    readSeq(line, sr, part);
                }
            }
            drawPart(part);
        }

        private static void readSeq(string line, StreamReader sr, Part part)
        {
            if (line.StartsWith("BorderSequence"))
            {
                part.Border = BorderSeq.LoadSeq(line);
            }
            else if (line.StartsWith("VdrillSequence"))
            {
                part.Vdrillings.Add(VdrillSeq.LoadSeq(line));
            }
            else if (line.StartsWith("HDrillSequence"))
            {
                part.Hdrillings.Add(HdrillSeq.LoadSeq(line));
            }
            else if (line.StartsWith("RouteSetMillSequence"))
            {
                RouteSeq rs = new RouteSeq(line);
                while (((line = sr.ReadLine()) != null)
                    && (line.StartsWith("RouteSequence")))
                {
                    rs.AddRoute(line);
                }
                part.Routes.Add(rs);

                if (line != null)
                { readSeq(line, sr, part); }
            }
            else if (line.StartsWith("EndSequence"))
            {
                return;
            }
        }
        //private static void getEdgeBanding(Part part, out string eU, out string eB, out string eL, out string eR)
        //{
        //    if (part.Border.MachinePoint == "1" || part.Border.MachinePoint == "7M")
        //    {
        //        eU = part.Border.Edge4;
        //        eB = part.Border.Edge3;
        //        eL = part.Border.Edge1;
        //        eR = part.Border.Edge2;
        //        return;
        //    }
        //    if (part.Border.MachinePoint == "4" || part.Border.MachinePoint == "6M")
        //    {
        //        eU = part.Border.Edge3;
        //        eB = part.Border.Edge4;
        //        eL = part.Border.Edge1;
        //        eR = part.Border.Edge2;
        //        return;
        //    }
        //    if (part.Border.MachinePoint == "5" || part.Border.MachinePoint == "3M")
        //    {
        //        eU = part.Border.Edge3;
        //        eB = part.Border.Edge4;
        //        eL = part.Border.Edge2;
        //        eR = part.Border.Edge1;
        //        return;
        //    }
        //    if (part.Border.MachinePoint == "8" || part.Border.MachinePoint == "2M")
        //    {
        //        eU = part.Border.Edge4;
        //        eB = part.Border.Edge3;
        //        eL = part.Border.Edge2;
        //        eR = part.Border.Edge1;
        //        return;
        //    }
        //    if (part.Border.MachinePoint == "2" || part.Border.MachinePoint == "4M")
        //    {
        //        eU = part.Border.Edge1;
        //        eB = part.Border.Edge2;
        //        eL = part.Border.Edge4;
        //        eR = part.Border.Edge3;
        //        return;
        //    }
        //    if (part.Border.MachinePoint == "6" || part.Border.MachinePoint == "8M")
        //    {
        //        eU = part.Border.Edge2;
        //        eB = part.Border.Edge1;
        //        eL = part.Border.Edge3;
        //        eR = part.Border.Edge4;
        //        return;
        //    }
        //    if (part.Border.MachinePoint == "3" || part.Border.MachinePoint == "1M")
        //    {
        //        eU = part.Border.Edge1;
        //        eB = part.Border.Edge2;
        //        eL = part.Border.Edge3;
        //        eR = part.Border.Edge4;
        //        return;
        //    }
        //    if (part.Border.MachinePoint == "7" || part.Border.MachinePoint == "5M")
        //    {
        //        eU = part.Border.Edge2;
        //        eB = part.Border.Edge1;
        //        eL = part.Border.Edge4;
        //        eR = part.Border.Edge3;
        //        return;
        //    }

        //    throw new Exception("Unknown machine point!");
        //}

        private static void drawPart(Part part)
        {

            buffer = new Bitmap(1980, 1300);
            using (Graphics g = Graphics.FromImage(buffer))
            {
                g.FillRectangle(new SolidBrush(Color.White), 0, 0, 1980, 1300);

                Drawer drawer = new Drawer(g, part, 1980, 1300);
                drawer.Draw();
            }
            if (File.Exists(Path.ChangeExtension(outfile, ".jpg")))
                File.Delete(Path.ChangeExtension(outfile, ".jpg"));
            buffer.Save(Path.ChangeExtension(outfile, ".jpg"));
        }
    }
}
