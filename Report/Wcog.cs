using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows;

namespace NestixReport
{
    public class Wcog
    {
        public int PosNo { get; set; }
        public int Quantity { get; set; }
        public double Weight { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public string Panel { get; set; }
        public string Block { get; set; }
        public string Part { get; set; }
        public string Type { get; set; }
        public string Side { get; set; }
        public string StockNumber { get; set; }
        public string Quality { get; set; }
        public string Gps1 { get; set; }
        public string Gps2 { get; set; }
        public string Gps3 { get; set; }
        public string Gps4 { get; set; }
        public string Ship { get; set; }
        public string Ident { get; set; }
        public string NestedOn { get; set; }
        public double Area { get; set; }
        public double CircLength { get; set; }
        public double CircWidth { get; set; }
        public double Thickness { get; set; }
        public string Shape { get; set; }
        public string Dimension { get; set; }
        public double TotalLength { get; set; }
        public double MouldedLength { get; set; }

        /*public override string ToString()
        {
            return $"PosNo: {PosNo}, Block: {Block}, T: {GetThickness()}, Qual: {Quality}, Qty: {Quantity}";
        }*/

        public double GetThickness()
        {
            if (Dimension.Length > 0)
            {
                var spl = Dimension.Split('*');
                return double.Parse(spl.Length > 1 ? spl[1] : Dimension, CultureInfo.InvariantCulture);
            }

            return Thickness;
        }

        public static Dictionary<int, Wcog> Read(string file)
        {
            var wcog = new Dictionary<int, Wcog>();

            var wcogFileLines = File.ReadAllLines(file);

            for (var i = 2; i < wcogFileLines.Length; i++)
            {
                var c = wcogFileLines[i].Split(',');

                if (c.Length < 27)
                {
                    continue;
                }

                var pos = GetPos(c[0]);

                if (wcog.ContainsKey(pos))
                {
                    wcog[pos].Quantity++;
                    continue;
                }

                var l = new Wcog();
                l.Panel = c[5];
                l.Block = c[6];
                l.Part = c[7];
                l.Type = c[8];
                l.Side = c[9];
                l.StockNumber = c[10];
                l.Quality = c[11].Replace("_B", "");
                l.Gps1 = c[12];
                l.Gps2 = c[13];
                l.Gps3 = c[14];
                l.Gps4 = c[15];
                l.Ship = c[16];
                l.Ident = c[17];
                l.NestedOn = c[18];
                l.Shape = c[23];
                l.Dimension = c[24];
                l.MouldedLength = double.TryParse(c[26], NumberStyles.Any, CultureInfo.InvariantCulture, out var val1) ? val1 : 0.0;
                l.TotalLength = double.TryParse(c[25],  NumberStyles.Any, CultureInfo.InvariantCulture, out var val2) ? val2 : 0.0;
                l.Thickness = double.TryParse(c[22], NumberStyles.Any, CultureInfo.InvariantCulture, out var val3) ? val3 : 0.0;
                l.CircWidth = double.TryParse(c[21], NumberStyles.Any, CultureInfo.InvariantCulture, out var val4) ? val4 : 0.0;
                l.CircLength = double.TryParse(c[20], NumberStyles.Any, CultureInfo.InvariantCulture, out var val5) ? val5 : 0.0;
                l.Area = double.TryParse(c[19],  NumberStyles.Any, CultureInfo.InvariantCulture, out var val6) ? val6 : 0.0;
                l.Z = double.TryParse(c[4],  NumberStyles.Any, CultureInfo.InvariantCulture, out var val7) ? val7 : 0.0;
                l.Y = double.TryParse(c[3],  NumberStyles.Any, CultureInfo.InvariantCulture, out var val8) ? val8 : 0.0;
                l.X = double.TryParse(c[2],  NumberStyles.Any, CultureInfo.InvariantCulture, out var val9) ? val9 : 0.0;
                l.Weight = double.TryParse(c[1],  NumberStyles.Any, CultureInfo.InvariantCulture, out var val10) ? val10 : 0.0;
                l.Quantity = 1;
                l.PosNo = pos;

                wcog[pos] = l;
            }

            //wcog.Sort((x, y) => x.PosNo.CompareTo(y.PosNo));

            return wcog;
        }

        private static int GetPos(string s)
        {
            try
            {
                return int.Parse(s.Split('-')[^1].Replace("P", "").Replace("S", "").Replace("B", "").Replace("C", ""));
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error in GetPos(str to int convertion)");
                return -1;
            }
        }
    }
}
