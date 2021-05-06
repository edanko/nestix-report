using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Reflection;
using System.Windows;

namespace NestixReport
{
    partial class MainWindow
    {
        private void CalcGasesClick(object sender, RoutedEventArgs e)
		{
            var con = new SqlConnection(GetConnectionString());
            var com = new SqlCommand(Db.ForGas, con);
            com.Parameters.AddWithValue("name", GetFilter());

            try
            {
                con.Open();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            var reader = com.ExecuteReader();

            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };

            var xl = excelApp.Workbooks.Add(Missing.Value);
            var s = (Microsoft.Office.Interop.Excel.Worksheet)xl.Worksheets[1];

            s.Range["A1"].Value2 = "NC";
            s.Range["B1"].Value2 = "Technology";
            s.Range["C1"].Value2 = "Weight";
            s.Range["D1"].Value2 = "Кислород, кг";
            s.Range["E1"].Value2 = "Пропан, кг";
            s.Range["F1"].Value2 = "Азот, кг";
            s.Range["G1"].Value2 = "CO2 5% He 60% N2 35%, м3";

            s.Range["C2", "G2"].EntireColumn.NumberFormat = "0.000";
            //s.Range["C2", "G2"].EntireColumn.NumberFormat = "@";


			s.Range["A1", "G1"].Font.Bold = true;
			s.Range["A1", "G1"].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            var allNc = new List<string[]>();

			while (reader.Read())
			{
                var gas = GasAmount.GetGas(reader.GetInt32(3), reader.GetString(4), reader.GetDouble(2));

                var curNc = new[]
                {
                    reader.GetString(0),
                    reader.GetString(1),
                    reader.GetDouble(2).ToString(CultureInfo.InvariantCulture),
                    gas.Oxygen.ToString(CultureInfo.InvariantCulture),
                    gas.Propan.ToString(CultureInfo.InvariantCulture),
                    gas.Nitrogen.ToString(CultureInfo.InvariantCulture),
                    gas.LaserMix.ToString(CultureInfo.InvariantCulture)
                };
                allNc.Add(curNc);
            }

            var row = 2;
            foreach (var nc in allNc)
            {
                s.Range["A" + row].Value2 = nc[0];
                s.Range["B" + row].Value2 = nc[1];
                s.Range["C" + row].Value2 = nc[2];
                s.Range["D" + row].Value2 = nc[3];
                s.Range["E" + row].Value2 = nc[4];
                s.Range["F" + row].Value2 = nc[5];
                s.Range["G" + row].Value2 = nc[6];
                row++;
            }

            //s.Range["C2", "G" + row].NumberFormat = "0.00";

            s.Range["A1", "G1"].EntireColumn.AutoFit();

			reader.Close();
			con.Close();
			excelApp.Visible = true;
			excelApp.UserControl = true;
		}
    }

    public class GasAmount
    {
        public double Oxygen { get; private set; }
        public double Propan { get; private set; }
        public double Nitrogen { get; private set; }
        public double LaserMix { get; private set; }

        public static GasAmount GetGas(int machid, string quality, double weight)
        {
            const double plasmaOxygen = 5.56; // plasma - oxygen
            const double gasOxygen = 6.31; // gas - oxygen
            const int gasPropan = 2; // gas - propan
            const double laserNitrogenMetal = 0.16365879037204; // laser, n2, metal
            const double laserMixMetal = 0.000502461198510648; // laser - mix per kg, metal
            const double laserNitrogenPlywood = 22.8; // laser - n2, kg per list, plywood
            const double laserMixPlywood = 0.07; // laser - mix, m3, plywood

            var g = new GasAmount();

            switch (machid)
            {
                case 5:
                    g.Oxygen = plasmaOxygen * weight / 1000.0;
                    break;
                case 7:
                case 9:
                {
                    g.Oxygen = gasOxygen * weight / 1000.0;
                    g.Propan = gasPropan * weight / 1000.0;
                    break;
                }
                case 8:
                    if (quality == "PLYWOOD")
                    {
                        g.Nitrogen = laserNitrogenPlywood * 1.0;
                        g.LaserMix = laserMixPlywood * 1.0;
                    }
                    else
                    {
                        g.Nitrogen = laserNitrogenMetal * weight;
                        g.LaserMix = laserMixMetal * weight;
                    }
                    break;
            }
            return g;
        }
    }
}
