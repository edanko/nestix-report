using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Reflection;
using System.Windows;

namespace NestixReport
{
    partial class MainWindow
    {
        private void MaterialReport(object sender, RoutedEventArgs e)
        {
            var con = new SqlConnection(GetConnectionString());
            var com = new SqlCommand(Db.PlatesInfo, con);
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

            var allNc = new List<string[]>();

            while (reader.Read())
            {
                var curNc = new[]
                {
                    reader.GetString(0), // ncname
                    reader.GetString(1), // machine
                    reader.GetFloat(2).ToString(CultureInfo.InvariantCulture), // used
                    reader.GetFloat(3).ToString(CultureInfo.InvariantCulture), // machtime
                    reader.GetFloat(4).ToString(CultureInfo.InvariantCulture), // thick
                    reader.GetString(5), // quality
                    reader.GetDouble(6).ToString(CultureInfo.InvariantCulture), // len
                    reader.GetDouble(7).ToString(CultureInfo.InvariantCulture), // width
                    reader.GetFloat(8).ToString(CultureInfo.InvariantCulture) // matweight
                };
                allNc.Add(curNc);
            }

            reader.Close();
            con.Close();


            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };

            var xl = excelApp.Workbooks.Add(Missing.Value);
            var s = (Microsoft.Office.Interop.Excel.Worksheet)xl.Worksheets[1];

            s.Range["A1"].Value2 = "№ п/п";
            s.Range["B1"].Value2 = "Карта раскроя";
            s.Range["C1"].Value2 = "Тип оборудования";
            s.Range["D1"].Value2 = "Коэф. ракроя, %";
            s.Range["E1"].Value2 = "Машинное время, мин";
            s.Range["F1"].Value2 = "Толщина заготовки, мм";
            s.Range["G1"].Value2 = "Марка материала";
            s.Range["H1"].Value2 = "Длина заготовки, мм";
            s.Range["I1"].Value2 = "Ширина заготовки, мм";
            s.Range["J1"].Value2 = "Масса материала, кг";

            s.Range["A1", "J1"].Font.Bold = true;
            s.Range["A1", "J1"].EntireColumn.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            s.Range["A1", "J1"].EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
           
            s.Range["C1"].EntireColumn.NumberFormat = "0.0";
            s.Range["D1", "E1"].EntireColumn.NumberFormat = "0.00";
            s.Range["J1"].EntireColumn.NumberFormat = "0.00";


            var row = 2;

            foreach (var nc in allNc)
            { 
                s.Range["A" + row].Value2 = row - 1;
                s.Range["B" + row].Value2 = nc[0];

                var tech = nc[1] switch
                {
                    "PlasmaBevelOmniMatL8000" => "OmniMat L8000 (Plasma)",
                    "GasBevelOmniMatL8000" => "OmniMat L8000 (Gas)",
                    "LaserMatL4200" => "LaserMat L4200",
                    "OmniMatL7000" => "OmniMat L7000 (Gas)",
                    _ => nc[1]
                };
                s.Range["C" + row].Value2 = tech;
                
                s.Range["D" + row].Value2 = nc[2];
                s.Range["E" + row].Value2 = nc[3];
                s.Range["F" + row].Value2 = nc[4];
                s.Range["G" + row].Value2 = nc[5];
                s.Range["H" + row].Value2 = nc[6];
                s.Range["I" + row].Value2 = nc[7];
                s.Range["J" + row].Value2 = nc[8];

                row++;
            }

            s.Range["A1", "J1"].EntireColumn.AutoFit();

            excelApp.Visible = true;
            excelApp.UserControl = true;
        }
    }
}
