using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Windows;
using Microsoft.Office.Interop.Excel;

namespace NestixReport
{
    partial class MainWindow
    {
        private void BatchNestInfo(object sender, RoutedEventArgs e)
		{
            var con = new SqlConnection(GetConnectionString());
            var com = new SqlCommand(Db.BatchNestingInfo, con);
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


            var row = 5;
            var allNc = new List<string[]>();

            while (reader.Read())
            {
                var curNc = new[]
                {
                    reader.GetString(0),
                    reader.GetString(1),
                    reader.GetString(2),
                    reader.GetFloat(3).ToString(CultureInfo.InvariantCulture),
                    reader.GetString(4),
                    reader.GetString(5),
                    reader.GetString(6),
                    reader.GetString(7)
                };
                allNc.Add(curNc);
            }

            reader.Close();
            con.Close();

            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };

            var xl = excelApp.Workbooks.Add();
            var s = (Worksheet)xl.Worksheets[1];

            // s.Range["A3"].Value2 = "№";
            // s.Range["B3"].Value2 = "УП";
            // s.Range["C3"].Value2 = "Заказ-Секция";
            // s.Range["D3"].Value2 = "МТР";
            // s.Range["E3"].Value2 = "Толщина";
            // s.Range["F3"].Value2 = "Марка";
            // s.Range["G3"].Value2 = "Габариты заготовки, мм";
            // s.Range["H3"].Value2 = "Инфо";
            // s.Range["I3"].Value2 = "Номера отходов";

            s.Range["A3"].Value2 = "№";
            s.Range["A4"].Value2 = "#";
            
            s.Range["B3"].Value2 = "Карта раскроя";
            s.Range["B4"].Value2 = "CP";

            s.Range["C3"].Value2 = "Тех.\nкомплект";
            s.Range["C4"].Value2 = "-";

            s.Range["D3"].Value2 = "Заказ";
            s.Range["D4"].Value2 = "Order";
            
            s.Range["E3"].Value2 = "Секция";
            s.Range["E4"].Value2 = "Section";
            
            s.Range["F3"].Value2 = "МТР";
            s.Range["F4"].Value2 = "Cutting Machine";
            
            s.Range["G3"].Value2 = "Толщина 1";
            s.Range["G4"].Value2 = "Thickness 1";
            
            s.Range["H3"].Value2 = "Толщина 2";
            s.Range["H4"].Value2 = "Thickness 2";
            
            s.Range["I3"].Value2 = "Ширина 1";
            s.Range["I4"].Value2 = "Width 1";
            
            s.Range["J3"].Value2 = "Ширина 2";
            s.Range["J4"].Value2 = "Width 2";
            
            s.Range["K3"].Value2 = "Длина";
            s.Range["K4"].Value2 = "Length";
            
            s.Range["L3"].Value2 = "Марка\nпроката";
            s.Range["L4"].Value2 = "Quality";
            
            s.Range["M3"].Value2 = "№ материала";
            s.Range["M4"].Value2 = "Mark No.";
            
            s.Range["N3"].Value2 = "Кол-во";
            s.Range["N4"].Value2 = "Quantity";
            

            switch (DbComboBox.Text)
            {
                case "NxSC_Zvezda_120K":
                case "NxSC_Zvezda_69K":
                    s.Range["O3"].Value2 = "№ SHI";
                    s.Range["O4"].Value2 = "CP SHI";
                    break;

                case "NxSC_Zvezda_AFRA":
                case "NxSC_Zvezda_MR":
                    s.Range["O3"].Value2 = "№ HSHI";
                    s.Range["O4"].Value2 = "CP HSHI";
                    break;

                default:
                    s.Range["O3"].Value2 = "Примечание";
                    s.Range["O4"].Value2 = "Notes";

                    break;
            }




            //s.Range["A3", "O4"].Font.Bold = true;
            s.Range["A3", "O4"].Interior.Color = Color.LightGray;
            
            s.Range["D1"].EntireColumn.NumberFormat = "@";
            s.Range["E1"].EntireColumn.NumberFormat = "@";

            foreach (var nc in allNc)
            {
                nc[2] = nc[2] switch
                {
                    "PlasmaBevelOmniMatL8000" => "OM8000 (Plasma)",
                    "GasBevelOmniMatL8000" => "OM8000 (Gas)",
                    "LaserMatL4200" => "LM4200",
                    "GasOmniMatL7000" => "OM7000 (Gas)",
                    _ => nc[2]
                };
            }

            var n = 1;
            foreach (var nc in allNc)
            {
                s.Range["A" + row].Value2 = n;
                s.Range["B" + row].Value2 = nc[0];
                s.Range["C" + row].Value2 = "-";
                s.Range["D" + row].Value2 = nc[1].Split("-")[0];
                s.Range["E" + row].Value2 = nc[1].Split("-")[1];
                s.Range["F" + row].Value2 = nc[2];
                s.Range["G" + row].Value2 = nc[3];
                s.Range["H" + row].Value2 = "";
                s.Range["I" + row].Value2 = nc[5].Split("x")[1];
                s.Range["J" + row].Value2 = "";
                s.Range["K" + row].Value2 = nc[5].Split("x")[0];
                s.Range["L" + row].Value2 = nc[4];

                if (nc[6].Contains(';'))
                {
                    var l = nc[6].Split(';');

                    s.Range["M" + row].Value2 = l[1];
                    s.Range["O" + row].Value2 = l[0];
                }
                else
                {
                    s.Range["M" + row].Value2 = nc[6];
                    s.Range["O" + row].Value2 = "";
                }

                s.Range["N" + row].Value2 = "1";

                n++;
                row++;
            }

            //s.Range["E3", "E" + row].NumberFormat = "0.00";

            row--;

            s.Range["A3", "O" + row].EntireColumn.VerticalAlignment = XlVAlign.xlVAlignCenter;
            s.Range["A3", "O" + row].EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            s.Range["A3", "O" + row].Borders.LineStyle = XlLineStyle.xlContinuous;
            s.Range["A3", "O" + row].BorderAround(XlLineStyle.xlContinuous);

            s.Range["D1"].Value2 = $"Заказ {s.Range["D5"].Value2} секция {s.Range["E5"].Value2}";
            s.Range["D1", "I2"].Merge();
            s.Range["D1", "I2"].Font.Size = 24;

            s.Range["A1", "I1"].EntireColumn.AutoFit();
            s.Range["A1", "I" + row].EntireRow.AutoFit();

            excelApp.Visible = true;
            excelApp.UserControl = true;
        }
    }
}
