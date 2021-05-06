using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Windows;

namespace NestixReport
{
    partial class MainWindow
    {
        private void PickingListClick(object sender, RoutedEventArgs e)
        {
            var con = new SqlConnection(GetConnectionString());
            var ncCom = new SqlCommand(Db.GetNxPathIds, con);
            ncCom.Parameters.AddWithValue("name", GetFilter());

            try
            {
                con.Open();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            var parts = new List<string[]>();

            var ncReader = ncCom.ExecuteReader();

            var ncWithPathId = new Dictionary<string, int>();

            while (ncReader.Read())
            {
                ncWithPathId.Add((string) ncReader["name"], (int) ncReader["id"]);
            }

            ncReader.Close();

            foreach (var k in ncWithPathId.Keys)
            {
                var partsCom = new SqlCommand(Db.PickingList, con);
                partsCom.Parameters.AddWithValue("pathid", ncWithPathId[k]);
                var partsReader = partsCom.ExecuteReader();

                while (partsReader.Read())
                {
                    var a = new[]
                    {
                        k, // ncname
                        partsReader.GetString(0), // order
                        partsReader.GetString(1), // section
                        partsReader.GetString(2), // pos 
                        partsReader.GetInt32(3).ToString(), // det count
                        partsReader.GetFloat(4).ToString(CultureInfo.InvariantCulture), // weight
                        partsReader.GetFloat(5).ToString(CultureInfo.InvariantCulture), // total weight
                        partsReader.GetFloat(6).ToString(CultureInfo.InvariantCulture), // thick
                        partsReader.GetString(7), // quality
                        partsReader.GetFloat(8).ToString(CultureInfo.InvariantCulture), // partlen
                        partsReader.GetFloat(9).ToString(CultureInfo.InvariantCulture) // part width
                    };
                    parts.Add(a);
                }

                partsReader.Close();
            }

            con.Close();

            ncWithPathId.Clear();

            if (parts.Count == 0)
            {
                MessageBox.Show("Карты не найдены");
                return;
            }

            try
            {
                parts.Sort((x, y) => int.Parse(x[3]).CompareTo(int.Parse(y[3])));
            }
            catch { }


            var tmp = new List<string[]> {parts[0]};

            for (var i = 1; i < parts.Count; i++)
            {
                if (parts[i][3] == parts[i - 1][3]) continue;

                if (parts[i][2] == parts[i - 1][2] && parts[i][1] == parts[i - 1][1])
                {
                    tmp.Add(parts[i]);
                }
            }


            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };

            var xl = excelApp.Workbooks.Add();
            var s = (Microsoft.Office.Interop.Excel.Worksheet) xl.Worksheets[1];

            s.Range["A1"].Value2 = "№ п/п";
            s.Range["B1"].Value2 = "№ чертежа";
            s.Range["C1"].Value2 = "Заказ";
            s.Range["D1"].Value2 = "Секция";
            s.Range["E1"].Value2 = "№ поз.";
            s.Range["F1"].Value2 = "Кол-во";
            s.Range["G1"].Value2 = "Толщина, мм";
            s.Range["H1"].Value2 = "Марка материала";
            s.Range["I1"].Value2 = "Карта раскроя";
            s.Range["J1"].Value2 = "Тип";
            s.Range["K1"].Value2 = "Масса 1 шт., кг";
            s.Range["L1"].Value2 = "Общ. масса, кг";
            s.Range["M1"].Value2 = "Длина, мм";
            s.Range["N1"].Value2 = "Ширина, мм";
            s.Range["O1"].Value2 = "Узел";
            s.Range["P1"].Value2 = "Маршрут";
            s.Range["Q1"].Value2 = "Примечание";




            s.Range["A1", "Q1"].Font.Bold = true;
            s.Range["A1", "Q1"].EntireColumn.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            s.Range["A1", "Q1"].EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            s.Range["B1", "E1"].EntireColumn.NumberFormat = "@";
            s.Range["K1", "L1"].EntireColumn.NumberFormat = "0.00";
            s.Range["M1", "N1"].EntireColumn.NumberFormat = "0.0";

            var row = 2;

            foreach (var p in tmp)
            {
                s.Range["A" + row].Value2 = row - 1;
                s.Range["B" + row].Value2 = "";
                s.Range["C" + row].Value2 = p[1];
                s.Range["D" + row].Value2 = p[2];
                s.Range["E" + row].Value2 = p[3];
                s.Range["F" + row].Value2 = p[4];
                s.Range["G" + row].Value2 = p[7];
                s.Range["H" + row].Value2 = p[8];

                var nclist = parts.FindAll(x => x[3] == p[3] && x[2] == p[2] && x[1] == p[1]);
                var nx = nclist.Select(nc => nc[0]).ToList();
                nx.Sort();

                s.Range["I" + row].Value2 = string.Join("\n", nx.ToArray());

                s.Range["J" + row].Value2 = "Лист";
                s.Range["K" + row].Value2 = p[5];
                s.Range["L" + row].Value2 = p[6];
                s.Range["M" + row].Value2 = p[9];
                s.Range["N" + row].Value2 = p[10];
                s.Range["O" + row].Value2 = "";
                s.Range["P" + row].Value2 = "";
                s.Range["Q" + row].Value2 = "";

                row++;
            }

            s.Range["A1", "Q1"].EntireColumn.AutoFit();
            s.Range["A1", "Q" + row].EntireRow.AutoFit();
        }
    }
}
