using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows;

namespace NestixReport
{
    public class MaterialPivot
    {
        public float Thickness { get; set; }
        public string Quality { get; set; }
        public double Length { get; set; }
        public double Width { get; set; }
        public int Quantity { get; set; }
    }
    
    public class Nc
    {
        public float Thickness { get; set; }
        public string Quality { get; set; }
        public double Length { get; set; }
        public double Width { get; set; }
    }
    
    
    partial class MainWindow
    {
        
        private void PlatePivotClick(object sender, RoutedEventArgs e)
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

            var allNc = new List<Nc>();

            while (reader.Read())
            {
                var curNc = new Nc();
                curNc.Length = reader.GetDouble(6);
                curNc.Width = reader.GetDouble(7);
                curNc.Quality = reader.GetString(5);
                curNc.Thickness = reader.GetFloat(4);
                
                allNc.Add(curNc);
            }

            reader.Close();
            con.Close();

            var list = new List<MaterialPivot>();

            foreach (var nc in allNc)
            {
                var exist = list.Find(x =>
                    Math.Abs(x.Thickness - nc.Thickness) < 0.0001 && x.Quality == nc.Quality && Math.Abs(x.Length - nc.Length) < 0.0001 &&
                    Math.Abs(x.Width - nc.Width) < 0.0001);

                if (exist != null)
                {
                    exist.Quantity++;
                }
                else
                {
                    var i = new MaterialPivot();
                    i.Length = nc.Length;
                    i.Width = nc.Width;
                    i.Quality = nc.Quality;
                    i.Thickness = nc.Thickness;
                    i.Quantity = 1;
                    
                    list.Add(i);
                }
            }

            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };

            var xl = excelApp.Workbooks.Add();
            var s = (Microsoft.Office.Interop.Excel.Worksheet)xl.Worksheets[1];

            s.Range["A1"].Value2 = "№";
            s.Range["B1"].Value2 = "Толщина";
            s.Range["C1"].Value2 = "Материал";
            s.Range["D1"].Value2 = "Длина";
            s.Range["E1"].Value2 = "Ширина";
            s.Range["F1"].Value2 = "Кол-во";

            s.Range["A1", "F1"].Font.Bold = true;
            s.Range["A1", "F1"].EntireColumn.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            s.Range["A1", "F1"].EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            s.Range["B1"].EntireColumn.NumberFormat = "0.0";

            var n = 1;
            var row = 2;
 
            foreach (var line in list)
            {
                s.Range["A" + row].Value2 = n;
                s.Range["B" + row].Value2 = line.Thickness;
                s.Range["C" + row].Value2 = line.Quality;
                s.Range["D" + row].Value2 = line.Length;
                s.Range["E" + row].Value2 = line.Width;
                s.Range["F" + row].Value2 = line.Quantity;

                n++;
                row++;
            }

            s.Range["A1", "F1"].EntireColumn.AutoFit();

            excelApp.Visible = true;
            excelApp.UserControl = true;
        }
    }
}
