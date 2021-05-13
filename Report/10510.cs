using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using Application = Microsoft.Office.Interop.Excel.Application;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace NestixReport
{
    partial class MainWindow
    {
        private void PickingList10510Click(object sender, RoutedEventArgs e)
        {
            if (DbComboBox.SelectedIndex != 0)
            {
                MessageBox.Show("Не выбран проект 10510!", "Фатал еррор");
                return;
            }

            var ofd = new CommonOpenFileDialog
            {
                IsFolderPicker = true
            };

            if (ofd.ShowDialog() != CommonFileDialogResult.Ok)
            {
                return;
            }

            var wcogPath = "";
            var docPath = "";

            var files = Directory.GetFiles(ofd.FileName);

            foreach (var file in files)
            {
                var lower = file.ToLowerInvariant();

                if (lower.Contains("wcog") && Path.GetExtension(lower) == ".csv")
                {
                    wcogPath = lower;
                }
                else if (Path.GetExtension(lower) == ".docx" && !lower.Contains("~"))
                {
                    docPath = lower;
                }
            }

            if (string.IsNullOrEmpty(wcogPath))
            {
                MessageBox.Show("Файл wcog не обнаружен");
                return;
            }

            var docx = Docx.Read(docPath);

            var nxparts = NestixPartlist.GetPartlistFromNestix(GetConnectionString(), GetFilter());

            var wcog = Wcog.Read(wcogPath);

            var wcogKeys = docx.Keys.ToList();
            wcogKeys.Sort();

            var excel = new Application
            {
                Visible = true
            };

            var xl = excel.Workbooks.Add();
            var s = (Worksheet) xl.Worksheets[1];


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
            s.Range["A1", "Q1"].EntireColumn.VerticalAlignment = XlVAlign.xlVAlignCenter;
            s.Range["A1", "Q1"].EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            s.Range["B1", "E1"].EntireColumn.NumberFormat = "@";
            s.Range["K1", "L1"].EntireColumn.NumberFormat = "0.00";
            s.Range["M1", "N1"].EntireColumn.NumberFormat = "0.0";

            var row = 2;
            foreach (var p in wcogKeys)
            {

                var nclist = new List<NestixPartlist>();

                int posno = 0;
                string block = "";
                int quantity = 0;
                double thickness = 0.0;
                string quality = "";
                string dimension = "";
                double weight = 0.0, totalWeight = 0.0;
                double length = 0.0, width = 0.0;
                
                if (wcog.ContainsKey(p))
                {
                    posno = wcog[p].PosNo;
                    block = wcog[p].Block;
                    quantity = wcog[p].Quantity;
                    thickness = wcog[p].GetThickness();
                    quality = wcog[p].Quality;
                    dimension = wcog[p].Shape.TrimEnd() switch
                    {
                        "FB" => $"Полоса {wcog[p].Dimension.Replace("*", "x").Replace(".0", "")}",
                        "PP" => $"Полособульб {wcog[p].Dimension.Replace("*", "x").Replace(".0", "")}",
                        _ => "Лист"
                    };
                    weight = wcog[p].Weight;
                    totalWeight = weight * quantity;
                }
                else
                {
                    posno = docx[p].PosNo;
                    block = docx[p].Block;
                    quantity = docx[p].Quantity;
                    thickness = docx[p].GetThickness();
                    quality = docx[p].Quality;
                    dimension = docx[p].Dimension;
                    weight = docx[p].Weight;
                    totalWeight = weight * quantity;
                }

                if (nxparts != null)
                {
                    nclist = nxparts.FindAll(x =>
                            x.PosNo == posno.ToString() && x.Section == block);
                }

                if (nclist.Count > 0)
                {
                    weight = nclist[0].Weight;
                    totalWeight = nclist[0].TotalWeight;
                }
                
                if (dimension == "Лист" || dimension.StartsWith("Полоса"))
                {
                    if (nclist.Count > 0)
                    {
                        length = nclist[0].Length;
                        width = nclist[0].Width;
                    }
                }
                else
                {
                    if (wcog.ContainsKey(p))
                    {
                        length = wcog[p].TotalLength;

                        width = double.Parse(wcog[p].Shape.TrimEnd() switch
                        {
                            "FB" => wcog[p].Dimension.Split("*")[0],
                            "PP" => wcog[p].Dimension.Split("*")[0],
                            _ => ""
                        }, CultureInfo.InvariantCulture);
                    }
                }

                s.Range["A" + row].Value2 = row - 1;
                s.Range["B" + row].Value2 = "";
                s.Range["C" + row].Value2 = "056001";
                s.Range["D" + row].Value2 = block;
                s.Range["E" + row].Value2 = posno;
                s.Range["F" + row].Value2 = nclist.Count > 0 ? nclist[0].Count : quantity;
                s.Range["G" + row].Value2 = thickness;
                s.Range["H" + row].Value2 = quality;

                var nx = nclist.Select(nc => nc.NcName).ToList();
                nx.Sort();
                s.Range["I" + row].Value2 = string.Join("\n", nx.ToArray());
                s.Range["J" + row].Value2 = dimension;
                s.Range["K" + row].Value2 = weight;
                s.Range["L" + row].Value2 = totalWeight;
                s.Range["M" + row].Value2 = length;
                s.Range["N" + row].Value2 = width;
                s.Range["O" + row].Value2 = docx[posno].Assembly;
                s.Range["P" + row].Value2 = "";
                s.Range["Q" + row].Value2 = "";

                row++;
            }

            s.Range["A1", "Q1"].EntireColumn.AutoFit();
            s.Range["A1", "Q" + row].EntireRow.AutoFit();


            var dir = Path.GetDirectoryName(docPath);
            var fname = Path.GetFileNameWithoutExtension(docPath);
            fname += " - Комплектовочная ведомость.xlsx";
            fname = Path.Join(dir, fname);

            xl.SaveAs(fname);
        }

        private void CheckWcogClick(object sender, RoutedEventArgs e)
        {
            var ofd = new CommonOpenFileDialog
            {
                IsFolderPicker = true
            };

            if (ofd.ShowDialog() != CommonFileDialogResult.Ok)
            {
                return;
            }

            var wcogPath = "";
            var docPath = "";

            var files = Directory.GetFiles(ofd.FileName);

            foreach (var file in files)
            {
                var lower = file.ToLowerInvariant();

                if (lower.Contains("wcog") && Path.GetExtension(lower) == ".csv")
                {
                    wcogPath = lower;
                }
                else if (Path.GetExtension(lower) == ".docx" && !lower.Contains("~"))
                {
                    docPath = lower;
                }
            }

            if (string.IsNullOrEmpty(wcogPath))
            {
                MessageBox.Show("Файл wcog не обнаружен");
                return;
            }

            var docx = Docx.Read(docPath);

            var docxKeys = docx.Keys.ToList();
            //docxKeys.Sort();

            var wcog = Wcog.Read(wcogPath);

            var wcogKeys = wcog.Keys.ToList();
            //wcogKeys.Sort();

            var log = new List<string[]>();

            var keysOnlyInDocx = docxKeys.Except(wcogKeys).ToList();
            keysOnlyInDocx.Sort();


            foreach (var k in keysOnlyInDocx)
            {
                log.Add(new[]
                {
                    docx[k].PosNo.ToString(CultureInfo.InvariantCulture),
                    "отсутствует в WCOG",
                    "-",
                    "-"
                });
            }

            var keysOnlyInWcog = wcogKeys.Except(docxKeys).ToList();
            keysOnlyInWcog.Sort();

            foreach (var k in keysOnlyInWcog)
            {
                log.Add(new[]
                {
                    wcog[k].PosNo.ToString(CultureInfo.InvariantCulture),
                    "отсутствует в спецификации",
                    "-",
                    "-"
                });
            }

            var common = wcogKeys.Intersect(docxKeys).ToList();
            common.Sort();

            foreach (var k in common)
            {
                if (wcog[k].Quantity != docx[k].Quantity)
                {
                    log.Add(new[]
                    {
                        wcog[k].PosNo.ToString(CultureInfo.InvariantCulture),
                        "конфликт кол-ва деталей",
                        wcog[k].Quantity.ToString(),
                        docx[k].Quantity.ToString()
                    });
                }

                if (wcog[k].Quality != docx[k].Quality)
                {
                    log.Add(new[]
                    {
                        wcog[k].PosNo.ToString(CultureInfo.InvariantCulture),
                        "конфликт материалов",
                        wcog[k].Quality,
                        docx[k].Quality
                    });
                }

                var plWeight = Math.Round(wcog[k].Weight, 1);
                var docWeight = Math.Round(docx[k].Weight, 1);

                if (Math.Abs(plWeight - docWeight) > 0.15)
                {
                    log.Add(new[]
                    {
                        wcog[k].PosNo.ToString(CultureInfo.InvariantCulture),
                        "конфликт массы",
                        plWeight.ToString("G"),
                        docWeight.ToString("G")
                    });
                }



                try
                {
                    var wcogThickness = wcog[k].GetThickness();
                    var docThickness = docx[k].GetThickness();

                    if (wcogThickness.CompareTo(docThickness) != 0)
                    {
                        log.Add(new[]
                        {
                            wcog[k].PosNo.ToString(CultureInfo.InvariantCulture),
                            "конфликт толщин",
                            wcogThickness.ToString("G"),
                            docThickness.ToString("G")
                        });
                    }
                }
                catch
                {
                    log.Add(new[]
                    {
                        wcog[k].PosNo.ToString(CultureInfo.InvariantCulture),
                        "в спецификации нет толщины",
                        "-",
                        "-"
                    });
                }
            }

            // try
            // {
            //     log.Sort((x, y) => int.Parse(x[0]).CompareTo(int.Parse(y[0])));
            // }
            // catch { }



            /*var spreadsheetDocument = SpreadsheetDocument.Create(@"e:\out.xlsx", SpreadsheetDocumentType.Workbook);

            var workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            var sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(workbookpart),
                SheetId = 1,
                Name = "sheetName"
            };

            sheets.AppendChild(sheet);

            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new SheetData());
            
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
            
            
            
            #region open pdf
            var psi = new ProcessStartInfo(resultFilename)
            {
                UseShellExecute = true
            };
            Process.Start(psi);
            #endregion
            */
            
            
            
            

            var excel = new Application
            {
                Visible = true
            };

            var xl = excel.Workbooks.Add();
            var s = (Worksheet) xl.Worksheets[1];

            s.Range["A1"].Value2 = "№ поз.";
            s.Range["B1"].Value2 = "Тип";
            s.Range["C1"].Value2 = "WCOG";
            s.Range["D1"].Value2 = "Спец.";

            s.Range["A1", "D1"].Font.Bold = true;
            s.Range["A1", "D1"].EntireColumn.VerticalAlignment = XlVAlign.xlVAlignCenter;
            s.Range["A1", "D1"].EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            var row = 2;
            foreach (var l in log)
            {
                s.Range["A" + row].Value2 = l[0];
                s.Range["B" + row].Value2 = l[1];
                s.Range["C" + row].Value2 = l[2];
                s.Range["D" + row].Value2 = l[3];

                row++;
            }

            //s.Range["A1", "D" + row].Borders.Weight = 2;
            s.Range["A1", "D1"].EntireColumn.AutoFit();
            s.Range["A1", "D" + row].EntireRow.AutoFit();

            s.Range["A1", "B1"].AutoFilter(1, null, XlAutoFilterOperator.xlFilterNoFill, null, true);

            var dir = Path.GetDirectoryName(docPath);
            var fname = Path.GetFileNameWithoutExtension(docPath);
            fname += " - Лог сравнения WCOG и спецификации.xlsx";
            fname = Path.Join(dir, fname);

            xl.SaveAs(fname);
        }

        private void QuantityCheckWcogAndNestix(object sender, RoutedEventArgs e)
        {
            if (DbComboBox.SelectedIndex != 0)
            {
                MessageBox.Show("Не выбран проект 10510!", "Фатал еррор");
                return;
            }

            var ofd = new CommonOpenFileDialog
            {
                IsFolderPicker = false
            };

            if (ofd.ShowDialog() != CommonFileDialogResult.Ok)
            {
                return;
            }

            var nxparts = NestixPartlist.GetPartlistFromNestix(GetConnectionString(), GetFilter());
            var parts = Wcog.Read(ofd.FileName);

            var excel = new Application
            {
                Visible = true
            };
            var xl = excel.Workbooks.Add();
            var s = (Worksheet)xl.Worksheets[1];

            s.Range["A1"].Value2 = "Секция";
            s.Range["B1"].Value2 = "№ поз.";
            s.Range["C1"].Value2 = "WCOG";
            s.Range["D1"].Value2 = "NESTIX";

            s.Range["A1", "D1"].Font.Bold = true;

            var row = 2;
            foreach (var n in nxparts)
            {

                var pos = int.Parse(n.PosNo);
                
                if (parts.ContainsKey(pos))
                {
                    if (parts[pos].Block != n.Section)
                    {
                        continue;
                    }

                    if (parts[pos].Quantity == n.Count)
                    {
                        continue;
                    }
                    
                    s.Range["A" + row].Value2 = parts[pos].Block;
                    s.Range["B" + row].Value2 = parts[pos].PosNo;
                    s.Range["C" + row].Value2 = parts[pos].Quantity;
                    s.Range["D" + row].Value2 = n.Count;

                    row++;
                }
            }

            s.Range["A1", "D1"].EntireColumn.AutoFit();
            s.Range["A1", "D" + row].EntireRow.AutoFit();
        }
    }
}
