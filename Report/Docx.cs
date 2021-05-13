using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DataTable = System.Data.DataTable;

namespace NestixReport
{
    public class Docx
    {
        public string Block { get; set; }
        public int PosNo { get; set; }
        public int Quantity { get; set; }
        public string Dimension { get; set; }
        public string Quality { get; set; }
        public double Weight { get; set; }
        public string Assembly { get; set; }

        public double GetThickness()
        {
            var spl = Dimension.Split('*');
            return double.Parse(spl.Length > 1 ? spl[1] : Dimension, CultureInfo.InvariantCulture);
        }
        
        public static Dictionary<int, Docx> Read(string filename)
        {
            const string sortament = "5,50*4.0,2.25	5.5,55*4.5,2.73	6,60*5.0,3.36	7,70*5.0,3.98	8,80*5.0,4.58	9,90*5.5,5.52	10,100*6.0,6.76	12,120*6.5,8.75	14а,140*7.0,11.05	14б,140*9.0,13.23	16а,160*8.0,14.08	16б,160*10.0,16.60	18а,180*9.0,17.41	18б,180*11.0,20.24	20а,200*10.0,21.47	20б,200*12.0,24.60	22а,220*11.0,25.75	22б,220*13.0,29.20	24а,240*12.0,30.42	24б,240*14.0,34.18	90*90*6.0,90*6.0,8.33	63*63*6.0,63*6.0,5.72	108*6.0,108*6.0,15.09	133*8.0,133*8.0,24.66	168*9.0,168*9.0,35.29";
            
            var dataTable = new DataTable("Sortament");
            dataTable.Columns.Add("NAME");
            dataTable.Columns.Add("DIM");
            var item = new DataColumn[2];
            item[0] = dataTable.Columns["NAME"];
            dataTable.PrimaryKey = item;
            var sortaments = sortament.Split("\t");
            foreach (var t in sortaments)
            {
                var sortamentsSplit = t.Split(',');
                dataTable.Rows.Add(sortamentsSplit[0], sortamentsSplit[1]);
            }
            
            
            
            var doc = WordprocessingDocument.Open(filename, false);
                
            var res = new Dictionary<int, Docx>();
                
            var assembly = "";

            var block = "";

            foreach (var table in doc.MainDocumentPart.Document.Body.Elements<Table>())
            {
                var columnsCount = table.Elements<TableGrid>().First().ChildElements.Count;
                
                if (columnsCount == 1)
                {
                    var row = table.Elements<TableRow>().First();
                    var s = row.Descendants<TableCell>().First();

                    block = s.InnerText.ToLowerInvariant().Replace("секция", "").Trim();
                }

                if (columnsCount != 20)
                {
                    continue;
                }

                foreach (var row in table.Elements<TableRow>())
                {
                    var s = row.Descendants<TableCell>().ToArray();
                    
                    if (s[1].InnerText.ToLowerInvariant().Contains("узлы на стапель"))
                    {
                        continue;
                    }

                    if (s.Length == 1 && s[0].InnerText == "")
                    {
                        continue;
                    }

                    if (s[0].InnerText.ToLowerInvariant().Contains("узел"))
                    {
                        assembly = s[0].InnerText;
                        continue;
                    }

                    assembly = s[1].InnerText.ToLowerInvariant() switch
                    {
                        { } a when a.Contains("узел") => s[1].InnerText,
                        { } b when b.Contains("листы но") => s[1].InnerText,
                        { } c when c.Contains("рж но") => s[1].InnerText,
                        { } e when e.Contains("листы второго дна") => s[1].InnerText,
                        { } f when f.Contains("рж второго дна") => s[1].InnerText,
                        { } g when g.Contains("россыпь на секцию") => s[1].InnerText,
                        { } h when h.Contains("россыпь на стапель") => s[1].InnerText,
                        "листы" => s[1].InnerText,
                        "рж" => s[1].InnerText,
                        { } m when m.Contains("россыпь на подсекцию") => s[1].InnerText,

                        _ => assembly
                    };
                    
                    if (s[1].InnerText.ToLowerInvariant() == "сводные данные")
                    {
                        break;
                    }
                    
                    if (string.IsNullOrEmpty(s[0].InnerText.Trim()))
                    {
                        continue;
                    }
                        
                    if (string.IsNullOrWhiteSpace(s[0].InnerText))
                    {
                        continue;
                    }

                    var str1 = Regexp(" " + s[1].InnerText + " ",
                            "(?<!,)(?<=[S,s])[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}(?=\\s{1,})|[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}[\\s]?[x|х][0-9]{1,}[\\.|,]{0,1}[0-9]{0,}(?=\\s{1,})|(?<=\\s{1})[R,r](?<!,)[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}[a,b,а,б]{0,1}(?=\\s{1,})|(?<=[П][р][у][т][о][к][ ])[0-9]{1,}(?=\\s{1,})|[0-9]{1,}[\\*][0-9]{1,3}[\\.|,]{0,1}[0-9]{0,}")
                        .ToLower().Replace(",", ".");

                    if (!string.IsNullOrEmpty(str1))
                    {
                        if (str1.Contains("x"))
                        {
                            str1 = str1.Replace(" ", "");

                            var split = str1.Split('x');
                            if (!split[0].Contains("."))
                            {
                                split[0] += ".0";
                            }

                            str1 = split[1] + "*" + split[0];
                        }
                        else if (str1.Contains("х"))
                        {
                            str1 = str1.Replace(" ", "");

                            var split = str1.Split('х');
                            if (!split[0].Contains("."))
                            {
                                split[0] += ".0";
                            }

                            str1 = split[1] + "*" + split[0];
                        }

                        else if (str1.Contains("r") || str1.Contains("р"))
                        {
                            var dataRow = dataTable.Rows.Find(str1.Replace("r", "").Replace("р", "").Replace("a", "а"));
                            if (dataRow != null)
                            {
                                str1 = dataRow[1].ToString();
                            }
                        }
                    }


                    
                    
                    var pos = new Docx();
                    pos.Block = block;
                    pos.PosNo = int.Parse(s[0].InnerText, CultureInfo.InvariantCulture);
                    
                    pos.Quantity = string.IsNullOrWhiteSpace(s[4].InnerText) ? 1 : int.Parse(s[4].InnerText, CultureInfo.InvariantCulture);
                    
                    
                    
                    pos.Dimension = str1;
                    pos.Quality = RenameMaterials(s[11].InnerText);

                    if (pos.Quantity > 1)
                    {
                        if (s[5].InnerText != "")
                        {
                            pos.Weight = double.Parse(s[5].InnerText, CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            pos.Weight = double.Parse(s[6].InnerText, CultureInfo.InvariantCulture)/pos.Quantity;
                        }
                    }
                    else
                    {
                        pos.Weight = double.Parse(s[6].InnerText, CultureInfo.InvariantCulture);
                    }


                    //pos.Weight = int.Parse(s[4]) > 1 ? double.Parse(s[5], CultureInfo.InvariantCulture) : double.Parse(s[6], CultureInfo.InvariantCulture);
                    pos.Assembly = assembly;

                    if (res.ContainsKey(pos.PosNo))
                    {
                        res[pos.PosNo].Quantity += pos.Quantity;
                    }
                    else
                    {
                        res[pos.PosNo] = pos;
                    }
                }
            }
            
            
            



            /*foreach (var s in specTable!.Select(t => t.Split('\t')))
            {
                if (s.Length == 1)
                {
                    continue;
                }

                if (s[1].ToLowerInvariant().Contains("узлы на стапель"))
                {
                    continue;
                }

                if (s.Length == 1 && s[0] == "")
                {
                    continue;
                }

                if (s[0].ToLowerInvariant().Contains("узел"))
                {
                    assembly = s[0];
                    continue;
                }

                assembly = s[1].ToLowerInvariant() switch
                {
                    { } a when a.Contains("узел") => s[1],
                    { } b when b.Contains("листы но") => s[1],
                    { } c when c.Contains("рж но") => s[1],
                    { } e when e.Contains("листы второго дна") => s[1],
                    { } f when f.Contains("рж второго дна") => s[1],
                    { } g when g.Contains("россыпь на секцию") => s[1],
                    { } h when h.Contains("россыпь на стапель") => s[1],
                    { } d when d == "листы" => s[1],
                    { } l when l == "рж" => s[1],
                    { } m when m.Contains("россыпь на подсекцию") => s[1],


                    _ => assembly
                };







                if (s[1].ToLower().Contains("сводные данные"))
                {
                    break;
                }

                if (string.IsNullOrEmpty(s[0].Trim()))
                {
                    continue;
                }


                var str1 = Regexp(" " + s[1] + " ",
                        "(?<!,)(?<=[S,s])[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}(?=\\s{1,})|[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}[\\s]?[x|х][0-9]{1,}[\\.|,]{0,1}[0-9]{0,}(?=\\s{1,})|(?<=\\s{1})[R,r](?<!,)[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}[a,b,а,б]{0,1}(?=\\s{1,})|(?<=[П][р][у][т][о][к][ ])[0-9]{1,}(?=\\s{1,})|[0-9]{1,}[\\*][0-9]{1,3}[\\.|,]{0,1}[0-9]{0,}")
                    .ToLower().Replace(",", ".");

                if (!string.IsNullOrEmpty(str1))
                {
                    if (str1.Contains("x"))
                    {
                        str1 = str1.Replace(" ", "");

                        var split = str1.Split('x');
                        if (!split[0].Contains("."))
                        {
                            split[0] += ".0";
                        }

                        str1 = split[1] + "*" + split[0];
                    }
                    else if (str1.Contains("х"))
                    {
                        str1 = str1.Replace(" ", "");

                        var split = str1.Split('х');
                        if (!split[0].Contains("."))
                        {
                            split[0] += ".0";
                        }

                        str1 = split[1] + "*" + split[0];
                    }

                    else if (str1.Contains("r") || str1.Contains("р"))
                    {
                        var dataRow = dataTable.Rows.Find(str1.Replace("r", "").Replace("р", "").Replace("a", "а"));
                        if (dataRow != null)
                        {
                            str1 = dataRow[1].ToString();
                        }
                    }
                }

                if (string.IsNullOrWhiteSpace(s[4]))
                {
                    s[4] = "1";
                }

                var pos = new Docx();
                pos.PosNo = int.Parse(s[0], CultureInfo.InvariantCulture);
                pos.Quantity = int.Parse(s[4], CultureInfo.InvariantCulture);
                pos.Dimension = str1;
                pos.Quality = RenameMaterials(s[11]);

                if (int.Parse(s[4]) > 1)
                {
                    if (s[5] != "")
                    {
                        pos.Weight = double.Parse(s[5], CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        pos.Weight = double.Parse(s[6], CultureInfo.InvariantCulture)/int.Parse(s[4]);
                    }
                }
                else
                {
                    pos.Weight = double.Parse(s[6], CultureInfo.InvariantCulture);
                }


                //pos.Weight = int.Parse(s[4]) > 1 ? double.Parse(s[5], CultureInfo.InvariantCulture) : double.Parse(s[6], CultureInfo.InvariantCulture);
                pos.Assembly = assembly;

                res[pos.PosNo] = pos;
            }*/

            return res;
        }

        private static string RenameMaterials(string s)
        {
            s = s.ToUpper();

            s = s.Replace("PCA", "A");
            s = s.Replace("PCA36", "A36");
            s = s.Replace("PCA40", "A40");
            s = s.Replace("PCD40", "D40");
            s = s.Replace("PCD500W", "DW");
            s = s.Replace("РСЕ500WP", "EPW");
            s = s.Replace("PCE500W", "EW");
            s = s.Replace("20", "ST20");
            s = s.Replace("PCE500W", "EW");
            s = s.Replace("CТ3CП2", "SP3PS_143");
            s = s.Replace("EW-П", "EPW");
            s = s.Replace("CТST20", "ST20");
            s = s.Replace("PCE40Z35", "EZ");
            //s = s.Replace("H10_B", "H10");

            s = s.Replace("DWARC40", "D");
            s = s.Replace("EWARC30", "EW");

            //DWARC40 - D
            //EWARC30 - EW


            s = s.Replace("EW-P", "EPW");
            s = s.Replace("08X18H10T", "H10");

            return s;
        }

        private static string Regexp(string s, string exp)
        {
            var regex = new Regex(exp);
            string value;
            var matchCollections = regex.Matches(s);
            if (matchCollections.Count <= 1)
            {
                value = matchCollections.Count != 1 ? s : matchCollections[0].Value;
            }
            else
            {
                value = matchCollections[1].Value;
            }
            return value;
        }
    }
}
