using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;

namespace NestixReport
{
    partial class MainWindow
    {
        private string GetWorkingDirectory()
        {
            var f = DbComboBox.Text switch
            {
                "NxSC_Zvezda_10510" => "nxsc_zvezda_leader",
                _ => DbComboBox.Text
            };

            return $@"\\bk-ssk-nestix01.corp.local\{f}\master";
        }

        private string GetExePathForCutting()
        {
            var f = DbComboBox.Text switch
            {
                "NxSC_Zvezda_10510" => "nxsc_zvezda_leader",
                _ => DbComboBox.Text
            };

            return $@"\\bk-ssk-nestix01.corp.local\{f}\bin\Cutting.exe";
        }

        private void RunNestixButtonClick(object sender, RoutedEventArgs e)
        {
            var process = new Process
            {
                StartInfo =
                {
                    FileName = GetExePathForCutting(),
                    Arguments = "-nxsite=Zvezda",
                    WorkingDirectory = GetWorkingDirectory()
                }
            };
            process.Start();
        }

        private string GetExePathForOldReport()
        {
            var f = DbComboBox.Text switch
            {
                "NxSC_Zvezda_10510" => "nxsc_zvezda_leader",
                _ => DbComboBox.Text
            };

            return $@"\\bk-ssk-nestix01.corp.local\{f}\bin\DDRepoviewU.exe";
        }

        private string GetExePathForNewReport()
        {
            var f = DbComboBox.Text switch
            {
                "NxSC_Zvezda_10510" => "nxsc_zvezda_leader",
                _ => DbComboBox.Text
            };

            return $@"\\bk-ssk-nestix01.corp.local\{f}\bin\report\report.exe";
        }

        
        private void OpenOldReportClick(object sender, RoutedEventArgs e)
        {
            var con = new SqlConnection(GetConnectionString());
            var com = new SqlCommand(Db.GetNxPathIdsForReport, con);

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

            var idsList = new List<string>();

            while (reader.Read())
            {
                idsList.Add(reader.GetInt32(0).ToString());
            }
            reader.Close();
            con.Close();


            if (idsList.Count == 0)
            {
                MessageBox.Show("Карты не найдены");
                return;
            }

            var arr = new ArrayList();

            var cnt = 0;
            var tmpList = new List<string>();

            if (idsList.Count > 100)
            {
                foreach (var id in idsList)
                {
                    if (cnt == 100)
                    {
                        arr.Add(tmpList);
                        tmpList.Clear();
                        cnt = 0;
                    }

                    tmpList.Add(id);
                    cnt++;
                }
            }
            else
            {
                arr.Add(idsList);
            }

            foreach (List<string> curList in arr)
            {
                var process = new Process
                {
                    StartInfo =
                    {
                        FileName = GetExePathForOldReport(),
                        Arguments = $@"-INI=.\Settings\nestix2.ini -SEC=DD_SHIP_Report -PARAMS={string.Join(",",curList)} -NXLANG=Eng",
                        WorkingDirectory = GetWorkingDirectory()
                    }
                };
                process.Start();
            }
        }

        private void OpenNewReportClick(object sender, RoutedEventArgs e)
        {
            var con = new SqlConnection(GetConnectionString());
            var com = new SqlCommand(Db.GetNxPathIdsForReport, con);

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

            var idsList = new List<string>();

            while (reader.Read())
            {
                idsList.Add(reader.GetInt32(0).ToString());
            }
            reader.Close();
            con.Close();

            if (idsList.Count == 0)
            {
                MessageBox.Show("Карты не найдены");
                return;
            }

            var process = new Process
            {
                StartInfo =
                {
                    FileName = GetExePathForNewReport(),
                    Arguments = $@"-PARAMS={string.Join(",",idsList)}",
                    WorkingDirectory = GetWorkingDirectory()
                }
            };
            process.Start();
        }

    }
}
