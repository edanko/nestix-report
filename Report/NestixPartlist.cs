using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows;

namespace NestixReport
{
    public class NestixPartlist
    {
        public string NcName {get;set;}
        public string Order {get;set;}
        public string Section {get;set;}
        public string PosNo {get;set;}
        public int Count {get;set;}
        public double Weight {get;set;}
        public double TotalWeight {get;set;}
        public double Length {get;set;}
        public double Width {get;set;}







        public static List<NestixPartlist> GetPartlistFromNestix(string connectionString, string filter)
        {
            var con = new SqlConnection(connectionString);
            var ncCom = new SqlCommand(Db.GetNxPathIds, con);
            ncCom.Parameters.AddWithValue("name", filter);

            try
            {
                con.Open();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }

            var ncReader = ncCom.ExecuteReader();

            var ncWithPathId = new Dictionary<string, int>();

            while (ncReader.Read())
            {
                ncWithPathId.Add((string)ncReader["name"], (int)ncReader["id"]);
            }
            ncReader.Close();

            var partlist = new List<NestixPartlist>();

            foreach (var k in ncWithPathId.Keys)
            {
                var partsCom = new SqlCommand(Db.GetPartsFromNxPathId, con);
                partsCom.Parameters.AddWithValue("pathid", ncWithPathId[k]);
                var partsReader = partsCom.ExecuteReader();

                while (partsReader.Read())
                {
                    var p = new NestixPartlist
                    {
                        NcName = k,
                        Order = partsReader.GetString(0),
                        Section = partsReader.GetString(1),
                        PosNo = partsReader.GetString(2),
                        Count = partsReader.GetInt32(3),
                        Weight = partsReader.GetFloat(4),
                        TotalWeight = partsReader.GetFloat(5),
                        Length = partsReader.GetFloat(8),
                        Width = partsReader.GetFloat(9)
                    };
                    partlist.Add(p);
                }

                partsReader.Close();
            }
            con.Close();

            ncWithPathId.Clear();

            return partlist;
        }
    }
}
