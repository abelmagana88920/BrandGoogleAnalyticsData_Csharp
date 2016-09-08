using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.IO;
namespace BrandGoogleAnalyticsData
{
    public static class Library
    {
        public static void WriteErrorLog(string Message)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\logFile.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + Message);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }
        public static void Execute(string command)
        {
            string[] sep = { "provider connection string=" };
            string metadata_con = ConfigurationManager.ConnectionStrings["IFSReportingContext"].ConnectionString;
            string[] metadata_con_array = Explode(metadata_con ,sep);
            string datasource_conn = metadata_con_array[1].Replace("\"", "");

                System.Data.SqlClient.SqlConnection sqlConnection1 =
                new System.Data.SqlClient.SqlConnection(datasource_conn);

                System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                cmd.CommandType = System.Data.CommandType.Text;
                cmd.CommandText = command;
                cmd.Connection = sqlConnection1;

                sqlConnection1.Open();
                cmd.ExecuteNonQuery();
                sqlConnection1.Close();
         
        }

        public static string[] Explode(string field, string[] separators)
        {
            string[] field_array = field.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return field_array;
        }
    }
}