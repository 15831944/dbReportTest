using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace checkDwgs
{
    class dbUmmm
    {
        public static string SimpleUpdateDB(string connection, string sql)
        {
            //SqlConnection sqlConn1 = new SqlConnection(System.Configuration.ConfigurationManager.AppSettings["sqlConn.ConnectionString"]);
            if (!string.IsNullOrEmpty(sql.Trim()))
            {
                var connectionString = connection; //System.Configuration.ConfigurationManager.AppSettings[connection];
                var sqlConn1 = new SqlConnection(connectionString);
                
                sqlConn1.Open();
                SqlCommand cmd;
                cmd = new SqlCommand(sql, sqlConn1);
                try
                {
                    int result = cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw;
                }
                finally
                {
                    sqlConn1.Close();
                }

                return null;
            }

            return "";
        }


        //For getting a value for the db

        public static string GetParameterValue(object o)
        {
            if (o == null)
            {
                return "NULL";
            }

            if (o is TimeSpan)
            {
                return ((TimeSpan)o).Ticks.ToString();
            }

            if (o is bool)
            {
                return Convert.ToInt32((bool)o).ToString();
            }

            if (o is byte[])
            {
                return "0x" + BitConverter.ToString((byte[])o).Replace("-", "");
            }

            return "'" + o.ToString().Replace("'", "''") + "'";
        }


        //Example of using said functions
        /*
public void UpdateNetworkFolderInformation(Project project)
    {
        var projectSQL = @"UPDATE projects SET " +
                  @" network_folder = " + DBActions.GetParameterValue(project.NetworkFolder) +
                  @", last_updater = " + DBActions.GetParameterValue(EMISUtility.CurrentEMISUser.ShortNTID) +
                  @", last_updated = " + DBActions.GetParameterValue(DateTime.Now) +
                  @" WHERE prj_code = " + DBActions.GetParameterValue(project.ProjectCode);

        DBActions.SimpleUpdateDB(EMISUtility.EEIDWConnection, projectSQL);
    }
    */


    }

}
