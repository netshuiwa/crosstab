using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for SqlLayer
/// </summary>
public class SqlLayer
{
	public SqlLayer()
	{
		//
		// TODO: Add constructor logic here
		//
	}

    public static DataTable GetDataTable(string spName)
    {
        try
        {
            string strCon = ConfigurationManager.ConnectionStrings["PivotConnectionString"].ConnectionString;
            using(SqlConnection con = new SqlConnection(strCon))
            {
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter();
                SqlCommand cmd = new SqlCommand(spName, con);
                cmd.CommandType = CommandType.StoredProcedure;
                da.SelectCommand = cmd;

                DataSet ds = new DataSet();
                da.Fill(ds);

                if (ds.Tables.Count > 0)
                    return ds.Tables[0];
            }
        }
        catch (Exception)
        {
            throw;
        }
        return null;
    }
}