using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Windows.Forms;

namespace QSMS
{
    class SqlHelper
    {
        private static SqlConnection cn;
        private static SqlDataAdapter sda;
        private static DataTable dt;
        private static DataSet ds;
        private static SqlCommand cmd;

        public static DataTable ExecuteTable(string strSql, string strconnect)
        {
            using (cn = new SqlConnection(strconnect))
            {
                try
                {
                    sda = new SqlDataAdapter(strSql, cn);
                    sda.SelectCommand.CommandTimeout = 300;
                    dt = new DataTable();
                    sda.Fill(dt);
                    return dt;
                }
                catch (Exception ex)
                {
                    sda = null;
                    MessageBox.Show(ex.Message.ToString(), "DB Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }
        }

        public static DataTable ExecuteDataTable(string spName, SqlParameter[] paras, string strconnect)
        {
            using (cn = new SqlConnection(strconnect))
            {
                try
                {
                    sda = new SqlDataAdapter(spName, cn);
                    if (paras.Length > 0)
                    {
                        foreach (SqlParameter parameter in paras)
                        {
                            sda.SelectCommand.Parameters.Add(parameter);
                        }
                    }
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure;
                    sda.SelectCommand.CommandTimeout = 0;
                    dt = new DataTable();
                    sda.Fill(dt);
                    sda.Dispose();
                    return dt;
                }
                catch (Exception ex)
                {
                    sda = null;
                    MessageBox.Show(ex.Message.ToString(), "DB Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }
        }

        public static DataSet ExecuteDataSet(string strSql, string strconnect)
        {
            using (cn = new SqlConnection(strconnect))
            {
                try
                {
                    sda = new SqlDataAdapter(strSql, cn);
                    sda.SelectCommand.CommandTimeout = 300;
                    ds = new DataSet();
                    sda.Fill(ds);
                    return ds;
                }
                catch (Exception ex)
                {
                    sda = null;
                    MessageBox.Show(ex.Message.ToString(), "DB Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }
        }

        public static DataSet ExecuteDataSet(string strSql, SqlParameter[] paras, string strconnect)
        {
            using (cn = new SqlConnection(strconnect))
            {
                try
                {
                    sda = new SqlDataAdapter(strSql, cn);
                    if (paras.Length > 0)
                    {
                        foreach (SqlParameter parameter in paras)
                        {
                            sda.SelectCommand.Parameters.Add(parameter);
                        }
                    }
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure;
                    sda.SelectCommand.CommandTimeout = 300;
                    ds = new DataSet();
                    sda.Fill(ds);
                    sda.Dispose();
                    return ds;
                }
                catch (Exception ex)
                {
                    sda = null;
                    MessageBox.Show(ex.Message.ToString(), "DB Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }
        }

        public static void Executeless(string strSql, CommandType commandType, SqlParameter[] paras, string strconnect)
        {
            cn = null;
            try
            {
                cn = CreateConnection(strconnect);
                cmd = new SqlCommand(strSql, cn);
                if (paras != null)
                {
                    foreach (var s in paras)
                    {
                        cmd.Parameters.Add(s);
                    }
                }
                cmd.CommandType = commandType;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                CloseConnection(cn);
            }
        }  

        private static void CloseConnection(SqlConnection cn)
        {
            try
            {
                if (cn == null)
                {
                    return;
                }
                if (cn.State != ConnectionState.Closed)
                {
                    cn.Dispose();
                    cn.Close();
                }
            }
            finally
            {
                cn = null;
            }
        }  

        private static SqlConnection CreateConnection(string strConn)
        {
            try
            {
                if (String.IsNullOrEmpty(strConn))
                {
                    return null;
                }
                else
                {
                    SqlConnection cn = new SqlConnection(strConn);
                    cn.Open();
                    return cn;
                }
            }
            catch
            {
                return null;
            }
        } 
    }
}
