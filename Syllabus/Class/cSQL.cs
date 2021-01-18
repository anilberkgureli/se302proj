using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace Syllabus
{
    class cSQL
    {
        /*ParamaterName, ParamaterValue, ParameterType, ParameterSize*/
        public bool CallProcedure(string sProcedureName, List<Tuple<string, string, SqlDbType, int>> sParameters)
        {
            try
            {
                SqlConnection conn = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=" + Directory.GetCurrentDirectory() + @"\Database.mdf");
                conn.Open();
                SqlTransaction transaction = conn.BeginTransaction();
                SqlCommand cmd = new SqlCommand(sProcedureName, conn, transaction);
                try
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (sParameters != null)
                    {
                        foreach (Tuple<string, string, SqlDbType, int> ls in sParameters)
                        {
                            if (ls.Item3.Equals(SqlDbType.Date))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Date).Value = Convert.ToDateTime(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.Time))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Time).Value = TimeSpan.Parse(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.VarChar))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.VarChar, ls.Item4).Value = ls.Item2;
                            else if (ls.Item3.Equals(SqlDbType.Char))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Char, ls.Item4).Value = ls.Item2;
                            else if (ls.Item3.Equals(SqlDbType.NChar))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Char, ls.Item4).Value = ls.Item2;
                            else if (ls.Item3.Equals(SqlDbType.TinyInt))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.TinyInt).Value = Convert.ToByte(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.SmallInt))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.SmallInt).Value = Convert.ToInt16(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.Int))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Int).Value = Convert.ToInt32(ls.Item2);
                        }
                    }
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                    transaction.Commit();
                    return true;
                }
                catch (SqlException sSQL)
                {
                    transaction.Rollback();
                    MessageBox.Show(sSQL.Message);
                    return false;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                    cmd.Dispose();
                    transaction.Dispose();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                return false;
            }
        }

        public int CallProcedure(string sProcedureName, List<Tuple<string, string, SqlDbType, int>> sParameters, bool isOutput)
        {
            try
            {
                SqlConnection conn = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=" + Directory.GetCurrentDirectory() + @"\Database.mdf");
                conn.Open();
                SqlTransaction transaction = conn.BeginTransaction();
                SqlCommand cmd = new SqlCommand(sProcedureName, conn, transaction);
                try
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (sParameters != null)
                    {
                        foreach (Tuple<string, string, SqlDbType, int> ls in sParameters)
                        {
                            if (ls.Item3.Equals(SqlDbType.Date))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Date).Value = Convert.ToDateTime(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.Time))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Time).Value = TimeSpan.Parse(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.VarChar))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.VarChar, ls.Item4).Value = ls.Item2;
                            else if (ls.Item3.Equals(SqlDbType.Char))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Char, ls.Item4).Value = ls.Item2;
                            else if (ls.Item3.Equals(SqlDbType.NChar))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Char, ls.Item4).Value = ls.Item2;
                            else if (ls.Item3.Equals(SqlDbType.TinyInt))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.TinyInt).Value = Convert.ToByte(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.SmallInt))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.SmallInt).Value = Convert.ToInt16(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.Int))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Int).Value = Convert.ToInt32(ls.Item2);
                        }
                    }
                    cmd.Parameters.Add("@iResult", SqlDbType.TinyInt);
                    cmd.Parameters["@iResult"].Direction = ParameterDirection.Output;
                    cmd.ExecuteNonQuery();
                    transaction.Commit();
                    return Convert.ToInt16(cmd.Parameters["@iResult"].Value);

                }
                catch (SqlException sSQL)
                {
                    transaction.Rollback();
                    MessageBox.Show(sSQL.Message);
                    return 0;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                    cmd.Dispose();
                    transaction.Dispose();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                return 0;
            }
        }

        public DataTable CallProcedure(string sProcedureName, List<Tuple<string, string, SqlDbType, int>> sParameters, bool isOutput, bool isTable)
        {
            try
            {
                SqlConnection conn = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=" + Directory.GetCurrentDirectory() + @"\Database.mdf");
                DataTable dt = new DataTable();
                conn.Open();
                SqlTransaction transaction = conn.BeginTransaction();
                SqlCommand cmd = new SqlCommand(sProcedureName, conn, transaction);
                try
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (sParameters != null)
                    {
                        foreach (Tuple<string, string, SqlDbType, int> ls in sParameters)
                        {
                            if (ls.Item3.Equals(SqlDbType.Date))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Date).Value = Convert.ToDateTime(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.Time))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Time).Value = TimeSpan.Parse(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.VarChar))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.VarChar, ls.Item4).Value = ls.Item2;
                            else if (ls.Item3.Equals(SqlDbType.Char))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Char, ls.Item4).Value = ls.Item2;
                            else if (ls.Item3.Equals(SqlDbType.NChar))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Char, ls.Item4).Value = ls.Item2;
                            else if (ls.Item3.Equals(SqlDbType.TinyInt))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.TinyInt).Value = Convert.ToByte(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.SmallInt))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.SmallInt).Value = Convert.ToInt16(ls.Item2);
                            else if (ls.Item3.Equals(SqlDbType.Int))
                                cmd.Parameters.Add("@" + ls.Item1, SqlDbType.Int).Value = Convert.ToInt32(ls.Item2);
                        }
                    }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    transaction.Commit();
                    return dt;

                }
                catch (SqlException sSQL)
                {
                    transaction.Rollback();
                    MessageBox.Show(sSQL.Message);
                    return null;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                    cmd.Dispose();
                    dt.Dispose();
                    transaction.Dispose();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                return null;
            }
        }
    }
}
