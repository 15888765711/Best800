using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _800Best.ExcelHelpDAL
{


       
        public static class SqlHelper
        {
            private static readonly string constr = ConfigurationManager.ConnectionStrings["Sqlcon"].ConnectionString;

            public static int ExecuteNonQuery(string sql, CommandType cmdType, params SqlParameter[] sp)
            {
                using (SqlConnection connection = new SqlConnection(constr))
                {
                    return ExecuteNonQuery(connection, sql, cmdType, sp);
                }
            }

            public static int ExecuteNonQuery(SqlConnection conn, string sql, CommandType cmdType, params SqlParameter[] ps)
            {
                using (SqlCommand command = new SqlCommand(sql, conn))
                {
                    command.CommandType = cmdType;
                    conn.Open();
                    if (ps != null)
                    {
                        command.Parameters.AddRange(ps);
                    }
                    return command.ExecuteNonQuery();
                }
            }

            public static SqlDataReader ExecuteReader(string sql, CommandType cmdType, params SqlParameter[] sp)
            {
                SqlDataReader reader;
                SqlConnection connection = new SqlConnection(constr);
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    command.CommandType = cmdType;
                    command.Parameters.Clear();
                    if (sp != null)
                    {
                        command.Parameters.AddRange(sp);
                    }
                    try
                    {
                        connection.Open();
                    command.CommandTimeout = 0;
                        reader = command.ExecuteReader(CommandBehavior.CloseConnection);
                    }
                    catch (Exception exception)
                    {
                        connection.Close();
                        connection.Dispose();
                        throw exception;
                    }
                }
                return reader;
            }

            public static object ExecuteScale(string sql, CommandType cmdType, params SqlParameter[] ps)
            {
                using (SqlConnection connection = new SqlConnection(constr))
                {
                    return ExecuteScale(connection, sql, cmdType, ps);
                }
            }

            public static object ExecuteScale(SqlConnection conn, string sql, CommandType cmdType, params SqlParameter[] ps)
            {
                using (SqlCommand command = new SqlCommand(sql, conn))
                {
                    command.CommandType = cmdType;
                    conn.Open();
                    if (ps != null)
                    {
                        command.Parameters.Add(ps);
                    }
                    return command.ExecuteScalar();
                }
            }

            public static DataTable GetTable(string sql, params SqlParameter[] ps)
            {
                DataTable dataTable = new DataTable();
                using (SqlDataAdapter adapter = new SqlDataAdapter(sql, constr))
                {
                    if (ps != null)
                    {
                        adapter.SelectCommand.Parameters.Add(ps);
                    }
                    adapter.Fill(dataTable);
                }
                return dataTable;
            }

            public static int SqlBulkCopyInsert(string strTableName, DataTable dtData)
            {
                try
                {
                    using (SqlBulkCopy copy = new SqlBulkCopy(constr))
                    {   
                        copy.DestinationTableName = strTableName;
                        copy.NotifyAfter = dtData.Rows.Count;
                        copy.BatchSize = dtData.Rows.Count;
                        copy.WriteToServer(dtData);
                        copy.Close();
                        return 1;
                    }
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
        }
    }





