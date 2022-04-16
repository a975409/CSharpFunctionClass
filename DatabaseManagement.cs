using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace Northwind
{
    public static class DatabaseManagement
    {
        public static readonly string conStr = ConfigurationManager.ConnectionStrings["NorthwindConnectionString"].ConnectionString;

        /// <summary>
        /// 透過Sql語法取得某一段區間的查詢結果
        /// </summary>
        /// <param name="page">頁碼</param>
        /// <param name="size">單一頁最大資料筆數</param>
        /// <param name="sql">sql語法</param>
        /// <param name="parameters">套用到where條件查詢的參數</param>
        /// <returns></returns>
        public static DataTable SqlQueryDataTablePageSize(int page, int size, string sql, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(conStr))
            {
                conn.Open();
                sql += " OFFSET @offset ROWS FETCH NEXT @pageSize ROWS ONLY";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@offset", (page * size));
                    cmd.Parameters.AddWithValue("@pageSize", size);

                    if (parameters.Length > 0)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        dt.Load(reader);
                        return dt;
                    }
                }
            }
        }

        public static DataTable SqlQueryDataTable(string sql, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(conStr))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    if (parameters.Length > 0)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        dt.Load(reader);
                        return dt;
                    }
                }
            }
        }

    }
}
