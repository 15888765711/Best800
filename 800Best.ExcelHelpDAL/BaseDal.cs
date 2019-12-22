using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace _800Best.ExcelHelpDAL
{
    public class BaseDal<T> where T: class, new()
    {
        //CUD操作
        public void ExecuteNonQuery()
        {
            

        }
        #region 批量上传
        /// <summary>
        /// 批量上传
        /// </summary>
        /// <param name="tableName">数据库中的表名</param>
        /// <param name="dt">自己读取出来的datatable</param>
        /// <returns></returns>
        public int MySqlBulkCopy(string tableName, DataTable dt)
        {

            return SqlHelper.SqlBulkCopyInsert(tableName, dt);

        }
        #endregion

        #region 读取数据
        /// <summary>
        /// 用sql语句读取1个datatable
        /// </summary>
        /// <param name="sqlstr"></param>
        /// <param name="starttime"></param>
        /// <param name="endtime"></param>
        /// <returns></returns>
        public DataTable GetDataTable( string sqlstr, DateTime starttime, DateTime endtime)
        {
            SqlParameter[] sp = new SqlParameter[] {
                new SqlParameter("@starttime",starttime),
                new SqlParameter("@endtime",endtime),

            };
            return SqlHelper.GetTable(sqlstr, sp);
        }

        #endregion

        #region 更新数据

        #endregion



    }
}
