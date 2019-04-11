using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections;


/// <summary>
/// DbHelperSQL 的摘要说明
/// </summary>
namespace OpcTest
{
    public abstract class DbHelperSQL
    {
        public static string connectionString = "server=192.168.3.89;database=OPCTEST;uid=sa;pwd=sa123";
        //public static string connectionString = "server=192.168.3.88;database=OPCTEST;uid=sa;pwd=";
        //public static string connectionString = "server=(local);database=OPCTEST;Integrated Security=True";

        #region 公用方法

        /// <summary>
        /// 在一个表中寻找一列数据的最大值
        /// </summary>
        /// <param name="FieldName">列名</param>
        /// <param name="TableName">表名</param>
        /// <returns></returns>
        public static int GetMaxID(string FieldName, string TableName)
        {
            string strsql = string.Format("select max({0})+1 from {1}", FieldName, TableName);
            object obj = GetSingle(strsql);
            if (obj == null)
            {
                return 1;
            }
            else
            {
                return int.Parse(obj.ToString());
            }
        }
        /// <summary>
        /// 在一个表中寻找一列数据的最小值
        /// </summary>
        /// <param name="FieldName">列名</param>
        /// <param name="TableName">表名</param>
        /// <returns></returns>
        public static int GetMinID(string FieldName, string TableName) 
        {
            string strsql = string.Format("select min({0})+1 from {1}", FieldName, TableName);
            object obj = GetSingle(strsql);
            if (obj == null)
            {
                return 1;
            }
            else
            {
                return int.Parse(obj.ToString());
            }
        }
        /// <summary>
        /// 判断GetSingle函数返回的值是否为空
        /// </summary>
        /// <param name="strSql"></param>
        /// <param name="cmdParms"></param>
        /// <returns></returns>
        public static bool Exists(string strSql, params SqlParameter[] cmdParms)
        {
            object obj = GetSingle(strSql, cmdParms);
            int cmdresult;
            if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
            {
                cmdresult = 0;
            }
            else
            {
                cmdresult = int.Parse(obj.ToString());
            }
            if (cmdresult == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        #endregion

        #region  执行简单SQL语句
        /// <summary>
        /// 执行SQL语句，返回影响的记录数
        /// </summary>
        /// <param name="SQLString">数据库语句</param>
        /// <returns>影响的记录数</returns>
        public static int ExecuteSql(string SQLString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(SQLString, connection))
                {
                    try
                    {
                        connection.Open();
                        int rows = cmd.ExecuteNonQuery();
                        return rows;
                    }
                    catch (SqlException E)
                    {
                        connection.Close();
                        throw new Exception(E.Message);
                    }
                }
            }
        }
        /// <summary>
        /// 执行SQL语句，不返回影响的记录数
        /// </summary>
        /// <param name="SQLString"></param>
        public static void ExecuteSqlxx(string SQLString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(SQLString, connection))
                {
                    try
                    {
                        connection.Open();
                        int rows = cmd.ExecuteNonQuery();
                    }
                    catch (SqlException E)
                    {
                        connection.Close();
                        throw new Exception(E.Message);
                    }
                }
            }
        }
        /// <summary>
        /// 执行多条SQL语句，实现数据库事务。
        /// </summary>
        /// <param name="SQLStringList">数据库语句集合</param>
        public static void ExecuteSqlTran(ArrayList SQLStringList)
        {

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand() { Connection = conn };
                SqlTransaction tx = conn.BeginTransaction();
                cmd.Transaction = tx;
                try
                {
                    for (int n = 0; n < SQLStringList.Count; n++)
                    {
                        string strsql = SQLStringList[n].ToString();
                        if (strsql.Trim().Length > 1)
                        {
                            cmd.CommandText = strsql;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    tx.Commit();
                }
                catch (SqlException E)
                {
                    tx.Rollback();
                    throw new Exception(E.Message);
                }

            }
        }
        /// <summary>
        /// 通过临时表，一次进行插入多条数据库语句
        /// </summary>
        /// <param name="dtSource"></param>临时表来源
        /// <param name="DestinationTableName"></param>目标数据库表的名称
        public static void ExecuteSqlBulkCopy(DataTable dtSource, string DestinationTableName)
        {         
            using (SqlBulkCopy copy = new SqlBulkCopy(connectionString))
            {              
                copy.DestinationTableName = DestinationTableName;            
                copy.ColumnMappings.Add("virtual_address", "VIRTUAL_ADDRESS");
                copy.ColumnMappings.Add("device_key", "DEVICE_KEY");
                copy.ColumnMappings.Add("address_value", "ADDRESS_VALUE");
                copy.ColumnMappings.Add("methed", "COLLECT_METHOD");
                copy.ColumnMappings.Add("collect_equip", "COLLECT_EQUIP");
                copy.ColumnMappings.Add("protocol", "PROTOCOL");
                copy.ColumnMappings.Add("ip", "IP_ADDRESS");
                copy.ColumnMappings.Add("physical_address", "PHYSICAL_ADDRESS");
                copy.ColumnMappings.Add("enable", "ENABLE");
                copy.ColumnMappings.Add("state", "STATE");
                copy.ColumnMappings.Add("write", "WRITE");          
                copy.ColumnMappings.Add("description", "DESCRIPTION");
                copy.ColumnMappings.Add("remarks", "REMARKS");                       
                copy.WriteToServer(dtSource);
            }
        }
        /// <summary>
        /// 执行带一个存储过程参数的的SQL语句。返回影响的记录数
        /// </summary>
        /// <param name="SQLString">数据库语句</param>
        /// <param name="content"></param>
        /// <returns></returns>
        public static int ExecuteSql(string SQLString, string content)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(SQLString, connection);
                SqlParameter myParameter = new SqlParameter("@content", SqlDbType.NText) { Value = content /* <param name="content">参数内容,比如一个字段是格式复杂的文章，有特殊符号，可以通过这个方式添加</param>*/ };
                cmd.Parameters.Add(myParameter);
                try
                {
                    connection.Open();
                    int rows = cmd.ExecuteNonQuery();
                    return rows;//影响的记录数
                }
                catch (SqlException E)
                {
                    throw new Exception(E.Message);
                }
                finally
                {
                    cmd.Dispose();
                    connection.Close();
                }
            }
        }
        /// <summary>
        /// 向数据库里插入图像格式的字段(和上面情况类似的另一种实例)
        /// </summary>
        /// <param name="strSQL">数据库语句</param>
        /// <param name="fs"></param>
        /// <returns>影响的记录数</returns>
        public static int ExecuteSqlInsertImg(string strSQL, byte[] fs)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(strSQL, connection);
                SqlParameter myParameter = new SqlParameter("@fs", SqlDbType.Image) { Value = fs /* <param name="fs">图像字节,数据库的字段类型为image的情况</param>*/ };
                cmd.Parameters.Add(myParameter);
                try
                {
                    connection.Open();
                    int rows = cmd.ExecuteNonQuery();
                    return rows;//影响的记录数
                }
                catch (SqlException E)
                {
                    throw new Exception(E.Message);
                }
                finally
                {
                    cmd.Dispose();
                    connection.Close();
                }
            }
        }
        /// <summary>
        /// 执行查询，并返回查询所返回的结果集中第一行的第一列（object类型）。忽略其他列或行。
        /// </summary>
        /// <param name="SQLString">数据库查询语句</param>
        /// <returns>结果集中第一行的第一列（object类型）</returns>
        public static object GetSingle(string SQLString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(SQLString, connection))// <param name="SQLString">计算查询结果语句</param>
                {
                    try
                    {
                        connection.Open();
                        object obj = cmd.ExecuteScalar();
                        if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
                        {
                            return null;
                        }
                        else
                        {
                            return obj;//查询结果
                        }
                    }
                    catch (SqlException e)
                    {
                        connection.Close();
                        throw new Exception(e.Message);
                    }
                }
            }
        }
        /// <summary>
        /// 执行查询语句，返回SqlDataReader
        /// </summary>
        /// <param name="strSQL">查询语句</param>
        /// <returns>SqlDataReader</returns>
        public static SqlDataReader ExecuteReader(string strSQL)// 执行查询语句，返回SqlDataReader
        {
            SqlConnection connection = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(strSQL, connection);
            try
            {
                connection.Open();
                SqlDataReader myReader = cmd.ExecuteReader();
                return myReader;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }

        }
        /// <summary>
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="SQLString">数据库查询语句</param>
        /// <returns>DataSet</returns>
        public static DataSet Query(string SQLString)// 执行查询语句，返回DataSet
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                DataSet ds = new DataSet();

                try
                {
                    connection.Open();
                    SqlDataAdapter command = new SqlDataAdapter(SQLString, connection);
                    command.Fill(ds, "ds");
                }
                catch (SqlException ex)
                {
                    throw new Exception(ex.Message);
                }
                return ds;
            }
        }

        /// <summary>
        /// 执行查询语句，返回DataTable
        /// </summary>
        /// <param name="sql">数据库查询语句</param>
        /// <returns>Datatable</returns>
        public static DataTable OpenTable(string SQLString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlDataAdapter da = new SqlDataAdapter(SQLString, connection);
                DataTable dt = new DataTable();

                try
                {
                    da.Fill(dt);
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                finally
                {
                    da.Dispose();
                    connection.Close();
                }

                return dt;
            }
        }
        #endregion

        #region 执行带参数的SQL语句

        /// <summary>
        /// 执行SQL语句，返回受影响的记录数
        /// </summary>
        /// <param name="SQLString">SQL语句</param>
        /// <returns>受影响的记录数</returns>
        public static int ExecuteSql(string SQLString, params SqlParameter[] cmdParms)// 执行SQL语句，返回影响的记录数
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    try
                    {
                        PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                        int rows = cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                        return rows;
                    }
                    catch (SqlException E)
                    {
                        throw new Exception(E.Message);
                    }
                }
            }
        }
        /// <summary>
        /// 执行多条SQL语句，实现数据库事务。
        /// </summary>
        /// <param name="SQLStringList">SQL语句的哈希表（key为sql语句，value是该语句的SqlParameter[]）</param>
        public static void ExecuteSqlTran(Hashtable SQLStringList)// 执行多条SQL语句，实现数据库事务。
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    SqlCommand cmd = new SqlCommand();
                    try
                    {
                        //循环
                        foreach (DictionaryEntry myDE in SQLStringList)
                        {
                            string cmdText = myDE.Key.ToString();
                            SqlParameter[] cmdParms = (SqlParameter[])myDE.Value;
                            PrepareCommand(cmd, conn, trans, cmdText, cmdParms);
                            int val = cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();

                            trans.Commit();
                        }
                    }
                    catch
                    {
                        trans.Rollback();
                        throw;
                    }
                }
            }
        }
        /// <summary>
        /// 执行一条计算查询结果语句，返回查询结果（object）。
        /// </summary>
        /// <param name="SQLString">计算查询结果语句</param>
        /// <returns>查询结果（object）</returns>
        public static object GetSingle(string SQLString, params SqlParameter[] cmdParms)// 执行一条计算查询结果语句，返回查询结果（object）
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    try
                    {
                        PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                        object obj = cmd.ExecuteScalar();
                        cmd.Parameters.Clear();
                        if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
                        {
                            return null;
                        }
                        else
                        {
                            return obj;
                        }
                    }
                    catch (SqlException e)
                    {
                        throw new Exception(e.Message);
                    }
                }
            }
        }
        /// <summary>
        /// 执行查询语句，返回SqlDataReader
        /// </summary>
        /// <param name="strSQL">查询语句</param>
        /// <returns>SqlDataReader</returns>
        public static SqlDataReader ExecuteReader(string SQLString, params SqlParameter[] cmdParms)// 执行查询语句，返回SqlDataReader
        {
            SqlConnection connection = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            try
            {
                PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                SqlDataReader myReader = cmd.ExecuteReader();
                cmd.Parameters.Clear();
                return myReader;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }

        }
        /// <summary>
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="SQLString">查询语句</param>
        /// <returns>DataSet</returns>
        public static DataSet Query(string SQLString, params SqlParameter[] cmdParms)// 执行查询语句，返回DataSet
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand();
                PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    DataSet ds = new DataSet();
                    try
                    {
                        da.Fill(ds, "ds");
                        cmd.Parameters.Clear();
                    }
                    catch (SqlException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    return ds;
                }
            }
        }

        private static void PrepareCommand(SqlCommand cmd, SqlConnection conn, SqlTransaction trans, string cmdText, SqlParameter[] cmdParms)
        {
            if (conn.State != ConnectionState.Open)
                conn.Open();
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            if (trans != null)
                cmd.Transaction = trans;
            cmd.CommandType = CommandType.Text;//cmdType;
            if (cmdParms != null)
            {
                foreach (SqlParameter parm in cmdParms)
                    cmd.Parameters.Add(parm);
            }
        }

        #endregion

        #region 存储过程操作

        /// <summary>
        /// 执行存储过程
        /// </summary>
        /// <param name="storedProcName">存储过程名</param>
        /// <param name="parameters">存储过程参数</param>
        /// <returns>SqlDataReader</returns>
        public static SqlDataReader RunProcedure(string storedProcName, IDataParameter[] parameters)//执行存储过程,返回SqlDataReader
        {
            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataReader returnReader;
            connection.Open();
            SqlCommand command = BuildQueryCommand(connection, storedProcName, parameters);
            command.CommandType = CommandType.StoredProcedure;
            returnReader = command.ExecuteReader();
            return returnReader;
        }
        /// <summary>
        /// 执行存储过程
        /// </summary>
        /// <param name="storedProcName">存储过程名</param>
        /// <param name="parameters">存储过程参数</param>
        /// <param name="tableName">DataSet结果中的表名</param>
        /// <returns>DataSet</returns>
        public static DataSet RunProcedure(string storedProcName, IDataParameter[] parameters, string tableName) //执行存储过程,返回DataSet结果中的表名
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                DataSet dataSet = new DataSet();
                connection.Open();
                SqlDataAdapter sqlDA = new SqlDataAdapter();
                sqlDA.SelectCommand = BuildQueryCommand(connection, storedProcName, parameters);
                sqlDA.Fill(dataSet, tableName);
                connection.Close();
                return dataSet;
            }
        }
        /// <summary>
        /// 构建 SqlCommand 对象(用来返回一个结果集，而不是一个整数值)
        /// </summary>
        /// <param name="connection">数据库连接</param>
        /// <param name="storedProcName">存储过程名</param>
        /// <param name="parameters">存储过程参数</param>
        /// <returns>SqlCommand</returns>
        private static SqlCommand BuildQueryCommand(SqlConnection connection, string storedProcName, IDataParameter[] parameters)
        // 构建 SqlCommand 对象(用来返回一个结果集，而不是一个整数值)，返回SqlCommand
        {
            SqlCommand command = new SqlCommand(storedProcName, connection) { CommandType = CommandType.StoredProcedure };
            foreach (SqlParameter parameter in parameters)
            {
                command.Parameters.Add(parameter);
            }
            return command;
        }
        /// <summary>
        /// 执行存储过程，返回影响的行数		
        /// </summary>
        /// <param name="storedProcName">存储过程名</param>
        /// <param name="parameters">存储过程参数</param>
        /// <param name="rowsAffected">影响的行数</param>
        /// <returns></returns>
        public static int RunProcedure(string storedProcName, IDataParameter[] parameters, out int rowsAffected)
        // 执行存储过程，返回影响的行数
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                int result;
                connection.Open();
                SqlCommand command = BuildIntCommand(connection, storedProcName, parameters);
                rowsAffected = command.ExecuteNonQuery();
                result = (int)command.Parameters["ReturnValue"].Value;
                //Connection.Close();
                return result;
            }
        }
        /// <summary>
        /// 创建 SqlCommand 对象实例(用来返回一个整数值)	
        /// </summary>
        /// <param name="storedProcName">存储过程名</param>
        /// <param name="parameters">存储过程参数</param>
        /// <returns>SqlCommand 对象实例</returns>
        private static SqlCommand BuildIntCommand(SqlConnection connection, string storedProcName, IDataParameter[] parameters)
        // 创建 SqlCommand 对象实例(用来返回一个整数值)	
        {
            SqlCommand command = BuildQueryCommand(connection, storedProcName, parameters);
            command.Parameters.Add(new SqlParameter("ReturnValue",
                SqlDbType.Int, 4, ParameterDirection.ReturnValue,
                false, 0, 0, string.Empty, DataRowVersion.Default, null));
            return command;
        }
        #endregion

        #region 通用查询
        /// <summary>
        /// 根据条件精确查询某列是某个值的数据条数
        /// </summary>
        /// <param name="TableName">表名</param>
        /// <param name="ColumnName">列名</param>
        /// <param name="KeyValue">值</param>
        /// <returns></returns>
        public static int GetObjectCountByKeyValue(string TableName, string ColumnName, string KeyValue)
        {
            if (KeyValue != "" || KeyValue != null)
            {
                string SqlStr = string.Format("select * from {0} where {1} ='{2}'", TableName, ColumnName, KeyValue);
                DataTable dt = OpenTable(SqlStr);
                return dt.Rows.Count;
            }
            else
                return 0;
        }

        /// <summary>
        /// 根据条件精确查询某列是某个值的数据
        /// </summary>
        /// <param name="TableName">表名</param>
        /// <param name="ColumnName">列名</param>
        /// <param name="KeyValue">值</param>
        /// <returns></returns>
        public static DataTable GetObjectDataByKeyValue(string TableName, string ColumnName, string KeyValue)
        {
            DataTable ReturnTable = new DataTable();
            if (KeyValue.ToString() != "" || KeyValue != null)
            {
                string SqlStr = string.Format("select * from {0} where {1} ='{2}'", TableName, ColumnName, KeyValue);
                ReturnTable = OpenTable(SqlStr);
            }
            return ReturnTable;
        }

        /// <summary>
        /// 根据条件模糊查询某列是某个值的数据
        /// </summary>
        /// <param name="TableName">表名</param>
        /// <param name="ColumnName">列名</param>
        /// <param name="KeyValue">值</param>
        /// <returns></returns>
        public static DataTable GetObjectDataLikeKeyValue(string TableName, string ColumnName, string KeyValue)
        {
            DataTable ReturnTable = new DataTable();
            if (KeyValue.ToString() != "" || KeyValue != null)
            {
                string SqlStr = string.Format("select * from {0} where {1} like '%{2}%'", TableName, ColumnName, KeyValue);
                ReturnTable = OpenTable(SqlStr);
            }
            return ReturnTable;
        }

        /// <summary>
        /// 根据某列的值查询指定列的值
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <param name="columnName">列名</param>
        /// <param name="value">值</param>
        /// <param name="goalcolumn">目标列</param>
        /// <returns></returns>
        public static string GetStrValue(string TableName, string ColumnName, string Value, string GoalColumn)
        {
            //string ReturnStr = "";
            //string SqlStr = string.Format("select {0} from {1} where {2} = '{3}'", GoalColumn, TableName, ColumnName, Value);
            //DataTable dt = data.DBQuery.OpenTable1(string.Format("select {0} from {1} where {2} = '{3}'", GoalColumn, TableName, ColumnName, Value));
            //if (dt.Rows.Count > 0)
            //{
            //    ReturnStr = dt.Rows[0][0].ToString();
            //}
            //return ReturnStr;
            return "";
        }

        /// <summary>
        /// 分页返回DataTable
        /// </summary>
        /// <param name="SQLString">查询的sql语句</param>
        /// <param name="start">从第start行开始返回</param>
        /// <param name="count">共返回count行记录</param>
        /// <param name="tablename">返回DataSet中的表名</param>
        /// <returns>返回DataTable</returns>
        public static DataTable GetTableFY(string SQLString, int start, int count, string tablename)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                DataSet ds = new DataSet();

                try
                {
                    connection.Open();
                    SqlDataAdapter command = new SqlDataAdapter(SQLString, connection);
                    command.Fill(ds, start, count, tablename);
                }
                catch (SqlException ex)
                {
                    throw new Exception(ex.Message);
                }
                return ds.Tables[tablename];
            }
        }
        #endregion
    }
}
