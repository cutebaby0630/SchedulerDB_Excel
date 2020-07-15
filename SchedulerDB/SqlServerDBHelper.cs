using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using Dapper.Contrib.Extensions;
using System.ComponentModel.DataAnnotations.Schema;

namespace SqlServerHelper.Core
{
    public class SqlServerDBHelper
    {
        #region -- Property --

        /// <summary>
        /// 设置SqlServer的连接字符串
        /// </summary>
        public static readonly string _connString = "Data Source={0};DataBase={1};User ID={2};Password={3}";

        private static readonly int _connectionTimeout = 600;

        private string _connectingString;

        /// <summary>
        /// The data source
        /// </summary>
        private String _dataSource, _dataBase, _userID, _userPwd;

        /// <summary>
        /// The connection
        /// </summary>
        private SqlConnection _connection;

        /// <summary>
        /// Gets a value indicating whether this instance is connected.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is connected; otherwise, <c>false</c>.
        /// </value>
        public bool IsConnected
        {
            get { return (_connection != null && _connection.State == ConnectionState.Open); }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is closed.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is closed; otherwise, <c>false</c>.
        /// </value>
        public bool IsClosed
        {
            get { return (_connection == null || _connection.State == ConnectionState.Closed); }
        }

        /// <summary>
        /// Gets or sets the last error MSG.
        /// </summary>
        /// <value>
        /// The last error MSG.
        /// </value>
        public String LastErrorMsg { get; set; }

        #endregion

        #region -- Constructor --

        /// <summary>
        /// 禁止实例化
        /// </summary>
        private SqlServerDBHelper()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlServerDBHelper"/> class.
        /// </summary>
        /// <param name="dataSource">The data source.</param>
        /// <param name="dataBase">The data base.</param>
        /// <param name="userID">The user identifier.</param>
        /// <param name="userPwd">The user password.</param>
        public SqlServerDBHelper(String dataSource, String dataBase, String userID, String userPwd)
        {
            _dataSource = dataSource;
            _dataBase = dataBase;
            _userID = userID;
            _userPwd = userPwd;

            _connectingString = String.Format(_connString, _dataSource, _dataBase, _userID, _userPwd);
            _connection = GetSqlDbConnection();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlServerDBHelper"/> class.
        /// </summary>
        /// <param name="connectString">The connect string.</param>
        public SqlServerDBHelper(String connectString)
        {
            _connectingString = connectString;
            _connection = GetSqlDbConnection();
        }

        /// <summary>
        /// Finalizes an instance of the <see cref="SqlServerDBHelper"/> class.
        /// </summary>
        ~SqlServerDBHelper()
        {
            if (!IsClosed)
            {
                DisConnectToDB();
            }
        }

        #endregion

        #region -- Private DbConnection --

        /// <summary>
        /// 建立資料庫連線
        /// </summary>
        /// <returns>資料庫連線</returns>
        private IDbConnection GetDbConnection()
        {
            IDbConnection dBConnection = null;

            try
            {

                dBConnection = new SqlConnection(_connectingString);

                if (dBConnection.State != ConnectionState.Open)
                {
                    dBConnection.Open();
                }
            }
            catch (SqlException myEx)
            {
                _connection = null;
                LastErrorMsg += myEx.ToString();
            }
            catch (Exception ex)
            {
                _connection = null;
                LastErrorMsg += ex.ToString();
            }

            return dBConnection;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="readOnlyConnection"></param>
        /// <returns></returns>
        private SqlConnection GetSqlDbConnection()
        {
            var dBConnection = new SqlConnection(_connectingString);

            if (dBConnection.State != ConnectionState.Open)
            {
                dBConnection.Open();
            }

            return dBConnection;
        }

        /// <summary>
        /// Dises the connect to database.
        /// </summary>
        /// <returns></returns>
        private bool DisConnectToDB()
        {
            //_connection.Close();
            return (IsClosed);
        }

        #endregion

        #region -- Public ExecuteNonQueryAsync --

        public async Task<int> ExecuteNonQueryAsync(string excuteSql, object param = null, bool enableTransaction = false, CommandType commandType = CommandType.Text)
        {
            try
            {
                using (IDbConnection con = GetDbConnection())
                {
                    if (!enableTransaction)
                    {
                        return await con.ExecuteAsync(excuteSql, param, null, _connectionTimeout, commandType).ConfigureAwait(false);
                    }

                    using (var trans = con.BeginTransaction())
                    {
                        try
                        {
                            var retValue = await con.ExecuteAsync(excuteSql, param, trans, _connectionTimeout, commandType).ConfigureAwait(false);
                            trans.Commit();
                            return retValue;
                        }
                        catch
                        {
                            trans.Rollback();
                            throw;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                LastErrorMsg = ex.ToString();

                return -1;
            }
        }

        /// <summary>
        /// ExecuteScalar，執行查詢並傳回第一個資料列的第一個資料行中查詢所傳回的結果
        /// </summary>
        /// <param name="excuteSql">SQL敘述</param>
        /// <param name="param">參數物件</param>
        /// <param name="enableTransaction">包Transaction執行</param>
        /// <param name="commandType">敘述類型</param>
        /// <returns>執行回覆結果</returns>
        public async Task<object> ExecuteScalarAsync(string excuteSql, object param = null, bool enableTransaction = false, CommandType commandType = CommandType.Text)
        {
            using (IDbConnection con = GetDbConnection())
            {
                if (!enableTransaction)
                {
                    return await con.ExecuteScalarAsync(excuteSql, param, null, _connectionTimeout, commandType).ConfigureAwait(false);
                }
                else
                {
                    using (var trans = con.BeginTransaction())
                    {
                        try
                        {
                            var result = await con.ExecuteScalarAsync(excuteSql, param, trans, _connectionTimeout, commandType).ConfigureAwait(false);
                            trans.Commit();
                            return result;
                        }
                        catch
                        {
                            trans.Rollback();
                            throw;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 執行交易 query
        /// </summary>
        /// <param name="taskList">任務清單</param>
        /// <returns></returns>
        public bool ExecuteTransactionQuery(params Action<IDbConnection, IDbTransaction>[] taskList)
        {
            using (IDbConnection con = GetDbConnection())
            {
                using (var transaction = con.BeginTransaction(IsolationLevel.ReadUncommitted))
                {
                    try
                    {
                        foreach (var act in taskList)
                        {
                            act(con, transaction);
                        }

                        transaction.Commit();
                        return true;
                    }
                    catch
                    {
                        // Make sure connection is opend, prevent to throw ZombieCheck exception
                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }

                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        #endregion

        #region -- Public QueryAsync --

        /// <summary>
        /// 查詢資料
        /// </summary>
        /// <typeparam name="TReturn">回覆的資料類型</typeparam>
        /// <param name="querySql">SQL敘述</param>
        /// <param name="param">查詢參數物件</param>
        /// <param name="timeoutSecs">SQL執行Timeout秒數</param>
        /// <param name="readOnlyConnection">是否使用 Read Only Connetion</param>
        /// <param name="commandType">敘述類型</param>
        /// <returns>資料物件</returns>x
        public async Task<IEnumerable<TReturn>> QueryAsyncwithTimeoutTime<TReturn>(string querySql, object param = null, int timeoutSecs = 20, bool readOnlyConnection = false, CommandType commandType = CommandType.Text)
        {
            using (IDbConnection con = GetDbConnection())
            {
                return await con.QueryAsync<TReturn>(querySql, param, null, timeoutSecs, commandType).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// 查詢資料
        /// </summary>
        /// <typeparam name="TReturn">回覆的資料類型</typeparam>
        /// <param name="querySql">SQL敘述</param>
        /// <param name="param">查詢參數物件</param>
        /// <param name="readOnlyConnection">是否使用 Read Only Connetion</param>
        /// <param name="commandType">敘述類型</param>
        /// <returns>資料物件</returns>x
        public async Task<IEnumerable<TReturn>> QueryAsync<TReturn>(string querySql, object param = null, CommandType commandType = CommandType.Text)
        {
            using (IDbConnection con = GetDbConnection())
            {
                return await con.QueryAsync<TReturn>(querySql, param, null, _connectionTimeout, commandType).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// 查詢第一筆資料
        /// (無結果回傳Null)
        /// </summary>
        /// <typeparam name="TResult">回傳的資料型態</typeparam>
        /// <param name="querySql">SQL敘述</param>
        /// <param name="param">查詢參數</param>
        /// <param name="readOnlyConnection">是否使用 Read Only Connetion</param>
        /// <param name="commandType">敘述類型</param>
        /// <returns>資料物件</returns>
        public async Task<TResult> QueryFirstOrDefaultAsync<TResult>(string querySql, object param = null, CommandType commandType = CommandType.Text)
        {
            using (IDbConnection con = GetDbConnection())
            {
                return await con.QueryFirstOrDefaultAsync<TResult>(querySql, param, null, _connectionTimeout, commandType).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// 查詢全部資料
        /// </summary>
        /// <typeparam name="TResult">資料封裝的物件類型</typeparam>
        /// <returns>資料物件</returns>
        public async Task<IEnumerable<TResult>> QueryAll<TResult>() where TResult : class
        {
            using (IDbConnection con = GetDbConnection())
            {
                return await con.GetAllAsync<TResult>(null, _connectionTimeout).ConfigureAwait(false);
            }
        }


        #endregion

        #region -- Public InsertAsync --

        /// <summary>
        /// 新增單筆或多筆.
        /// </summary>
        /// <typeparam name="T">新增資料物件Type or IEnumable</typeparam>
        /// <param name="insertEntity">新增物件</param>
        /// <param name="enableTransaction">是否使用Transaction</param>
        /// <returns>The ID(primary key) of the newly inserted record if it is identity using the defined type, otherwise null</returns>
        public async Task<int> InsertAsync<T>(T insertEntity, bool enableTransaction = false) where T : class
        {
            using (IDbConnection con = GetDbConnection())
            {
                if (!enableTransaction)
                {
                    return await con.InsertAsync(insertEntity, null, _connectionTimeout).ConfigureAwait(false);
                }
                else
                {
                    using (var transaction = con.BeginTransaction(IsolationLevel.ReadUncommitted))
                    {
                        try
                        {
                            var insertResult = await con.InsertAsync(insertEntity, transaction, _connectionTimeout).ConfigureAwait(false);
                            transaction.Commit();
                            return insertResult;
                        }
                        catch (Exception)
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }
        }

        ///// <summary>
        ///// 新增主次關係資料 (自動將Parent Key帶入 Child的 foreign key)
        ///// </summary>
        ///// <typeparam name="TParent">Parent 只能限制一筆</typeparam>
        ///// <typeparam name="TChild">Child 可以單筆或多筆</typeparam>
        ///// <param name="parent">Parent Data</param>
        ///// <param name="child">Child Data</param>
        ///// <returns></returns>
        //public async Task<bool> InsertAsync<TParent, TChild>(TParent parent, TChild child)
        //    where TParent : class
        //    where TChild : class
        //{
        //    bool isOk = false;
        //    using (IDbConnection con = GetDbConnection())
        //    {
        //        using (var transaction = con.BeginTransaction(IsolationLevel.ReadUncommitted))
        //        {
        //            try
        //            {
        //                if (await con.InsertAsync(parent, transaction, _connectionTimeout).ConfigureAwait(false) > 0)
        //                {
        //                    SetParentKeytoChildRefKey(ref parent, ref child);
        //                    if (await con.InsertAsync(child, transaction, _connectionTimeout).ConfigureAwait(false) > 0)
        //                    {
        //                        transaction.Commit();
        //                        isOk = true;
        //                    }
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                transaction.Rollback();
        //                throw;
        //            }
        //        }

        //        return isOk;
        //    }
        //}


        #endregion

        #region -- Public UpdateAsync --

        /// <summary>
        /// 更新單筆資料
        /// </summary>
        /// <typeparam name="T">更新資料物件Type</typeparam>
        /// <param name="updateEntity">更新物件</param>
        /// <returns></returns>
        public async Task<bool> UpdateAsync<T>(T updateEntity) where T : class
        {
            using (IDbConnection con = GetDbConnection())
            {
                return await con.UpdateAsync(updateEntity, null, _connectionTimeout).ConfigureAwait(false);
            }
        }


        #endregion

        #region -- Public Delete Async --

        /// <summary>
        /// 刪除單筆或多筆資料
        /// </summary>
        /// <typeparam name="T">刪除單筆資料 or 刪除多筆資料</typeparam>
        /// <param name="deleteEntity">單筆資料 or 多筆資料</param>
        /// <returns></returns>
        public async Task<bool> DeleteAsync<T>(T deleteEntity) where T : class
        {
            using (IDbConnection con = GetDbConnection())
            {
                return await con.DeleteAsync(deleteEntity, null, _connectionTimeout).ConfigureAwait(false);
            }
        }
        #endregion

        #region -- Public FillTableAsync --

        /// <summary>
        /// Fills the table.
        /// </summary>
        /// <param name="sql">The SQL.</param>
        /// <returns></returns>
        public async Task<DataTable> FillTableAsync(String sql)
        {
            using (SqlCommand cmd = new SqlCommand(sql, _connection))
            {
                DataTable dt = new DataTable();
                if (_connection.State != ConnectionState.Open) {
                    _connection.Open();
                }
                var reader = await cmd.ExecuteReaderAsync().ConfigureAwait(false);

                dt.Load(reader);

                return dt;
            }
        }

        #endregion

        #region -- Public SqlBulkCopyAsync --

        /// <summary>
        /// 新增多筆資料 - 使用 SqlBulkCopy, 当需要插入的数据很多的时候使用
        /// </summary>
        /// <typeparam name="T">資料物件類別</typeparam>
        /// <param name="entities">資料物件集合</param>
        /// <returns></returns>
        public async Task<int> SqlBulkCopyAsync<T>(IEnumerable<T> entities)
        {
            int rowCnt = -1;

            try
            {
                using (IDbConnection conn = GetSqlDbConnection())
                {
                    var tableAtt = (System.ComponentModel.DataAnnotations.Schema.TableAttribute)typeof(T).GetCustomAttributes(typeof(System.ComponentModel.DataAnnotations.Schema.TableAttribute), true).FirstOrDefault();
                    // default table name
                    string tableName = (tableAtt == null) ? typeof(T).Name : tableAtt.Name;

                    string sql = $"SELECT COUNT(*) FROM {tableName};";

                    var countStart = Convert.ToInt32(conn.ExecuteScalar(sql));

                    using (var bulkCopy = new SqlBulkCopy(conn as SqlConnection))
                    {

                        // 資料實體對應的資料表名稱;
                        bulkCopy.DestinationTableName = tableName;

                        var table = new DataTable();

                        var properties = typeof(T).GetProperties();

                        foreach (var prop in properties)
                        {
                            var dc = new DataColumn(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                            bulkCopy.ColumnMappings.Add(dc.ColumnName, dc.ColumnName); // 用ColumnName強制對應，避免 Table欄位 與 entities 的順序不一致
                            table.Columns.Add(dc);
                        }

                        foreach (T item in entities)
                        {
                            DataRow row = table.NewRow();
                            foreach (var prop in properties)
                            {
                                row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                            }

                            table.Rows.Add(row);
                        }

                        await bulkCopy.WriteToServerAsync(table).ConfigureAwait(false);

                        var countEnd = Convert.ToInt32(conn.ExecuteScalar(sql));

                        rowCnt = countEnd - countStart;
                    }
                }
            }
            catch (Exception ex)
            {
                LastErrorMsg = ex.ToString();

                rowCnt = -1;
            }

            return rowCnt;
        }

        private async Task<int> SqlBulkCopyAsync(string tableName, DataTable dt)
        {
            int rowCnt = dt.Rows.Count;

            try
            {
                using (SqlConnection conn = GetSqlDbConnection())
                {
                    using (SqlTransaction trans = conn.BeginTransaction())
                    {
                        //SqlBulkCopy批次處理新增 沒有檢驗比對處理
                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(_connection, SqlBulkCopyOptions.KeepIdentity, trans))
                        {
                            bulkCopy.DestinationTableName = tableName;
                            await bulkCopy.WriteToServerAsync(dt).ConfigureAwait(false);
                        }

                        trans.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                LastErrorMsg = ex.ToString();

                rowCnt = -1;
            }

            return rowCnt;
        }

        #endregion

        
    }
}
