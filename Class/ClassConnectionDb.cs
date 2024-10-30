using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace Class
{
    public class ClassConnectionDb : IDisposable
    {
        static public string ConnectionString()
        {
            return new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionConfig")["ConnString"] ?? "";
        }
        public SqlConnection conn;
        public SqlTransaction trans;
        public void OpenConnection()
        {
            if (conn == null)
            {
                conn = new SqlConnection(ConnectionString());
            }

            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
        }
        public void CloseConnection()
        {
            if (conn.State != ConnectionState.Closed)
            {
                conn.Close();
                conn.Dispose();
            }
        }
        public void BeginTransaction()
        {
            if (trans == null)
            {
                trans = conn.BeginTransaction();
            }
        }
        public void Commit()
        {
            if (trans != null)
            {
                trans.Commit();
            }
        }
        public void Rollback()
        {
            if (trans != null)
            {
                trans.Rollback();
            }
        }
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    trans?.Dispose();
                    conn?.Dispose();
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        public DataSet ExecuteAdapter(SqlCommand cmd)
        {
            if (cmd.CommandType != CommandType.StoredProcedure)
            {
                cmd.CommandType = CommandType.Text;
            }
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);

            if (ds != null)
            {
                if (ds.Tables.Count > 0)
                {
                    foreach (DataColumn column in ds.Tables[0].Columns)
                    {
                        column.ColumnName = column.ColumnName.ToLower();
                    }
                }
            }
            return ds;
        }

        private bool IsAuthorizedRole(string userName, string roleType)
        {
            // ตรวจสอบว่า userName และ roleType มีค่าไม่เป็นค่าว่างหรือไม่
            if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(roleType))
            {
                return false; // ถ้า userName หรือ roleType เป็นค่าว่าง ให้ถือว่าไม่ได้รับอนุญาต
            }

            // ตรวจสอบสิทธิ์บทบาท (เฉพาะ admin, approver, employee เท่านั้นที่อนุญาต ยกเว้น reviewer)
            var allowedRoles = new List<string> { "admin", "approver", "employee" }; // รายการบทบาทที่อนุญาต 
            if (allowedRoles.Contains(roleType))
            {
                return true; // หาก roleType อยู่ในรายการที่อนุญาต ให้อนุญาต
            }
            else
            {
                return false; // หากไม่อยู่ในรายการที่อนุญาต ให้ปฏิเสธ
            }
        }

        public string ExecuteNonQuerySQL(SqlCommand sqlCommand, string user_name, string role_type)
        {
            string ret = "";
            string query = sqlCommand.CommandText;
            try
            {
                if (sqlCommand.CommandType != CommandType.StoredProcedure)
                {
                    sqlCommand.CommandType = CommandType.Text;
                } 
                if (!IsAuthorizedRole(user_name, role_type))
                {
                    return "User is not authorized to perform this action.";
                }
                else
                {
                    if (trans != null)
                    {
                        sqlCommand.Transaction = trans;
                    }
                    sqlCommand.ExecuteNonQuery();

                    ret = "true";
                }

            }
            catch (Exception ex)
            {
                ret = ex.ToString();
                // query error 
                foreach (SqlParameter p in sqlCommand.Parameters)
                {
                    string value = p.Value == null ? "NULL" : p.Value?.ToString() ?? "";
                    if (p.SqlDbType == SqlDbType.VarChar || p.SqlDbType == SqlDbType.Char)
                    {
                        value = $"'{value}'";
                    }
                    query = query.Replace(p.ParameterName, value);
                }
                //// Write the query to a text file
                //System.IO.File.WriteAllText("executed_query.txt", query);
            }
            return ret;
        }

        public static SqlParameter CreateSqlParameter(string parameterName, SqlDbType dbType, object value, int length = 0)
        {
            if (value == null)
            {
                return new SqlParameter(parameterName, dbType) { Value = DBNull.Value };
            }

            if (dbType == SqlDbType.Int && int.TryParse(value.ToString(), out int intValue))
            {
                return new SqlParameter(parameterName, dbType) { Value = intValue };
            }

            if (dbType == SqlDbType.Decimal && decimal.TryParse(value.ToString(), out decimal decimalValue))
            {
                return new SqlParameter(parameterName, dbType) { Value = decimalValue };
            }

            if (dbType == SqlDbType.DateTime && DateTime.TryParse(value.ToString(), out DateTime dateTimeValue))
            {
                return new SqlParameter(parameterName, dbType) { Value = dateTimeValue };
            }

            if (dbType == SqlDbType.VarChar)
            {
                if (length > 0)
                {
                    return new SqlParameter(parameterName, dbType, length) { Value = value.ToString() };
                }
                else
                {
                    return new SqlParameter(parameterName, dbType) { Value = value.ToString() };
                }
            }

            // Default case for other data types
            return new SqlParameter(parameterName, dbType) { Value = value.ToString() };
        }

        //static private void ChangeColumnNamesToLowerCase(DataTable table)
        //{
        //    foreach (DataColumn column in table.Columns)
        //    {
        //        column.ColumnName = column.ColumnName.ToLower();
        //    }
        //}

        //static public DataTable ExecuteAdapterSQLDataTable(string sqlStatement, List<SqlParameter>? parameters, string tableName = "Table1", bool isStoredProcedure = false)
        //{
        //    DataSet dssql = new DataSet();
        //    try
        //    {
        //        string connStrSQL = ConnectionString();

        //        if (string.IsNullOrEmpty(connStrSQL))
        //        {
        //            throw new ApplicationException("Connection string cannot be null or empty.");
        //        }

        //        if (string.IsNullOrEmpty(sqlStatement) || string.IsNullOrWhiteSpace(sqlStatement))
        //        {
        //            throw new ApplicationException("SQL statement cannot be null or empty.");
        //        }

        //        if (!isStoredProcedure && string.IsNullOrWhiteSpace(sqlStatement))
        //        {
        //            throw new ArgumentException("SQL statement cannot be null or empty.");
        //        }


        //        using (SqlConnection connsql = new SqlConnection(connStrSQL))
        //        {
        //            connsql.Open();
        //            try
        //            {
        //                //if (!isStoredProcedure)
        //                //{
        //                //ValidateSqlStatement(sqlStatement, parameters);
        //                // Check for common SQL injection patterns
        //                string[] suspiciousPatterns = new[] { "--", "/*", "*/", ";", "DROP", "TRUNCATE", "EXEC", "EXECUTE" };

        //                foreach (var pattern in suspiciousPatterns)
        //                {
        //                    if (sqlStatement.ToUpperInvariant().Contains(pattern))
        //                    {
        //                        throw new SecurityException($"Potential SQL injection detected: {pattern}");
        //                    }
        //                }

        //                //}
        //                using (SqlCommand cmd = new SqlCommand(sqlStatement, connsql))
        //                {
        //                    //cmd.CommandType = isStoredProcedure ? CommandType.StoredProcedure : CommandType.Text;
        //                    if (isStoredProcedure) { cmd.CommandType = CommandType.StoredProcedure; }
        //                    cmd.CommandTimeout = 300;

        //                    // เพิ่มพารามิเตอร์ถ้ามี
        //                    if (parameters != null && parameters.Count > 0)
        //                    {
        //                        //AddParametersToCommand(cmd, parameters); 
        //                        if (parameters != null && parameters.Any())
        //                        {
        //                            Boolean paramlist = false;
        //                            foreach (var param in parameters)
        //                            {
        //                                if (param != null && !cmd.Parameters.Contains(param.ParameterName))
        //                                {
        //                                    cmd.Parameters.Add(param);
        //                                    paramlist = true;
        //                                }
        //                            }
        //                            if (paramlist)
        //                            {
        //                                //if (!isStoredProcedure)
        //                                //{
        //                                //    //ValidateSqlStatement(sqlStatement, parameters);
        //                                //    // Check for common SQL injection patterns
        //                                //    string[] suspiciousPatterns = new[] { "--", "/*", "*/", ";", "DROP", "TRUNCATE", "EXEC", "EXECUTE" };

        //                                //    foreach (var pattern in suspiciousPatterns)
        //                                //    {
        //                                //        if (sqlStatement.ToUpperInvariant().Contains(pattern))
        //                                //        {
        //                                //            throw new SecurityException($"Potential SQL injection detected: {pattern}");
        //                                //        }
        //                                //    }

        //                                // Ensure all parameters in the SQL statement are in the parameters list
        //                                var parameterNames = new Regex(@"@\w+").Matches(sqlStatement).Cast<Match>().Select(m => m.Value).Distinct().ToList();
        //                                if (parameters != null)
        //                                {
        //                                    foreach (var paramName in parameterNames)
        //                                    {
        //                                        if (!parameters.Any(p => p.ParameterName.Equals(paramName, StringComparison.OrdinalIgnoreCase)))
        //                                        {
        //                                            throw new SecurityException($"SQL parameter {paramName} is not provided in the parameters list");
        //                                        }
        //                                    }
        //                                }
        //                                else if (parameterNames.Any())
        //                                {
        //                                    throw new SecurityException("SQL statement contains parameters but no parameters were provided");
        //                                }
        //                                //}
        //                                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
        //                                {
        //                                    da.Fill(dssql, tableName);
        //                                }
        //                            }
        //                            else { throw new ArgumentException("Parameters must be provided for a parameterized SQL query."); }
        //                        }
        //                        else if (cmd.CommandText.Contains("@"))
        //                        {
        //                            throw new ArgumentException("Parameters must be provided for a parameterized SQL query.");
        //                        }
        //                    }
        //                    else
        //                    {
        //                        //if (!isStoredProcedure)
        //                        //{
        //                        //    string[] suspiciousPatterns = new[] { "--", "/*", "*/", ";", "DROP", "TRUNCATE", "EXEC", "EXECUTE" };

        //                        //    foreach (var pattern in suspiciousPatterns)
        //                        //    {
        //                        //        if (sqlStatement.ToUpperInvariant().Contains(pattern))
        //                        //        {
        //                        //            throw new SecurityException($"Potential SQL injection detected: {pattern}");
        //                        //        }
        //                        //    }
        //                        //}
        //                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
        //                        {
        //                            da.Fill(dssql, tableName);
        //                        }
        //                    }
        //                }
        //            }
        //            catch (SqlException ex)
        //            {
        //                throw new ApplicationException("An error occurred during the SQL operation.", ex);
        //            }
        //            catch (Exception ex)
        //            {
        //                throw new ApplicationException("An unexpected error occurred while executing the SQL command.", ex);
        //            }
        //        }
        //    }
        //    catch (Exception ex_function)
        //    {
        //        throw new ApplicationException("Function An unexpected error occurred while executing the SQL command.", ex_function);
        //    }

        //    if (dssql.Tables.Count > 0)
        //    {
        //        ChangeColumnNamesToLowerCase(dssql.Tables[0]);
        //        dssql.Tables[0].TableName = tableName;
        //        return dssql.Tables[0];
        //    }
        //    else
        //    {
        //        return new DataTable(tableName);
        //    }
        //}


        //public string ExecuteNonQuerySQLTrans(string user_name, string role_type, string sqlStatement, List<SqlParameter> parameters, SqlConnection conn, SqlTransaction? trans = null, bool isStoredProcedure = false)
        //{
        //    // ตรวจสอบสิทธิ์ก่อนดำเนินการ
        //    if (!ClassLogin.IsAuthorized(user_name))
        //    {
        //        return "User is not authorized to perform this action.";
        //    }

        //    // ตรวจสอบว่า SQL statement ไม่ว่างเปล่า
        //    if (string.IsNullOrEmpty(sqlStatement))
        //    {
        //        return "SQL statement cannot be null or empty.";
        //    }

        //    // ตรวจสอบว่า SqlConnection ไม่เป็น null
        //    if (conn == null)
        //    {
        //        return "Database connection cannot be null.";
        //    }
        //    try
        //    {
        //        //if (!isStoredProcedure)
        //        //{
        //        //ValidateSqlStatement(sqlStatement, parameters);
        //        // Check for common SQL injection patterns
        //        string[] suspiciousPatterns = new[] { "--", "/*", "*/", ";", "DROP", "TRUNCATE", "EXEC", "EXECUTE" };

        //        foreach (var pattern in suspiciousPatterns)
        //        {
        //            if (sqlStatement.ToUpperInvariant().Contains(pattern))
        //            {
        //                throw new SecurityException($"Potential SQL injection detected: {pattern}");
        //            }
        //        }
        //        //}
        //        using (SqlCommand cmd = new SqlCommand(sqlStatement, conn, trans))
        //        {
        //            cmd.CommandTimeout = 300;
        //            if (isStoredProcedure) { cmd.CommandType = CommandType.StoredProcedure; }
        //            try
        //            {
        //                // เพิ่มพารามิเตอร์ถ้ามี
        //                if (parameters != null && parameters.Count > 0)
        //                {
        //                    //AddParametersToCommand(cmd, parameters);
        //                    if (parameters != null && parameters.Any())
        //                    {
        //                        Boolean paramlist = false;
        //                        foreach (var param in parameters)
        //                        {
        //                            if (param != null && !cmd.Parameters.Contains(param.ParameterName))
        //                            {
        //                                cmd.Parameters.Add(param);
        //                                paramlist = true;
        //                            }
        //                        }
        //                        if (paramlist)
        //                        {
        //                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
        //                            {
        //                                return "User is not authorized to perform this action.";
        //                            }
        //                            else
        //                            {
        //                                // รันคำสั่ง SQL
        //                                var iret = cmd.ExecuteNonQuery();
        //                                //if (iret > 0) { return "true"; } else { return "false"; } 
        //                                return "true";
        //                            }
        //                        }
        //                        else { return "Parameters must be provided for a parameterized SQL query."; }
        //                    }
        //                    else if (cmd.CommandText.Contains("@"))
        //                    {
        //                        return "Parameters must be provided for a parameterized SQL query.";
        //                    }
        //                }
        //                else
        //                {
        //                    if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
        //                    {
        //                        return "User is not authorized to perform this action.";
        //                    }
        //                    else
        //                    {
        //                        // รันคำสั่ง SQL
        //                        cmd.ExecuteNonQuery();
        //                        return "true";
        //                    }
        //                }
        //            }
        //            catch (SqlException ex)
        //            {
        //                return $"error: SQL operation failed. {ex.Message}";
        //            }
        //            catch (Exception ex)
        //            {
        //                return $"error: An unexpected error occurred. {ex.Message}";
        //            }
        //            return "true";
        //        }
        //    }
        //    catch (Exception ex_function)
        //    {
        //        return "Function An unexpected error occurred while executing the SQL command." + ex_function.Message.ToString();
        //    }
        //}


    }
}
