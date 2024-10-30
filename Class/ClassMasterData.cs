
using Model;
using dotnet_epha_api.Class;
using dotnet6_epha_api.Class;
using Newtonsoft.Json;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

namespace Class
{
    public class ClassMasterData
    {
        private string sqlstr = "";
        private string jsper = "";
        private ClassFunctions cls = new ClassFunctions();
        private ClassJSON cls_json = new ClassJSON();
        private ClassConnectionDb _conn = new ClassConnectionDb();
        private ClassHazop clshazop = new ClassHazop();
        private DataSet dsData = new DataSet();
        private DataTable dt = new DataTable();
        private DataTable dtma = new DataTable();

        #region function
        string[] sMonth = ("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC").Split(',');
        private object ConvertToDBNull(object value)
        {
            return value == null || value == DBNull.Value ? DBNull.Value : value;
        }
        private object ConvertToIntOrDBNull(object value)
        {
            try
            {
                return int.TryParse(value.ToString(), out int result) ? (object)result : DBNull.Value;
            }
            catch { return DBNull.Value; }
        }
        private static DataTable refMsg(string status, string remark)
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_ref");
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            return dtMsg;
        }
        private static DataTable refMsg(string status, string remark, string seq_new)
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_new");
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            dtMsg.Rows[0]["seq_new"] = seq_new;
            return dtMsg;
        }
        private DataTable ConvertDStoDT(DataSet _ds, string _table_name)
        {
            DataTable _dt = new DataTable();
            try
            {
                _dt = dsData.Tables[_table_name].Copy();
            }
            catch { }
            return _dt;
        }
        private int get_max(string table_name, string Neverkey = "")
        {
            if (string.IsNullOrEmpty(table_name))
            {
                return 0;
            }
            try
            {

                // ตรวจสอบชื่อ table ให้มีเฉพาะอักขระที่ปลอดภัย
                if (!Regex.IsMatch(table_name, @"^[a-zA-Z0-9_]+$"))
                {
                    throw new ArgumentException("Invalid table name format.");
                }

                // เรียกใช้ stored procedure โดยใช้ชื่อ table
                List<SqlParameter> parameters = new List<SqlParameter>();

                parameters.Add(new SqlParameter("@TableName", SqlDbType.NVarChar) { Value = table_name });
                parameters.Add(new SqlParameter("@NextId", SqlDbType.Int) { Direction = ParameterDirection.Output });
                if (!string.IsNullOrEmpty(Neverkey))
                {
                    parameters.Add(new SqlParameter("@Neverkey", SqlDbType.NVarChar) { Value = Neverkey });
                }

                #region Execute to Datable
                try
                {
                    _conn = new ClassConnectionDb();
                    _conn = new ClassConnectionDb(); _conn.OpenConnection();
                    try
                    {
                        var command = _conn.conn.CreateCommand();
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = "usp_GetNextId";
                        //command.Parameters.Add(":costcenter", costcenter); 
                        if (parameters != null && parameters?.Count > 0)
                        {
                            foreach (var _param in parameters)
                            {
                                if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                {
                                    command.Parameters.Add(_param);
                                }
                            }
                        }
                        dt = new DataTable();
                        dt = _conn.ExecuteAdapter(command).Tables[0];
                        dt.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable



                return Convert.ToInt32(parameters[1].Value);  // ค่า @NextId ที่ถูกส่งออกมา
            }
            catch
            {
                return 0;
            }
        }



        private void ConvertJSONListresultToDataSet(string user_name, string role_type, ref string msg, ref string ret, ref DataSet dsData, DataMasterListModel param)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name)) { msg = "User is not authorized to perform this action."; ret = "Error"; return; }

            string table_name = param.json_name ?? "";
            jsper = param.json_data ?? "";
            if (jsper == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                cls_json = new ClassJSON();
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(user_name, role_type, jsper);
                if (dt == null) { dt = new DataTable(); }
                if (dt != null)
                {
                    dt.TableName = table_name;
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

        }
        private void ConvertJSONresultToData(string user_name, string role_type, ref string msg, ref string ret, ref DataSet dsData, SetDataMasterModel param, string table_name = "data")
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name)) { msg = "User is not authorized to perform this action."; ret = "Error"; return; }

            jsper = param.json_data + "";
            if (jsper == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                cls_json = new ClassJSON();
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(user_name, role_type, jsper);
                if (dt == null) { dt = new DataTable(); }
                if (dt != null)
                {
                    dt.TableName = table_name;
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

        }
        private void ConvertJSONresultToDataManageuser(string user_name, string role_type, ref string msg, ref string ret, ref DataSet dsData, SetManageuser param)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name)) { msg = "User is not authorized to perform this action."; ret = "Error"; return; }

            jsper = param.json_register_account + "";
            if (jsper == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                cls_json = new ClassJSON();
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(user_name, role_type, jsper);
                if (dt == null) { dt = new DataTable(); }
                if (dt != null)
                {
                    dt.TableName = "register_account";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

        }
        private void ConvertJSONresultToDataAuthorizationSetting(string user_name, string role_type, ref string msg, ref string ret, ref DataSet dsData, SetAuthorizationSetting param)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name)) { msg = "User is not authorized to perform this action."; ret = "Error"; return; }

            jsper = param.json_role_type + "";
            if (jsper == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                cls_json = new ClassJSON();
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(user_name, role_type, jsper);
                if (dt == null) { dt = new DataTable(); }
                if (dt != null)
                {
                    dt.TableName = "role_type";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_menu_setting + "";
            try
            {
                cls_json = new ClassJSON();
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(user_name, role_type, jsper);
                if (dt == null) { dt = new DataTable(); }
                if (dt != null)
                {
                    dt.TableName = "menu_setting";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_role_setting + "";
            try
            {
                cls_json = new ClassJSON();
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(user_name, role_type, jsper);
                if (dt == null) { dt = new DataTable(); }
                if (dt != null)
                {
                    dt.TableName = "role_setting";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

        }
        public DataTable refMsgSaveMaster(string status, string remark, string seq_new)
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_new");
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            dtMsg.Rows[0]["seq_new"] = seq_new;
            return dtMsg;
        }
        private string MapPathFiles(string _folder)
        {
            return (Path.Combine(Directory.GetCurrentDirectory(), "") + _folder.Replace("~", ""));
        }
        public (DataRow[] rows, int iRows) FilterDataTable(DataTable dt, Dictionary<string, object> filterParameters)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                // คืนค่าเป็น array ว่าง ๆ และจำนวนแถวเป็น 0 ถ้า DataTable ไม่มีข้อมูล
                return (new DataRow[0], 0);
            }

            // ตรวจสอบว่า DataTable มีคอลัมน์ที่ตรงกับคีย์ใน filterParameters หรือไม่
            foreach (var key in filterParameters.Keys)
            {
                if (!dt.Columns.Contains(key))
                {
                    throw new ArgumentException($"Column '{key}' does not exist in the DataTable.");
                }
            }

            // ใช้ LINQ ในการกรองข้อมูล
            var filteredRows = dt.AsEnumerable()
                                 .Where(row =>
                                 {
                                     foreach (var param in filterParameters)
                                     {
                                         var columnValue = row[param.Key];

                                         // ตรวจสอบว่าค่าที่อยู่ในคอลัมน์ไม่เป็น null และตรงกับค่าใน filterParameters
                                         if (columnValue == DBNull.Value || columnValue.ToString() != param.Value.ToString())
                                         {
                                             return false;
                                         }
                                     }
                                     return true;
                                 })
                                 .ToArray();

            // คืนค่าแถวที่กรองแล้วและจำนวนแถว
            return (filteredRows, filteredRows.Length);
        }

        #endregion function

        #region Manage User
        public string get_manageuser(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหาการ Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = param.user_name;

            clshazop = new ClassHazop();
            clshazop.get_employee_list(false, ref dsData);

            // ดึงค่าขั้นสูงสุดของ epha_register_account
            int iMaxSeqRegister = get_max("epha_register_account");

            // สร้าง SQL query เพื่อดึงข้อมูลจาก epha_register_account
            string sqlstr = @"select a.*, 'update' as action_type, 0 as action_change
                      from epha_register_account a 
                      order by seq";

            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt.Rows.Count == 0)
            {
                // กรณีที่ไม่มีข้อมูลใน DataTable หรือเป็นใบงานใหม่
                DataRow newRow = dt.NewRow();
                newRow["seq"] = iMaxSeqRegister;
                newRow["id"] = iMaxSeqRegister;
                newRow["create_by"] = user_name;
                newRow["action_type"] = "insert";
                newRow["action_change"] = 0;
                dt.Rows.Add(newRow);

                dt.AcceptChanges();

                // เพิ่มลำดับ `iMaxSeqRegister`
                iMaxSeqRegister += 1;
            }

            // ตั้งชื่อ Table ว่า register_account และเพิ่มลงใน dsData
            dt.TableName = "register_account";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();

            // ตั้งค่า max_id สำหรับ register_account
            clshazop.set_max_id(ref dtma, "register_account", (iMaxSeqRegister + 1).ToString());
            dtma.TableName = "max";

            // เพิ่ม DataTable ที่เก็บค่า max ลงใน dsData
            dsData.Tables.Add(dtma.Copy());
            dsData.AcceptChanges();

            // ตั้งชื่อ DataSet
            dsData.DataSetName = "dsData";
            dsData.AcceptChanges();

            // แปลง DataSet เป็น JSON
            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        public string set_manageuser(SetManageuser param)
        {
            // ตรวจสอบค่า เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "");
            string role_type = ClassLogin.GetUserRoleFromDb(user_name); //(param.role_type + "");
            string json_register_account = (param.json_register_account + "");

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            SetManageuser param_def = new SetManageuser();
            param_def.user_name = user_name;
            param_def.role_type = role_type;
            param_def.json_register_account = json_register_account;

            ConvertJSONresultToDataManageuser(user_name, role_type, ref msg, ref ret, ref dsData, param_def);

            if (ret.ToLower() == "error") { }
            else
            {
                if (dsData != null)
                {
                    //set role_type,role_setting,menu_setting 
                    var iMaxSeqRoleRole = get_max("epha_register_account");
                    ret = set_register_account(user_name, role_type, ref dsData, ref iMaxSeqRoleRole);
                }

            }
            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }
        public string set_register_account(string user_name, string role_type, ref DataSet dsData, ref int seq_now)
        {
            string ret = "";

            #region update data 
            DataTable dt = ConvertDStoDT(dsData, "register_account");

            string connectionString = ClassConnectionDb.ConnectionString();


            using (ClassConnectionDb transaction = new ClassConnectionDb())
            {
                transaction.OpenConnection();
                transaction.BeginTransaction();

                try
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string action_type = row["action_type"]?.ToString() ?? "";
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (action_type == "insert")
                        {
                            sqlstr = @"
                            INSERT INTO EPHA_REGISTER_ACCOUNT (
                                SEQ, ID, REGISTER_TYPE, ACCEPT_STATUS, USER_NAME, USER_DISPLAYNAME, USER_EMAIL, USER_PASSWORD, USER_PASSWORD_CONFIRM,
                                CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY
                            ) VALUES (
                                @SEQ, @ID, @REGISTER_TYPE, @ACCEPT_STATUS, @USER_NAME, @USER_DISPLAYNAME, @USER_EMAIL, @USER_PASSWORD, @USER_PASSWORD_CONFIRM,
                                GETDATE(), NULL, @CREATE_BY, @UPDATE_BY
                            )";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@REGISTER_TYPE", SqlDbType.Int) { Value = row["REGISTER_TYPE"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ACCEPT_STATUS", SqlDbType.Int) { Value = row["ACCEPT_STATUS"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 50) { Value = row["USER_NAME"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.NVarChar, 4000) { Value = row["USER_DISPLAYNAME"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@USER_EMAIL", SqlDbType.NVarChar, 50) { Value = row["USER_EMAIL"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@USER_PASSWORD", SqlDbType.NVarChar, 50) { Value = row["USER_PASSWORD"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@USER_PASSWORD_CONFIRM", SqlDbType.NVarChar, 50) { Value = row["USER_PASSWORD_CONFIRM"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });

                            seq_now += 1;
                        }
                        else if (action_type == "update")
                        {
                            seq_now = Convert.ToInt32(row["SEQ"]);

                            sqlstr = @"
                            UPDATE EPHA_REGISTER_ACCOUNT SET
                                ACCEPT_STATUS = @ACCEPT_STATUS, USER_NAME = @USER_NAME, USER_DISPLAYNAME = @USER_DISPLAYNAME,
                                USER_EMAIL = @USER_EMAIL, USER_PASSWORD = @USER_PASSWORD, USER_PASSWORD_CONFIRM = @USER_PASSWORD_CONFIRM
                            WHERE SEQ = @SEQ AND ID = @ID";

                            parameters.Add(new SqlParameter("@ACCEPT_STATUS", SqlDbType.Int) { Value = row["ACCEPT_STATUS"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 50) { Value = row["USER_NAME"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.NVarChar, 4000) { Value = row["USER_DISPLAYNAME"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@USER_EMAIL", SqlDbType.NVarChar, 50) { Value = row["USER_EMAIL"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@USER_PASSWORD", SqlDbType.NVarChar, 50) { Value = row["USER_PASSWORD"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@USER_PASSWORD_CONFIRM", SqlDbType.NVarChar, 50) { Value = row["USER_PASSWORD_CONFIRM"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                        }
                        else if (action_type == "delete")
                        {
                            sqlstr = @"
                            DELETE FROM EPHA_REGISTER_ACCOUNT
                            WHERE SEQ = @SEQ AND ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            if (!string.IsNullOrEmpty(action_type))
                            {
                                if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                {
                                    ret = "User is not authorized to perform this action.";
                                    break;
                                }
                                else
                                {
                                    #region ExecuteNonQuerySQL Data
                                    var command = transaction.conn.CreateCommand();
                                    //command.CommandType = CommandType.StoredProcedure;
                                    command.CommandText = sqlstr;
                                    if (parameters != null && parameters?.Count > 0)
                                    {
                                        foreach (var _param in parameters)
                                        {
                                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                            {
                                                command.Parameters.Add(_param);
                                            }
                                        }
                                        //command.Parameters.AddRange(parameters?.ToArray());
                                    }
                                    ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                    #endregion  ExecuteNonQuerySQL Data
                                }
                                if (ret != "true") break;
                            }
                        }
                    }

                    if (ret == "true")
                    {
                        if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                            transaction.Commit();
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                    }
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    ret = "error: " + ex.Message;
                    return ret;
                }
            }

            #endregion update data 

            return ret;
        }

        #endregion Manage User

        #region AuthorizationSetting
        public string get_authorizationsetting(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = param.user_name;

            int iMaxSeqRoleType = 0;
            int iMaxSeqMenuSetting = 0;
            int iMaxSeqRoleSetting = 0;

            #region menu -> fix data
            string sqlstr = @"   select a.page_type, a.page_controller
                         , case when a.page_type in ('main') then 0 else 1 end disable_page
                         , 0 as choos_menu
                         , a.seq, a.name
                         from epha_m_menu a 
                         where a.active_type = 1 
                         order by a.page_type, a.seq";

            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable
            if (dt != null)
            {
                dt.AcceptChanges();
                dt.TableName = "menu";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }
            #endregion menu -> fix data

            #region role_type
            iMaxSeqRoleType = get_max("epha_m_role_type");

            sqlstr = @"  select a.*, 'update' as action_type, 0 as action_change 
                 from epha_m_role_type a 
                 order by seq";

            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable
            if (dt == null || dt.Rows.Count == 0)
            {
                // กรณีที่เป็นข้อมูลใหม่
                dt = new DataTable();
                DataRow newRow = dt.NewRow();
                newRow["seq"] = iMaxSeqRoleType;
                newRow["id"] = iMaxSeqRoleType;
                newRow["create_by"] = user_name;
                newRow["action_type"] = "insert";
                newRow["action_change"] = 0;
                dt.Rows.Add(newRow);

                dt.AcceptChanges();
                iMaxSeqRoleType += 1;
            }

            dt.AcceptChanges();
            dt.TableName = "role_type";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();
            #endregion role_type

            #region menu_setting
            iMaxSeqMenuSetting = get_max("epha_m_menu_setting");

            sqlstr = @"  select a.*, 'update' as action_type, 0 as action_change 
                 , 1 as choos_data 
                 from epha_m_menu_setting a 
                 order by id_role_group, seq ";

            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable
            if (dt == null || dt.Rows.Count == 0)
            {
                // กรณีที่เป็นข้อมูลใหม่
                dt = new DataTable();
                DataRow newRow = dt.NewRow();
                newRow["seq"] = iMaxSeqMenuSetting;
                newRow["id"] = iMaxSeqMenuSetting;
                newRow["id_role_group"] = iMaxSeqRoleType;
                newRow["create_by"] = user_name;
                newRow["action_type"] = "insert";
                newRow["action_change"] = 0;
                dt.Rows.Add(newRow);

                dt.AcceptChanges();
                iMaxSeqMenuSetting += 1;
            }

            dt.AcceptChanges();
            dt.TableName = "menu_setting";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();
            #endregion menu_setting

            #region role_setting
            iMaxSeqRoleSetting = get_max("epha_m_role_setting");

            sqlstr = @"  select a.*, 'update' as action_type, 0 as action_change
                 , emp.user_displayname
                 , 'assets/img/team/avatar.webp' as user_img
                 from epha_m_role_setting a 
                 left join vw_epha_person_details emp on lower(emp.user_name)  = lower(a.user_name) 
                 order by id_role_group, seq  ";

            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable
            if (dt == null || dt.Rows.Count == 0)
            {
                // กรณีที่เป็นข้อมูลใหม่
                dt = new DataTable();
                DataRow newRow = dt.NewRow();
                newRow["seq"] = iMaxSeqRoleSetting;
                newRow["id"] = iMaxSeqRoleSetting;
                newRow["id_role_group"] = iMaxSeqRoleType;
                newRow["create_by"] = user_name;
                newRow["action_type"] = "insert";
                newRow["action_change"] = 0;
                dt.Rows.Add(newRow);

                dt.AcceptChanges();
                iMaxSeqRoleSetting += 1;
            }

            dt.AcceptChanges();
            dt.TableName = "role_setting";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();
            #endregion role_setting

            #region add employee list and max values
            clshazop = new ClassHazop();
            clshazop.get_employee_list(false, ref dsData);

            clshazop.set_max_id(ref dtma, "role_type", (iMaxSeqRoleType + 1).ToString());
            clshazop.set_max_id(ref dtma, "menu_setting", (iMaxSeqMenuSetting + 1).ToString());
            clshazop.set_max_id(ref dtma, "role_setting", (iMaxSeqRoleSetting + 1).ToString());

            dtma.TableName = "max";
            dsData.Tables.Add(dtma.Copy());
            dsData.AcceptChanges();
            #endregion add employee list and max values

            dsData.DataSetName = "dsData";
            dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        public string set_authorizationsetting(SetAuthorizationSetting param)
        {

            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "true";
            string user_name = (param.user_name + "");
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_role_type = (param.json_role_type + "");
            string json_menu_setting = (param.json_menu_setting + "");
            string json_role_setting = (param.json_role_setting + "");

            SetAuthorizationSetting param_def = new SetAuthorizationSetting();
            param_def.user_name = user_name;
            param_def.role_type = role_type;
            param_def.json_role_type = json_role_type;
            param_def.json_menu_setting = json_menu_setting;
            param_def.json_role_setting = json_role_setting;

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            ConvertJSONresultToDataAuthorizationSetting(user_name, role_type, ref msg, ref ret, ref dsData, param_def);

            if (ret.ToLower() == "error") { goto Next_Line; }

            //set role_type,role_setting,menu_setting 
            var iMaxSeqRoleRole = get_max("epha_m_role_type");
            var iMaxSeqRoleSetting = get_max("epha_m_role_setting");
            var iMaxSeqMenuSetting = get_max("epha_m_menu_setting");

            if (dsData != null)
            {
                if (dsData?.Tables.Count > 0)
                {
                    try
                    {
                        using (ClassConnectionDb transaction = new ClassConnectionDb())
                        {
                            transaction.OpenConnection();
                            transaction.BeginTransaction();

                            try
                            {
                                //table role_setting, menu_setting => set id_role_group => action_type = insert only
                                ret = set_role_type(user_name, role_type, transaction, ref dsData, ref iMaxSeqRoleRole);
                                if (ret != "true") { goto Next_Line; }

                                ret = set_role_setting(user_name, role_type, transaction, ref dsData, ref iMaxSeqRoleSetting);
                                if (ret != "true") { goto Next_Line; }

                                ret = set_menu_setting(user_name, role_type, transaction, ref dsData, ref iMaxSeqMenuSetting);
                                if (ret != "true") { goto Next_Line; }

                                if (ret == "true")
                                {
                                    if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                                        transaction.Commit();
                                    }
                                    else
                                    {
                                        transaction.Rollback();
                                    }
                                }
                                else
                                {
                                    transaction.Rollback();
                                }
                            }
                            catch (Exception exTransaction)
                            {
                                transaction.Rollback();
                                ret = "error: " + exTransaction.Message;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ret = "error: " + ex.Message;
                    }
                }
            }

        Next_Line:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }
        public string set_role_type(string user_name, string role_type, ClassConnectionDb transaction, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            string ret = "true";

            try
            {
                DataTable dt = ConvertDStoDT(dsData, "role_type");
                DataTable dtrole_setting = ConvertDStoDT(dsData, "role_setting");//id_role_group
                DataTable dtmenu_setting = ConvertDStoDT(dsData, "menu_setting");

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = dt.Rows[i]["action_type"]?.ToString() ?? "";
                    string sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (action_type == "insert")
                    {
                        sqlstr = "INSERT INTO EPHA_M_ROLE_TYPE " +
                                 "(SEQ, ID, NAME, DESCRIPTIONS, DEFAULT_TYPE, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @NAME, @DESCRIPTIONS, @DEFAULT_TYPE, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 255) { Value = dt.Rows[i]["NAME"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DEFAULT_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["DEFAULT_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });

                        dt.Rows[i]["SEQ"] = seq_now;
                        dt.Rows[i]["ID"] = seq_now;
                        dt.AcceptChanges();

                        string seq_def = dt.Rows[i]["SEQ"]?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(seq_def))
                        {
                            if (dtrole_setting != null)
                            {
                                if (dtrole_setting?.Rows.Count > 0)
                                {
                                    for (int ii = 0; ii < dtrole_setting?.Rows.Count; ii++)
                                    {
                                        if (seq_def == (dtrole_setting?.Rows[ii]["id_role_group"]?.ToString() ?? ""))
                                        {
                                            dtrole_setting.Rows[ii]["id_role_group"] = seq_now.ToString();
                                        }
                                    }
                                    dtrole_setting.AcceptChanges();
                                }
                            }
                            if (dtmenu_setting != null)
                            {
                                if (dtmenu_setting?.Rows.Count > 0)
                                {
                                    for (int ii = 0; ii < dtmenu_setting?.Rows.Count; ii++)
                                    {
                                        if (seq_def == (dtmenu_setting?.Rows[ii]["id_role_group"]?.ToString() ?? ""))
                                        {
                                            dtmenu_setting.Rows[ii]["id_role_group"] = seq_now.ToString();
                                        }
                                    }
                                    dtmenu_setting.AcceptChanges();
                                }
                            }
                        }

                        seq_now += 1;

                    }
                    else if (action_type == "update")
                    {
                        sqlstr = "UPDATE EPHA_M_ROLE_TYPE SET " +
                                 "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, DEFAULT_TYPE = @DEFAULT_TYPE, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID";

                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 255) { Value = dt.Rows[i]["NAME"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DEFAULT_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["DEFAULT_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] ?? DBNull.Value });
                    }
                    else if (action_type == "delete")
                    {
                        sqlstr = "DELETE FROM EPHA_M_ROLE_TYPE WHERE SEQ = @SEQ AND ID = @ID";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] ?? DBNull.Value });
                    }

                    if (!string.IsNullOrEmpty(sqlstr))
                    {
                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                ret = "User is not authorized to perform this action.";
                                break;
                            }
                            else
                            {
                                #region ExecuteNonQuerySQL Data
                                var command = transaction.conn.CreateCommand();
                                //command.CommandType = CommandType.StoredProcedure;
                                command.CommandText = sqlstr;
                                if (parameters != null && parameters?.Count > 0)
                                {
                                    foreach (var _param in parameters)
                                    {
                                        if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                        {
                                            command.Parameters.Add(_param);
                                        }
                                    }
                                    //command.Parameters.AddRange(parameters?.ToArray());
                                }
                                ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                #endregion  ExecuteNonQuerySQL Data
                            }
                            if (ret != "true") break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        public string set_role_setting(string user_name, string role_type, ClassConnectionDb transaction, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "true";
            try
            {
                DataTable dt = ConvertDStoDT(dsData, "role_setting");

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = dt.Rows[i]["action_type"]?.ToString() ?? "";
                    string sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (action_type == "insert")
                    {
                        sqlstr = "INSERT INTO EPHA_M_ROLE_SETTING " +
                                 "(SEQ, ID_ROLE_GROUP, USER_NAME, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID_ROLE_GROUP, @USER_NAME, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                        parameters.Add(new SqlParameter("@ID_ROLE_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_ROLE_GROUP"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["USER_NAME"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });

                        seq_now += 1;
                    }
                    else if (action_type == "update")
                    {
                        sqlstr = "UPDATE EPHA_M_ROLE_SETTING SET " +
                                 "ID_ROLE_GROUP = @ID_ROLE_GROUP, USER_NAME = @USER_NAME, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ";

                        parameters.Add(new SqlParameter("@ID_ROLE_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_ROLE_GROUP"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["USER_NAME"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                    }
                    else if (action_type == "delete")
                    {
                        sqlstr = "DELETE FROM EPHA_M_ROLE_SETTING WHERE SEQ = @SEQ";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                    }

                    if (!string.IsNullOrEmpty(sqlstr))
                    {
                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                ret = "User is not authorized to perform this action.";
                                break;
                            }
                            else
                            {
                                #region ExecuteNonQuerySQL Data
                                var command = transaction.conn.CreateCommand();
                                //command.CommandType = CommandType.StoredProcedure;
                                command.CommandText = sqlstr;
                                if (parameters != null && parameters?.Count > 0)
                                {
                                    foreach (var _param in parameters)
                                    {
                                        if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                        {
                                            command.Parameters.Add(_param);
                                        }
                                    }
                                    //command.Parameters.AddRange(parameters?.ToArray());
                                }
                                ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                #endregion  ExecuteNonQuerySQL Data
                            }
                            if (ret != "true") break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        public string set_menu_setting(string user_name, string role_type, ClassConnectionDb transaction, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }


            string ret = "true";
            string sqlstr = "";

            try
            {
                DataTable dt = ConvertDStoDT(dsData, "menu_setting");

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = dt.Rows[i]["action_type"]?.ToString() ?? "";
                    sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (action_type == "insert")
                    {
                        sqlstr = "INSERT INTO EPHA_M_MENU_SETTING " +
                                 "(SEQ, ID_ROLE_GROUP, ID_MENU, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID_ROLE_GROUP, @ID_MENU, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                        parameters.Add(new SqlParameter("@ID_ROLE_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_ROLE_GROUP"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ID_MENU", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["ID_MENU"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });

                        seq_now += 1;
                    }
                    else if (action_type == "update")
                    {
                        sqlstr = "UPDATE EPHA_M_MENU_SETTING SET " +
                                 "ID_ROLE_GROUP = @ID_ROLE_GROUP, ID_MENU = @ID_MENU, DESCRIPTIONS = @DESCRIPTIONS, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ";

                        parameters.Add(new SqlParameter("@ID_ROLE_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_ROLE_GROUP"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ID_MENU", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["ID_MENU"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                    }
                    else if (action_type == "delete")
                    {
                        sqlstr = "DELETE FROM EPHA_M_MENU_SETTING WHERE SEQ = @SEQ";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                    }

                    if (!string.IsNullOrEmpty(sqlstr))
                    {
                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                ret = "User is not authorized to perform this action.";
                                break;
                            }
                            else
                            {
                                #region ExecuteNonQuerySQL Data
                                var command = transaction.conn.CreateCommand();
                                //command.CommandType = CommandType.StoredProcedure;
                                command.CommandText = sqlstr;
                                if (parameters != null && parameters?.Count > 0)
                                {
                                    foreach (var _param in parameters)
                                    {
                                        if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                        {
                                            command.Parameters.Add(_param);
                                        }
                                    }
                                    //command.Parameters.AddRange(parameters?.ToArray());
                                }
                                ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                #endregion  ExecuteNonQuerySQL Data
                            }
                            if (ret != "true") break;
                        }
                    }
                }

                ret = "true";
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        #endregion AuthorizationSetting

        #region Manage User
        public string get_master_contractlist(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            // สร้าง DataTable และ DataSet
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = param.user_name;

            // ดึงค่าขั้นสูงสุดของ epha_person_details
            int iMaxSeq = get_max("epha_person_details", "employeeid");

            #region Contract List
            string sqlstr = @" select employeeid as seq, a.*, 'update' as action_type, 0 as action_change 
                       from epha_person_details a 
                       where a.user_type = 'contract'
                       order by a.employeeid";

            // ตรวจสอบผลลัพธ์ที่คืนค่ามาจากฐานข้อมูล
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            // ถ้า dt เป็น null หรือไม่มีแถว ให้สร้างแถวใหม่
            if (dt == null || dt.Rows.Count == 0)
            {
                if (dt == null) // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }

                // กรณีที่เป็นใบงานใหม่
                DataRow newRow = dt.NewRow();
                newRow["seq"] = iMaxSeq;
                newRow["employeeid"] = iMaxSeq;
                newRow["user_type"] = "contract";
                newRow["create_by"] = user_name;
                newRow["action_type"] = "insert";
                newRow["action_change"] = 0;
                dt.Rows.Add(newRow);

                dt.AcceptChanges();

                // เพิ่มลำดับ iMaxSeq
                iMaxSeq += 1;
            }

            dt.AcceptChanges();

            // ตั้งชื่อ Table ว่า data และเพิ่มลงใน dsData
            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();
            #endregion Contract List

            // ตั้งค่า max_id สำหรับ seq
            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            // เพิ่ม DataTable ที่เก็บค่า max ลงใน dsData
            dsData.Tables.Add(dtma.Copy());
            dsData.AcceptChanges();

            // ตั้งชื่อ DataSet
            dsData.DataSetName = "dsData";
            dsData.AcceptChanges();

            // แปลง DataSet เป็น JSON
            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        public string set_master_contractlist(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            DataMasterListModel param_def = new DataMasterListModel();
            param_def.json_name = "data";
            param_def.json_data = param.json_data;
            ConvertJSONListresultToDataSet(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }
            else
            {
                iMaxSeq = get_max("epha_person_details");
                ret = set_contractlist(user_name, role_type, ref dsData, ref iMaxSeq);
            }


        Next_Line_Convert:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));

        }
        public string set_contractlist(string user_name, string role_type, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            string ret = "";
            string connectionString = ClassConnectionDb.ConnectionString();
            string sqlstr = "";

            try
            {
                using (ClassConnectionDb transaction = new ClassConnectionDb())
                {
                    transaction.OpenConnection();
                    transaction.BeginTransaction();

                    try
                    {
                        DataTable dt = ConvertDStoDT(dsData, "data");

                        for (int i = 0; i < dt?.Rows.Count; i++)
                        {
                            string action_type = dt.Rows[i]["action_type"]?.ToString() ?? "";
                            sqlstr = "";
                            List<SqlParameter> parameters = new List<SqlParameter>();

                            if (action_type == "insert")
                            {
                                sqlstr = "INSERT INTO EPHA_PERSON_DETAILS " +
                                         "(SEQ, ID, NO, NO_DEVIATIONS, NO_GUIDE_WORDS, DEVIATIONS, GUIDE_WORDS, PROCESS_DEVIATION, AREA_APPLICATION, PARAMETER, ACTIVE_TYPE, DEF_SELECTED, " +
                                         "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                         "VALUES (@SEQ, @ID, @NO, @NO_DEVIATIONS, @NO_GUIDE_WORDS, @DEVIATIONS, @GUIDE_WORDS, @PROCESS_DEVIATION, @AREA_APPLICATION, @PARAMETER, @ACTIVE_TYPE, @DEF_SELECTED, " +
                                         "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = dt.Rows[i]["NO"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@NO_DEVIATIONS", SqlDbType.Int) { Value = dt.Rows[i]["NO_DEVIATIONS"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@NO_GUIDE_WORDS", SqlDbType.Int) { Value = dt.Rows[i]["NO_GUIDE_WORDS"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@DEVIATIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DEVIATIONS"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@GUIDE_WORDS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["GUIDE_WORDS"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@PROCESS_DEVIATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PROCESS_DEVIATION"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@AREA_APPLICATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["AREA_APPLICATION"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@PARAMETER", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PARAMETER"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@DEF_SELECTED", SqlDbType.Int) { Value = dt.Rows[i]["DEF_SELECTED"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });

                                seq_now += 1;
                            }
                            else if (action_type == "update")
                            {
                                seq_now = Convert.ToInt32(dt.Rows[i]["seq"]?.ToString() ?? "0");

                                sqlstr = "UPDATE EPHA_M_GUIDE_WORDS SET " +
                                         "NO = @NO, NO_DEVIATIONS = @NO_DEVIATIONS, NO_GUIDE_WORDS = @NO_GUIDE_WORDS, " +
                                         "DEVIATIONS = @DEVIATIONS, GUIDE_WORDS = @GUIDE_WORDS, PROCESS_DEVIATION = @PROCESS_DEVIATION, AREA_APPLICATION = @AREA_APPLICATION, PARAMETER = @PARAMETER, " +
                                         "ACTIVE_TYPE = @ACTIVE_TYPE, DEF_SELECTED = @DEF_SELECTED, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                         "WHERE SEQ = @SEQ AND ID = @ID";

                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = dt.Rows[i]["NO"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@NO_DEVIATIONS", SqlDbType.Int) { Value = dt.Rows[i]["NO_DEVIATIONS"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@NO_GUIDE_WORDS", SqlDbType.Int) { Value = dt.Rows[i]["NO_GUIDE_WORDS"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@DEVIATIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DEVIATIONS"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@GUIDE_WORDS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["GUIDE_WORDS"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@PROCESS_DEVIATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PROCESS_DEVIATION"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@AREA_APPLICATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["AREA_APPLICATION"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@PARAMETER", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PARAMETER"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@DEF_SELECTED", SqlDbType.Int) { Value = dt.Rows[i]["DEF_SELECTED"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] ?? DBNull.Value });
                            }
                            else if (action_type == "delete")
                            {
                                sqlstr = "DELETE FROM EPHA_M_GUIDE_WORDS WHERE SEQ = @SEQ AND ID = @ID";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] ?? DBNull.Value });
                            }
                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                if (!string.IsNullOrEmpty(action_type))
                                {
                                    if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        ret = "User is not authorized to perform this action.";
                                        break;
                                    }
                                    else
                                    {
                                        #region ExecuteNonQuerySQL Data
                                        var command = transaction.conn.CreateCommand();
                                        //command.CommandType = CommandType.StoredProcedure;
                                        command.CommandText = sqlstr;
                                        if (parameters != null && parameters?.Count > 0)
                                        {
                                            foreach (var _param in parameters)
                                            {
                                                if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                                {
                                                    command.Parameters.Add(_param);
                                                }
                                            }
                                            //command.Parameters.AddRange(parameters?.ToArray());
                                        }
                                        ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                        #endregion  ExecuteNonQuerySQL Data
                                    }
                                    if (ret != "true") break;
                                }
                            }
                        }

                        if (ret == "true")
                        {
                            if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                                transaction.Commit();
                                //// Mark the transaction scope as complete
                                //scope.Complete();
                            }
                            else
                            {
                                transaction.Rollback();
                            }
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        ret = "error: " + ex.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        #endregion Manage User

        #region Company, Department and Sections

        //get_master_systemwide
        public string get_master_company(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            int iMaxSeq = get_max("epha_m_company");

            #region company
            sqlstr = @" select * from epha_m_company t order by id ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "company";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion company

            #region Departments
            sqlstr = @" select distinct departments as id, functions +'-'+ departments as name, lower(a.departments) as text_check 
                        from vw_epha_person_details a
                        where isnull(functions,'') <> '' and isnull(departments,'') <> ''
                        order by   functions +'-'+ departments";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "departments";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Departments

            #region Sections
            sqlstr = @"select distinct emp.sections as id, emp.sections as name,
                       emp.functions, emp.departments, emp.sections
                       from vw_epha_person_details emp
                       where isnull(emp.functions,'') <> '' and isnull(emp.departments,'') <> '' and isnull(emp.sections,'') <> '' 
                       order by emp.functions, emp.departments, emp.sections";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "sections";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Sections

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }
        public string get_master_area(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            // สร้าง DataTable และ DataSet
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = param.user_name;

            // ดึงค่าขั้นสูงสุดของ epha_m_area
            int iMaxSeq = get_max("epha_m_area");

            #region Area Process Unit
            string sqlstr = @" select a.*, 'update' as action_type, 0 as action_change  
                       from epha_m_area a 
                       order by a.id";

            // ตรวจสอบผลลัพธ์ที่คืนค่ามาจากฐานข้อมูล
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            // ถ้า dt เป็น null หรือไม่มีแถว ให้สร้างแถวใหม่
            if (dt == null || dt.Rows.Count == 0)
            {
                if (dt == null) // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }

                // กรณีที่เป็นใบงานใหม่
                DataRow newRow = dt.NewRow();
                newRow["seq"] = iMaxSeq;
                newRow["id"] = iMaxSeq;
                newRow["create_by"] = user_name;
                newRow["action_type"] = "insert";
                newRow["action_change"] = 0;
                dt.Rows.Add(newRow);

                dt.AcceptChanges();
            }

            // ตั้งชื่อ Table ว่า area และเพิ่มลงใน dsData
            dt.TableName = "area";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();
            #endregion Area Process Unit

            // ตั้งค่า max_id สำหรับ seq
            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            // เพิ่ม DataTable ที่เก็บค่า max ลงใน dsData
            dsData.Tables.Add(dtma.Copy());
            dsData.AcceptChanges();

            // ตั้งชื่อ DataSet
            dsData.DataSetName = "dsData";
            dsData.AcceptChanges();

            // แปลง DataSet เป็น JSON
            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        public string get_master_toc(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            int iMaxSeq = get_max("epha_m_area_complex");
            int id_company = 1;
            int id_area = 1;

            #region Plant
            sqlstr = @" select t.seq as id, t.plant as name from epha_m_company t order by t.plant ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "plant";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Plant

            #region Area Process Unit
            sqlstr = @" select a.id, a.name, a.name as area_check  
                        from epha_m_area a 
                        order by a.id";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "area";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Area Process Unit

            #region Complex
            sqlstr = @"select t.*, 'update' as action_type, 0 as action_change
                       from epha_m_area_complex t 
                       order by t.id";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;
                dt.Rows[0]["id_company"] = id_company;
                dt.Rows[0]["id_area"] = id_area;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;

                dt.AcceptChanges();
            }
            dt.TableName = "toc";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            #endregion Complex


            //company
            #region company
            sqlstr = @" select * from epha_m_company t order by id ";

            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "company";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion company



            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }
        public string get_master_unit(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            int iMaxSeq = get_max("epha_m_business_unit");
            int id_company = 1;
            int id_area = 1;
            int id_plant_area = 1;

            #region Plant
            sqlstr = @" select t.seq as id, t.plant as name from epha_m_company t order by t.plant ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "plant";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Plant

            #region Area Process Unit
            sqlstr = @" select a.id, a.name, a.name as area_check  
                        from epha_m_area a 
                        order by a.id";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "area";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Area Process Unit

            #region Complex
            sqlstr = @" select a.id, a.name, a.name as area_check, a.id_company, a.id_area
                        from epha_m_area_complex a 
                        order by a.id";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "toc";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Complex

            //company
            #region company
            sqlstr = @" select * from epha_m_company t order by id ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "company";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion company

            #region Unit No
            sqlstr = @"select t.*, 'update' as action_type, 0 as action_change
                       from epha_m_business_unit t 
                       order by t.id";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;
                dt.Rows[0]["id_company"] = id_company;
                dt.Rows[0]["id_area"] = id_area;
                dt.Rows[0]["id_plant_area"] = id_plant_area;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;

                dt.AcceptChanges();
            }
            dt.TableName = "unit";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            #endregion Unit No

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        public string set_master_systemwide(SetDataMasterModel param)
        {
            string user_name = (param.user_name + "");
            string table_name = param.page_name ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);

            string ret = "";
            string msg = "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            dsData = new DataSet();
            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param, table_name);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }

            try
            {
                using (ClassConnectionDb transaction = new ClassConnectionDb())
                {
                    transaction.OpenConnection();
                    transaction.BeginTransaction();

                    try
                    {
                        if (table_name == "area") { ret = set_master_area(user_name, role_type, transaction, dsData, table_name, ref ret); }
                        else if (table_name == "toc") { ret = set_master_toc(user_name, role_type, transaction, dsData, table_name, ref ret); }
                        else if (table_name == "unit") { ret = set_master_unit(user_name, role_type, transaction, dsData, table_name, ref ret); }

                        if (ret == "true")
                        {
                            if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                transaction.Commit();
                            }
                            else
                            {
                                transaction.Rollback();
                                ret = "error";
                            }
                        }
                        else
                        {
                            transaction.Rollback();
                            ret = "error";
                        }
                    }
                    catch (Exception exTransaction)
                    {
                        transaction.Rollback();
                        ret = "error: " + exTransaction.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }


        Next_Line_Convert:;
            string json = ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
            return json;
        }

        public string set_master_area(string user_name, string role_type, ClassConnectionDb transaction, DataSet dsData, string json_table_name, ref string ret)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }
            ret = "true";
            try
            {
                #region update data 
                int seq_now = Convert.ToInt32(get_max("epha_m_area").ToString() ?? "1");
                DataTable dt = ConvertDStoDT(dsData, json_table_name);

                foreach (DataRow row in dt.Rows)
                {
                    string action_type = row["action_type"]?.ToString() ?? "";
                    string sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (action_type == "insert")
                    {
                        #region insert 
                        sqlstr = "INSERT INTO EPHA_M_AREA (SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });

                        seq_now += 1;
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        seq_now = Convert.ToInt32(row["seq"]?.ToString() ?? "0");

                        #region update
                        sqlstr = "UPDATE EPHA_M_AREA SET " +
                                 "ID = @ID, NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ";

                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_M_AREA WHERE SEQ = @SEQ";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                        #endregion delete
                    }

                    if (!string.IsNullOrEmpty(sqlstr))
                    {
                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                ret = "User is not authorized to perform this action.";
                                break;
                            }
                            else
                            {
                                #region ExecuteNonQuerySQL Data
                                var command = transaction.conn.CreateCommand();
                                //command.CommandType = CommandType.StoredProcedure;
                                command.CommandText = sqlstr;
                                if (parameters != null && parameters?.Count > 0)
                                {
                                    foreach (var _param in parameters)
                                    {
                                        if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                        {
                                            command.Parameters.Add(_param);
                                        }
                                    }
                                    //command.Parameters.AddRange(parameters?.ToArray());
                                }
                                ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                #endregion  ExecuteNonQuerySQL Data
                            }
                            if (ret != "true") break;
                        }
                    }
                }
                #endregion update data 
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        public string set_master_toc(string user_name, string role_type, ClassConnectionDb transaction, DataSet dsData, string json_table_name, ref string ret)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }
            ret = "true";
            try
            {
                #region update data 
                int seq_now = Convert.ToInt32(get_max("epha_m_area_complex").ToString() ?? "1");
                DataTable dt = ConvertDStoDT(dsData, json_table_name);

                foreach (DataRow row in dt.Rows)
                {
                    string action_type = row["action_type"]?.ToString() ?? "";
                    string sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (action_type == "insert")
                    {
                        #region insert 
                        sqlstr = "INSERT INTO epha_m_area_complex (SEQ, ID, ID_COMPANY, ID_AREA, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_COMPANY, @ID_AREA, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                        parameters.Add(new SqlParameter("@ID_COMPANY", SqlDbType.Int) { Value = row["ID_COMPANY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ID_AREA", SqlDbType.Int) { Value = row["ID_AREA"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });

                        seq_now += 1;
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        seq_now = Convert.ToInt32(row["seq"]?.ToString() ?? "0");

                        #region update
                        sqlstr = "UPDATE epha_m_area_complex SET " +
                                 "ID = @ID, NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ";

                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM epha_m_area_complex WHERE SEQ = @SEQ";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                        #endregion delete
                    }

                    if (!string.IsNullOrEmpty(sqlstr))
                    {
                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                ret = "User is not authorized to perform this action.";
                                break;
                            }
                            else
                            {
                                #region ExecuteNonQuerySQL Data
                                var command = transaction.conn.CreateCommand();
                                //command.CommandType = CommandType.StoredProcedure;
                                command.CommandText = sqlstr;
                                if (parameters != null && parameters?.Count > 0)
                                {
                                    foreach (var _param in parameters)
                                    {
                                        if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                        {
                                            command.Parameters.Add(_param);
                                        }
                                    }
                                    //command.Parameters.AddRange(parameters?.ToArray());
                                }
                                ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                #endregion  ExecuteNonQuerySQL Data
                            }
                            if (ret != "true") break;
                        }
                    }
                }
                #endregion update data 
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        public string set_master_unit(string user_name, string role_type, ClassConnectionDb transaction, DataSet dsData, string json_table_name, ref string ret)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }
            ret = "true";
            try
            {
                int seq_now = Convert.ToInt32(get_max("EPHA_M_BUSINESS_UNIT").ToString() ?? "1");

                #region update data 
                DataTable dt = ConvertDStoDT(dsData, json_table_name);

                foreach (DataRow row in dt.Rows)
                {
                    string action_type = row["action_type"]?.ToString() ?? "";
                    string sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (action_type == "insert")
                    {
                        #region insert 
                        sqlstr = "INSERT INTO EPHA_M_BUSINESS_UNIT (SEQ, ID, ID_COMPANY, ID_AREA, ID_PLANT_AREA, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_COMPANY, @ID_AREA, @ID_PLANT_AREA, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                        parameters.Add(new SqlParameter("@ID_COMPANY", SqlDbType.Int) { Value = row["ID_COMPANY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ID_AREA", SqlDbType.Int) { Value = row["ID_AREA"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ID_PLANT_AREA", SqlDbType.Int) { Value = row["ID_PLANT_AREA"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });

                        seq_now += 1;
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        seq_now = Convert.ToInt32(row["seq"]?.ToString() ?? "0");

                        #region update
                        sqlstr = "UPDATE EPHA_M_BUSINESS_UNIT SET " +
                                 "ID = @ID, NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ";

                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_M_BUSINESS_UNIT WHERE SEQ = @SEQ";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                        #endregion delete
                    }

                    if (!string.IsNullOrEmpty(sqlstr))
                    {
                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                ret = "User is not authorized to perform this action.";
                                break;
                            }
                            else
                            {
                                #region ExecuteNonQuerySQL Data
                                var command = transaction.conn.CreateCommand();
                                //command.CommandType = CommandType.StoredProcedure;
                                command.CommandText = sqlstr;
                                if (parameters != null && parameters?.Count > 0)
                                {
                                    foreach (var _param in parameters)
                                    {
                                        if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                        {
                                            command.Parameters.Add(_param);
                                        }
                                    }
                                    //command.Parameters.AddRange(parameters?.ToArray());
                                }
                                ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                #endregion  ExecuteNonQuerySQL Data
                            }
                            if (ret != "true") break;
                        }
                    }
                }
                #endregion update data 
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }
            return ret;
        }

        #endregion  Company, Department and Sections

        #region HAZOP Module : Functional Location
        public string get_master_functionallocation(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            // สร้าง DataTable และ DataSet
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = param.user_name;

            // ดึงค่าขั้นสูงสุดของ epha_m_functional_location
            int iMaxSeq = get_max("epha_m_functional_location");

            #region Area Process Unit
            string sqlstr = @"select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                      from epha_m_functional_location a 
                      order by seq";

            // ตรวจสอบผลลัพธ์ที่คืนค่ามาจากฐานข้อมูล
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt.Rows.Count == 0)
            {
                if (dt == null) // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }

                // กรณีที่เป็นใบงานใหม่
                DataRow newRow = dt.NewRow();
                newRow["seq"] = iMaxSeq;
                newRow["id"] = iMaxSeq;
                newRow["create_by"] = user_name;
                newRow["action_type"] = "insert";
                newRow["action_change"] = 0;
                dt.Rows.Add(newRow);

                dt.AcceptChanges();

                // เพิ่มลำดับ iMaxSeq
                iMaxSeq += 1;
            }

            dt.AcceptChanges();

            // ตั้งชื่อ Table ว่า data และเพิ่มลงใน dsData
            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();
            #endregion Area Process Unit

            #region Drawing
            sqlstr = @"select a.* , 'update' as action_type, 0 as action_change
               from epha_m_drawing a
               where a.module = 'functional_location' 
               order by a.seq";

            // ตรวจสอบผลลัพธ์ที่คืนค่ามาจากฐานข้อมูล
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            int id_drawing = get_max("epha_m_drawing");

            if (dt == null || dt.Rows.Count == 0)
            {
                if (dt == null) // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }

                // กรณีที่เป็นใบงานใหม่
                DataRow newRow = dt.NewRow();
                newRow["seq"] = id_drawing;
                newRow["id"] = id_drawing;
                newRow["module"] = "functional_location";
                newRow["create_by"] = user_name;
                newRow["action_type"] = "insert";
                newRow["action_change"] = 0;
                dt.Rows.Add(newRow);

                dt.AcceptChanges();
            }

            // ตั้งชื่อ Table ว่า drawing และเพิ่มลงใน dsData
            dt.TableName = "drawing";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();
            #endregion Drawing

            // ตั้งค่า max_id สำหรับ seq และ drawing
            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            clshazop.set_max_id(ref dtma, "drawing", (id_drawing + 1).ToString());
            dtma.TableName = "max";

            // เพิ่ม DataTable ที่เก็บค่า max ลงใน dsData
            dsData.Tables.Add(dtma.Copy());
            dsData.AcceptChanges();

            // ตั้งชื่อ DataSet
            dsData.DataSetName = "dsData";
            dsData.AcceptChanges();

            // แปลง DataSet เป็น JSON
            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        public string set_master_functionallocation(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            DataMasterListModel param_def = new DataMasterListModel();
            param_def.json_name = "data";
            param_def.json_data = param.json_data;
            ConvertJSONListresultToDataSet(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }

            param_def = new DataMasterListModel();
            param_def.json_name = "drawing";
            param_def.json_data = param.json_drawing;
            ConvertJSONListresultToDataSet(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }

            iMaxSeq = get_max("epha_m_functional_location");
            int iMaxSeqDrawing = get_max("epha_m_drawing");


            if (dsData != null)
            {
                if (dsData?.Tables.Count > 0)
                {
                    try
                    {
                        using (ClassConnectionDb transaction = new ClassConnectionDb())
                        {
                            transaction.OpenConnection();
                            transaction.BeginTransaction();

                            try
                            {

                                ret = set_functional_location(user_name, role_type, transaction, ref dsData, ref iMaxSeq);
                                if (ret != "true") throw new Exception(ret);

                                ret = set_master_drawing(user_name, role_type, transaction, ref dsData, ref iMaxSeqDrawing);
                                if (ret != "true") throw new Exception(ret);

                                if (ret == "true")
                                {
                                    if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                                        transaction.Commit();
                                    }
                                    else
                                    {
                                        transaction.Rollback();
                                    }
                                }
                                else
                                {
                                    transaction.Rollback();
                                }
                            }
                            catch (Exception exTransaction)
                            {
                                transaction.Rollback();
                                ret = "error: " + exTransaction.Message;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ret = "error: " + ex.Message;
                    }
                }
            }


        Next_Line_Convert:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }

        public string set_functional_location(string user_name, string role_type, ClassConnectionDb transaction, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "true";

            try
            {
                dt = new DataTable();
                dt = ConvertDStoDT(dsData, "data");

                foreach (DataRow row in dt.Rows)
                {
                    string action_type = row["action_type"]?.ToString() ?? "";
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string sqlstr = "";

                    if (!string.IsNullOrEmpty(action_type))
                    {
                        if (action_type == "insert")
                        {
                            sqlstr = "insert into EPHA_M_FUNCTIONAL_LOCATION " +
                                     "(SEQ, ID, NOTIF_DATE, NOTIF_TIME, NOTIFICATION, ORDERS, TYP, P, PLS, FUNCTIONAL_LOCATION, DESCRIPTIONS_FUNC, DESCRIPTIONS, " +
                                     "MN_WK, PLNT, REPORTED_BY, REQUIRED_START, REQUIRED_END, USER_STATUS, DATE_UPDATE, ACTIVE_TYPE, CREATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "values (@SEQ, @ID, @NOTIF_DATE, @NOTIF_TIME, @NOTIFICATION, @ORDERS, @TYP, @P, @PLS, @FUNCTIONAL_LOCATION, @DESCRIPTIONS_FUNC, @DESCRIPTIONS, " +
                                     "@MN_WK, @PLNT, @REPORTED_BY, @REQUIRED_START, @REQUIRED_END, @USER_STATUS, @DATE_UPDATE, @ACTIVE_TYPE, getdate(), @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@NOTIF_DATE", SqlDbType.NVarChar, 4000) { Value = row["NOTIF_DATE"] });
                            parameters.Add(new SqlParameter("@NOTIF_TIME", SqlDbType.NVarChar, 4000) { Value = row["NOTIF_TIME"] });
                            parameters.Add(new SqlParameter("@NOTIFICATION", SqlDbType.NVarChar, 4000) { Value = row["NOTIFICATION"] });
                            parameters.Add(new SqlParameter("@ORDERS", SqlDbType.NVarChar, 4000) { Value = row["ORDERS"] });
                            parameters.Add(new SqlParameter("@TYP", SqlDbType.NVarChar, 4000) { Value = row["TYP"] });
                            parameters.Add(new SqlParameter("@P", SqlDbType.NVarChar, 4000) { Value = row["P"] });
                            parameters.Add(new SqlParameter("@PLS", SqlDbType.NVarChar, 4000) { Value = row["PLS"] });
                            parameters.Add(new SqlParameter("@FUNCTIONAL_LOCATION", SqlDbType.NVarChar, 4000) { Value = row["FUNCTIONAL_LOCATION"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS_FUNC", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS_FUNC"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                            parameters.Add(new SqlParameter("@MN_WK", SqlDbType.NVarChar, 4000) { Value = row["MN_WK"] });
                            parameters.Add(new SqlParameter("@PLNT", SqlDbType.NVarChar, 4000) { Value = row["PLNT"] });
                            parameters.Add(new SqlParameter("@REPORTED_BY", SqlDbType.NVarChar, 4000) { Value = row["REPORTED_BY"] });
                            parameters.Add(new SqlParameter("@REQUIRED_START", SqlDbType.NVarChar, 4000) { Value = row["REQUIRED_START"] });
                            parameters.Add(new SqlParameter("@REQUIRED_END", SqlDbType.NVarChar, 4000) { Value = row["REQUIRED_END"] });
                            parameters.Add(new SqlParameter("@USER_STATUS", SqlDbType.NVarChar, 4000) { Value = row["USER_STATUS"] });
                            parameters.Add(new SqlParameter("@DATE_UPDATE", SqlDbType.NVarChar, 4000) { Value = row["DATE_UPDATE"] });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });

                            seq_now++;
                        }
                        else if (action_type == "update")
                        {
                            sqlstr = "update EPHA_M_FUNCTIONAL_LOCATION set " +
                                     "NOTIF_DATE = @NOTIF_DATE, NOTIF_TIME = @NOTIF_TIME, NOTIFICATION = @NOTIFICATION, ORDERS = @ORDERS, TYP = @TYP, P = @P, PLS = @PLS, " +
                                     "FUNCTIONAL_LOCATION = @FUNCTIONAL_LOCATION, DESCRIPTIONS_FUNC = @DESCRIPTIONS_FUNC, DESCRIPTIONS = @DESCRIPTIONS, " +
                                     "MN_WK = @MN_WK, PLNT = @PLNT, REPORTED_BY = @REPORTED_BY, REQUIRED_START = @REQUIRED_START, REQUIRED_END = @REQUIRED_END, " +
                                     "USER_STATUS = @USER_STATUS, DATE_UPDATE = @DATE_UPDATE, ACTIVE_TYPE = @ACTIVE_TYPE, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                     "where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                            parameters.Add(new SqlParameter("@NOTIF_DATE", SqlDbType.NVarChar, 4000) { Value = row["NOTIF_DATE"] });
                            parameters.Add(new SqlParameter("@NOTIF_TIME", SqlDbType.NVarChar, 4000) { Value = row["NOTIF_TIME"] });
                            parameters.Add(new SqlParameter("@NOTIFICATION", SqlDbType.NVarChar, 4000) { Value = row["NOTIFICATION"] });
                            parameters.Add(new SqlParameter("@ORDERS", SqlDbType.NVarChar, 4000) { Value = row["ORDERS"] });
                            parameters.Add(new SqlParameter("@TYP", SqlDbType.NVarChar, 4000) { Value = row["TYP"] });
                            parameters.Add(new SqlParameter("@P", SqlDbType.NVarChar, 4000) { Value = row["P"] });
                            parameters.Add(new SqlParameter("@PLS", SqlDbType.NVarChar, 4000) { Value = row["PLS"] });
                            parameters.Add(new SqlParameter("@FUNCTIONAL_LOCATION", SqlDbType.NVarChar, 4000) { Value = row["FUNCTIONAL_LOCATION"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS_FUNC", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS_FUNC"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                            parameters.Add(new SqlParameter("@MN_WK", SqlDbType.NVarChar, 4000) { Value = row["MN_WK"] });
                            parameters.Add(new SqlParameter("@PLNT", SqlDbType.NVarChar, 4000) { Value = row["PLNT"] });
                            parameters.Add(new SqlParameter("@REPORTED_BY", SqlDbType.NVarChar, 4000) { Value = row["REPORTED_BY"] });
                            parameters.Add(new SqlParameter("@REQUIRED_START", SqlDbType.NVarChar, 4000) { Value = row["REQUIRED_START"] });
                            parameters.Add(new SqlParameter("@REQUIRED_END", SqlDbType.NVarChar, 4000) { Value = row["REQUIRED_END"] });
                            parameters.Add(new SqlParameter("@USER_STATUS", SqlDbType.NVarChar, 4000) { Value = row["USER_STATUS"] });
                            parameters.Add(new SqlParameter("@DATE_UPDATE", SqlDbType.NVarChar, 4000) { Value = row["DATE_UPDATE"] });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });
                        }
                        else if (action_type == "delete")
                        {
                            sqlstr = "delete from EPHA_M_FUNCTIONAL_LOCATION where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            if (!string.IsNullOrEmpty(action_type))
                            {
                                if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                {
                                    ret = "User is not authorized to perform this action.";
                                    break;
                                }
                                else
                                {
                                    #region ExecuteNonQuerySQL Data
                                    var command = transaction.conn.CreateCommand();
                                    //command.CommandType = CommandType.StoredProcedure;
                                    command.CommandText = sqlstr;
                                    if (parameters != null && parameters?.Count > 0)
                                    {
                                        foreach (var _param in parameters)
                                        {
                                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                            {
                                                command.Parameters.Add(_param);
                                            }
                                        }
                                        //command.Parameters.AddRange(parameters?.ToArray());
                                    }
                                    ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                    #endregion  ExecuteNonQuerySQL Data
                                }
                                if (ret != "true") break;
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }


            return ret;
        }

        public string set_master_drawing(string user_name, string role_type, ClassConnectionDb transaction, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "true";


            try
            {
                if (dsData.Tables["drawing"] != null)
                {
                    dt = new DataTable();
                    dt = dsData?.Tables["drawing"]?.Copy() ?? new DataTable();
                    dt.AcceptChanges();

                    foreach (DataRow row in dt.Rows)
                    {
                        string action_type = row["action_type"]?.ToString() ?? "";
                        List<SqlParameter> parameters = new List<SqlParameter>();
                        string sqlstr = "";

                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (action_type == "insert")
                            {
                                sqlstr = "insert into EPHA_M_DRAWING " +
                                         "(SEQ, ID, MODULE, NAME, DESCRIPTIONS, DOCUMENT_FILE_SIZE, DOCUMENT_FILE_PATH, DOCUMENT_FILE_NAME, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                         "values (@SEQ, @ID, @MODULE, @NAME, @DESCRIPTIONS, @DOCUMENT_FILE_SIZE, @DOCUMENT_FILE_PATH, @DOCUMENT_FILE_NAME, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@MODULE", SqlDbType.NVarChar, 4000) { Value = row["MODULE"] });
                                parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] });
                                parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                                parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = row["DOCUMENT_FILE_SIZE"] });
                                parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_PATH"] });
                                parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_NAME"] });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });

                                seq_now++;
                            }
                            else if (action_type == "update")
                            {
                                sqlstr = "update EPHA_M_DRAWING set " +
                                         "MODULE = @MODULE, NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, DOCUMENT_FILE_SIZE = @DOCUMENT_FILE_SIZE, " +
                                         "DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH, DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME, " +
                                         "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                         "where SEQ = @SEQ and ID = @ID";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                parameters.Add(new SqlParameter("@MODULE", SqlDbType.NVarChar, 4000) { Value = row["MODULE"] });
                                parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] });
                                parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                                parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = row["DOCUMENT_FILE_SIZE"] });
                                parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_PATH"] });
                                parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_NAME"] });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });
                            }
                            else if (action_type == "delete")
                            {
                                sqlstr = "delete from EPHA_M_DRAWING where SEQ = @SEQ and ID = @ID";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                            }

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                if (!string.IsNullOrEmpty(action_type))
                                {
                                    if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        ret = "User is not authorized to perform this action.";
                                        break;
                                    }
                                    else
                                    {
                                        #region ExecuteNonQuerySQL Data
                                        var command = transaction.conn.CreateCommand();
                                        //command.CommandType = CommandType.StoredProcedure;
                                        command.CommandText = sqlstr;
                                        if (parameters != null && parameters?.Count > 0)
                                        {
                                            foreach (var _param in parameters)
                                            {
                                                if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                                {
                                                    command.Parameters.Add(_param);
                                                }
                                            }
                                            //command.Parameters.AddRange(parameters?.ToArray());
                                        }
                                        ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                        #endregion  ExecuteNonQuerySQL Data
                                    }
                                    if (ret != "true") break;
                                }

                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }


            return ret;
        }


        #endregion HAZOP Module : Functional Location

        #region HAZOP Module : Guide Words 
        public string get_master_guidewords(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            #region master data in page

            sqlstr = @" select p.id, p.name 
                        from epha_m_parameter p 
                        where p.active_type = 1
                        order by p.id";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt != null)
            {
                dt.TableName = "parameter";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }

            sqlstr = @" select a.id, a.name 
                        , a.id_parameter, p.name as parameter
                        from epha_m_area_application a 
                        inner join epha_m_parameter p on a.id_parameter = p.id
                        where a.active_type = 1
                        order by a.id";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable
            if (dt != null)
            {
                dt.TableName = "area_application";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }



            #endregion master data in page

            int iMaxSeq = get_max("epha_m_guide_words");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_guide_words a 
                        order by seq ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();
            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            #region drawing 
            sqlstr = @" select a.* , 'update' as action_type, 0 as action_change
                        from epha_m_drawing a
                        where a.module = 'guide_words' ";
            sqlstr += " order by a.seq ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            int id_drawing = get_max("epha_m_drawing");

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_drawing;
                dt.Rows[0]["id"] = id_drawing;

                dt.Rows[0]["module"] = "guide_words";

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }
            dt.TableName = "drawing";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion drawing

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            clshazop.set_max_id(ref dtma, "drawing", (id_drawing + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_guidewords(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";

            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }


            DataMasterListModel param_def = new DataMasterListModel
            {
                json_name = "data",
                json_data = param.json_data
            };
            ConvertJSONListresultToDataSet(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { return cls_json.SetJSONresult(ClassFile.refMsg("Error", ret)); }

            param_def = new DataMasterListModel
            {
                json_name = "drawing",
                json_data = param.json_drawing
            };
            ConvertJSONListresultToDataSet(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { return cls_json.SetJSONresult(ClassFile.refMsg("Error", ret)); }

            iMaxSeq = get_max("epha_m_guide_words");
            int iMaxSeqDrawing = get_max("epha_m_drawing");

            using (ClassConnectionDb transaction = new ClassConnectionDb())
            {
                transaction.OpenConnection();
                transaction.BeginTransaction();

                try
                {
                    #region update data guidewords
                    if (dsData.Tables["data"] != null)
                    {
                        dt = new DataTable();
                        dt = ConvertDStoDT(dsData, "data");

                        for (int i = 0; i < dt?.Rows.Count; i++)
                        {
                            string action_type = (dt.Rows[i]["action_type"]?.ToString() ?? "").ToLower();
                            string sqlstr = "";
                            List<SqlParameter> parameters = new List<SqlParameter>();

                            if (action_type == "insert")
                            {
                                #region insert
                                sqlstr = "insert into EPHA_M_GUIDE_WORDS " +
                                         "(SEQ, ID, NO, NO_DEVIATIONS, NO_GUIDE_WORDS, DEVIATIONS, GUIDE_WORDS, PROCESS_DEVIATION, AREA_APPLICATION, PARAMETER, ACTIVE_TYPE, DEF_SELECTED, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                         "values (@SEQ, @ID, @NO, @NO_DEVIATIONS, @NO_GUIDE_WORDS, @DEVIATIONS, @GUIDE_WORDS, @PROCESS_DEVIATION, @AREA_APPLICATION, @PARAMETER, @ACTIVE_TYPE, @DEF_SELECTED, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = iMaxSeq });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = iMaxSeq });
                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = dt.Rows[i]["NO"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@NO_DEVIATIONS", SqlDbType.Int) { Value = dt.Rows[i]["NO_DEVIATIONS"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@NO_GUIDE_WORDS", SqlDbType.Int) { Value = dt.Rows[i]["NO_GUIDE_WORDS"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@DEVIATIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DEVIATIONS"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@GUIDE_WORDS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["GUIDE_WORDS"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@PROCESS_DEVIATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PROCESS_DEVIATION"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@AREA_APPLICATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["AREA_APPLICATION"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@PARAMETER", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PARAMETER"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@DEF_SELECTED", SqlDbType.Int) { Value = dt.Rows[i]["DEF_SELECTED"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });

                                iMaxSeq += 1;
                                #endregion
                            }
                            else if (action_type == "update")
                            {
                                #region update
                                sqlstr = "update EPHA_M_GUIDE_WORDS set " +
                                         "NO = @NO, NO_DEVIATIONS = @NO_DEVIATIONS, NO_GUIDE_WORDS = @NO_GUIDE_WORDS, DEVIATIONS = @DEVIATIONS, GUIDE_WORDS = @GUIDE_WORDS, " +
                                         "PROCESS_DEVIATION = @PROCESS_DEVIATION, AREA_APPLICATION = @AREA_APPLICATION, PARAMETER = @PARAMETER, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                         "DEF_SELECTED = @DEF_SELECTED, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                         "where SEQ = @SEQ and ID = @ID";

                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = dt.Rows[i]["NO"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@NO_DEVIATIONS", SqlDbType.Int) { Value = dt.Rows[i]["NO_DEVIATIONS"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@NO_GUIDE_WORDS", SqlDbType.Int) { Value = dt.Rows[i]["NO_GUIDE_WORDS"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@DEVIATIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DEVIATIONS"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@GUIDE_WORDS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["GUIDE_WORDS"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@PROCESS_DEVIATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PROCESS_DEVIATION"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@AREA_APPLICATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["AREA_APPLICATION"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@PARAMETER", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PARAMETER"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@DEF_SELECTED", SqlDbType.Int) { Value = dt.Rows[i]["DEF_SELECTED"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                                #endregion
                            }
                            else if (action_type == "delete")
                            {
                                #region delete
                                sqlstr = "delete from EPHA_M_GUIDE_WORDS where SEQ = @SEQ and ID = @ID";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                                #endregion
                            }
                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                if (!string.IsNullOrEmpty(action_type))
                                {
                                    if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        ret = "User is not authorized to perform this action.";
                                        break;
                                    }
                                    else
                                    {
                                        #region ExecuteNonQuerySQL Data
                                        var command = transaction.conn.CreateCommand();
                                        //command.CommandType = CommandType.StoredProcedure;
                                        command.CommandText = sqlstr;
                                        if (parameters != null && parameters?.Count > 0)
                                        {
                                            foreach (var _param in parameters)
                                            {
                                                if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                                {
                                                    command.Parameters.Add(_param);
                                                }
                                            }
                                            //command.Parameters.AddRange(parameters?.ToArray());
                                        }
                                        ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                        #endregion  ExecuteNonQuerySQL Data
                                    }
                                    if (ret != "true") break;
                                }
                            }
                        }

                    }
                    #endregion update data guidewords

                    if (ret == "true")
                    {
                        #region update data drawing
                        if (dsData.Tables["drawing"] != null)
                        {
                            dt = new DataTable();
                            //dt = dsData.Tables["drawing"].Copy(); dt.AcceptChanges();
                            dt = ConvertDStoDT(dsData, "drawing");

                            for (int i = 0; i < dt?.Rows.Count; i++)
                            {
                                string action_type = (dt.Rows[i]["action_type"]?.ToString() ?? "").ToLower();
                                string sqlstr = "";
                                List<SqlParameter> parameters = new List<SqlParameter>();

                                if (action_type == "insert")
                                {
                                    #region insert
                                    sqlstr = "insert into EPHA_M_DRAWING (" +
                                             "SEQ, ID, MODULE, NAME, DESCRIPTIONS, DOCUMENT_FILE_SIZE, DOCUMENT_FILE_PATH, DOCUMENT_FILE_NAME, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                             "values (@SEQ, @ID, @MODULE, @NAME, @DESCRIPTIONS, @DOCUMENT_FILE_SIZE, @DOCUMENT_FILE_PATH, @DOCUMENT_FILE_NAME, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@MODULE", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["MODULE"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = dt.Rows[i]["DOCUMENT_FILE_SIZE"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DOCUMENT_FILE_PATH"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DOCUMENT_FILE_NAME"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                                    #endregion
                                }
                                else if (action_type == "update")
                                {
                                    #region update
                                    sqlstr = "update EPHA_M_DRAWING set " +
                                             "MODULE = @MODULE, NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, DOCUMENT_FILE_SIZE = @DOCUMENT_FILE_SIZE, " +
                                             "DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH, DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                             "where SEQ = @SEQ and ID = @ID";

                                    parameters.Add(new SqlParameter("@MODULE", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["MODULE"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = dt.Rows[i]["DOCUMENT_FILE_SIZE"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DOCUMENT_FILE_PATH"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DOCUMENT_FILE_NAME"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                                    #endregion
                                }
                                else if (action_type == "delete")
                                {
                                    #region delete
                                    sqlstr = "delete from EPHA_M_DRAWING where SEQ = @SEQ and ID = @ID";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                                    #endregion
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    if (!string.IsNullOrEmpty(action_type))
                                    {
                                        if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                        {
                                            ret = "User is not authorized to perform this action.";
                                            break;
                                        }
                                        else
                                        {
                                            #region ExecuteNonQuerySQL Data
                                            var command = transaction.conn.CreateCommand();
                                            //command.CommandType = CommandType.StoredProcedure;
                                            command.CommandText = sqlstr;
                                            if (parameters != null && parameters?.Count > 0)
                                            {
                                                foreach (var _param in parameters)
                                                {
                                                    if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                                    {
                                                        command.Parameters.Add(_param);
                                                    }
                                                }
                                                //command.Parameters.AddRange(parameters?.ToArray());
                                            }
                                            ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                            #endregion  ExecuteNonQuerySQL Data
                                        }
                                        if (ret != "true") break;
                                    }
                                }
                            }

                        }
                        #endregion update data drawing
                    }

                    if (ret == "true")
                    {
                        if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                            transaction.Commit();
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                    }
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    ret = "error: " + ex.Message;
                }
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }

        #endregion HAZOP Module : Guide Words 

        #region JSEA Module : Mandatory Note
        public string get_master_mandatorynote(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            int iMaxSeq = get_max("epha_m_mandatory_note");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change
                        from epha_m_mandatory_note a 
                        order by seq ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_mandatorynote(SetDataMasterModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name ?? "");
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = (param.json_data ?? "");

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }


            SetDataMasterModel param_def = new SetDataMasterModel
            {
                user_name = user_name,
                role_type = role_type,
                json_data = json_data
            };

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);

            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_mandatory_note");
                ret = set_mandatory_note(user_name, role_type, ref dsData, ref iMaxSeq);
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, iMaxSeq.ToString()));
        }

        public string set_mandatory_note(string user_name, string role_type, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "";

            #region update data
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            try
            {
                using (ClassConnectionDb transaction = new ClassConnectionDb())
                {
                    transaction.OpenConnection();
                    transaction.BeginTransaction();

                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"]?.ToString() ?? "").ToLower();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "insert into EPHA_M_MANDATORY_NOTE " +
                                     "(SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_DEF, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "values (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_DEF, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_DEF", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_DEF"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });

                            seq_now += 1;
                            #endregion
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "update EPHA_M_MANDATORY_NOTE set " +
                                     "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_DEF = @ACTIVE_DEF, ACTIVE_TYPE = @ACTIVE_TYPE, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                     "where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_DEF", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_DEF"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                            #endregion
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "delete from EPHA_M_MANDATORY_NOTE where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                            #endregion
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            if (!string.IsNullOrEmpty(action_type))
                            {
                                if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                {
                                    ret = "User is not authorized to perform this action.";
                                    break;
                                }
                                else
                                {
                                    #region ExecuteNonQuerySQL Data
                                    var command = transaction.conn.CreateCommand();
                                    //command.CommandType = CommandType.StoredProcedure;
                                    command.CommandText = sqlstr;
                                    if (parameters != null && parameters?.Count > 0)
                                    {
                                        foreach (var _param in parameters)
                                        {
                                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                            {
                                                command.Parameters.Add(_param);
                                            }
                                        }
                                        //command.Parameters.AddRange(parameters?.ToArray());
                                    }
                                    ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                    #endregion  ExecuteNonQuerySQL Data
                                }
                                if (ret != "true") break;
                            }
                        }
                    }

                    if (ret == "true")
                    {
                        if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                            transaction.Commit();
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }


        #endregion JSEA Module : Mandatory Note

        #region JSEA Module : Task Type
        public string get_master_tasktype(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            int iMaxSeq = get_max("epha_m_request_type");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change
                        from epha_m_request_type a 
                        order by seq ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_tasktype(SetDataMasterModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = param.user_name ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            SetDataMasterModel param_def = new SetDataMasterModel
            {
                user_name = user_name,
                role_type = role_type,
                json_data = json_data
            };

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);

            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_request_type");
                ret = set_request_type(user_name, role_type, ref dsData, ref iMaxSeq);
            }
            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, iMaxSeq.ToString()));
        }

        public string set_request_type(string user_name, string role_type, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "";

            #region update data
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            try
            {
                using (ClassConnectionDb transaction = new ClassConnectionDb())
                {
                    transaction.OpenConnection();
                    transaction.BeginTransaction();

                    try
                    {
                        for (int i = 0; i < dt?.Rows.Count; i++)
                        {
                            string action_type = (dt.Rows[i]["action_type"]?.ToString() ?? "").ToLower();
                            string sqlstr = "";
                            List<SqlParameter> parameters = new List<SqlParameter>();

                            if (action_type == "insert")
                            {
                                seq_now += 1;

                                #region insert
                                sqlstr = "insert into EPHA_M_REQUEST_TYPE " +
                                         "(SEQ, ID, NAME, DESCRIPTIONS, PHA_SUB_SOFTWARE, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                         "values (@SEQ, @ID, @NAME, @DESCRIPTIONS, @PHA_SUB_SOFTWARE, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@PHA_SUB_SOFTWARE", SqlDbType.NVarChar, 50) { Value = "JSEA" });
                                parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });

                                #endregion
                            }
                            else if (action_type == "update")
                            {
                                #region update
                                sqlstr = "update EPHA_M_REQUEST_TYPE set " +
                                         "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                         "where SEQ = @SEQ and ID = @ID";

                                parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                                #endregion
                            }
                            else if (action_type == "delete")
                            {
                                #region delete
                                sqlstr = "delete from EPHA_M_REQUEST_TYPE where SEQ = @SEQ and ID = @ID";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                                #endregion
                            }

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                if (!string.IsNullOrEmpty(action_type))
                                {
                                    if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        ret = "User is not authorized to perform this action.";
                                        break;
                                    }
                                    else
                                    {
                                        #region ExecuteNonQuerySQL Data
                                        var command = transaction.conn.CreateCommand();
                                        //command.CommandType = CommandType.StoredProcedure;
                                        command.CommandText = sqlstr;
                                        if (parameters != null && parameters?.Count > 0)
                                        {
                                            foreach (var _param in parameters)
                                            {
                                                if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                                {
                                                    command.Parameters.Add(_param);
                                                }
                                            }
                                            //command.Parameters.AddRange(parameters?.ToArray());
                                        }
                                        ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                        #endregion  ExecuteNonQuerySQL Data
                                    }
                                    if (ret != "true") break;
                                }
                            }
                        }

                        if (ret == "true")
                        {
                            if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                                transaction.Commit();
                            }
                            else
                            {
                                transaction.Rollback();
                            }
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    catch (Exception ex_transaction)
                    {
                        transaction.Rollback();
                        ret = "error: " + ex_transaction.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }


        #endregion JSEA Module : Task Type

        #region JSEA Module : Tag ID/Equipment
        public string get_master_tagid(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            int iMaxSeq = get_max("epha_m_tagid");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change
                        from epha_m_tagid a 
                        order by seq ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_tagid(SetDataMasterModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = param.user_name ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = user_name;
            param_def.role_type = role_type;
            param_def.json_data = json_data;

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_tagid");
                ret = set_tagid(user_name, role_type, ref dsData, ref iMaxSeq);
            }
            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, iMaxSeq.ToString()));
        }

        public string set_tagid(string user_name, string role_type, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }


            string ret = "";

            #region update data
            try
            {
                dt = new DataTable();
                dt = ConvertDStoDT(dsData, "data");
                if (dt != null)
                {
                    if (dt?.Rows.Count > 0)
                    {
                        using (ClassConnectionDb transaction = new ClassConnectionDb())
                        {
                            transaction.OpenConnection();
                            transaction.BeginTransaction();

                            try
                            {
                                for (int i = 0; i < dt?.Rows.Count; i++)
                                {
                                    string action_type = (dt.Rows[i]["action_type"]?.ToString() ?? "").ToLower();
                                    string sqlstr = "";
                                    List<SqlParameter> parameters = new List<SqlParameter>();

                                    if (action_type == "insert")
                                    {
                                        #region insert
                                        sqlstr = "insert into EPHA_M_TAGID " +
                                                 "(SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                                 "values (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"]?.ToString() ?? "" });
                                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"]?.ToString() ?? "" });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });

                                        seq_now += 1;
                                        #endregion
                                    }
                                    else if (action_type == "update")
                                    {
                                        #region update
                                        sqlstr = "update EPHA_M_TAGID set " +
                                                 "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                                 "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                                 "where SEQ = @SEQ and ID = @ID";

                                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"]?.ToString() ?? "" });
                                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                                        #endregion
                                    }
                                    else if (action_type == "delete")
                                    {
                                        #region delete
                                        sqlstr = "delete from EPHA_M_TAGID where SEQ = @SEQ and ID = @ID";

                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                                        #endregion
                                    }

                                    if (!string.IsNullOrEmpty(sqlstr))
                                    {
                                        if (!string.IsNullOrEmpty(action_type))
                                        {
                                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                            {
                                                ret = "User is not authorized to perform this action.";
                                                break;
                                            }
                                            else
                                            {
                                                #region ExecuteNonQuerySQL Data
                                                var command = transaction.conn.CreateCommand();
                                                //command.CommandType = CommandType.StoredProcedure;
                                                command.CommandText = sqlstr;
                                                if (parameters != null && parameters?.Count > 0)
                                                {
                                                    foreach (var _param in parameters)
                                                    {
                                                        if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                                        {
                                                            command.Parameters.Add(_param);
                                                        }
                                                    }
                                                    //command.Parameters.AddRange(parameters?.ToArray());
                                                }
                                                ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                                #endregion  ExecuteNonQuerySQL Data
                                            }
                                            if (ret != "true") break;
                                        }
                                    }
                                }


                                if (ret == "true")
                                {
                                    if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                                        transaction.Commit();
                                    }
                                    else
                                    {
                                        transaction.Rollback();
                                    }
                                }
                                else
                                {
                                    transaction.Rollback();
                                }
                            }
                            catch (Exception ex_transaction)
                            {
                                transaction.Rollback();
                                ret = "error: " + ex_transaction.Message;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }

        #endregion JSEA Module : Tag ID/Equipment

        #region HRA : Sections Group => Sub Area Group
        public string get_master_sections_group(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            int iMaxSeq = get_max("epha_m_sections_group");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_sections_group a 
                        order by seq ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }

                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_sections_group(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";

            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }


            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_sections_group");
                ret = set_sections_group(user_name, role_type, ref dsData, ref iMaxSeq);
            }
            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }

        public string set_sections_group(string user_name, string role_type, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }


            string ret = "";

            #region update data

            try
            {
                dt = new DataTable();
                dt = ConvertDStoDT(dsData, "data");

                if (dt != null)
                {
                    if (dt?.Rows.Count > 0)
                    {
                        using (ClassConnectionDb transaction = new ClassConnectionDb())
                        {
                            transaction.OpenConnection();
                            transaction.BeginTransaction();

                            try
                            {
                                for (int i = 0; i < dt?.Rows.Count; i++)
                                {
                                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                                    string sqlstr = "";
                                    List<SqlParameter> parameters = new List<SqlParameter>();

                                    if (action_type == "insert")
                                    {
                                        #region insert
                                        sqlstr = "insert into EPHA_M_SECTIONS_GROUP " +
                                                 "(SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                                 "values (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });

                                        seq_now += 1;
                                        #endregion
                                    }
                                    else if (action_type == "update")
                                    {
                                        #region update
                                        sqlstr = "update EPHA_M_SECTIONS_GROUP set " +
                                                 "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                                 "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                                 "where SEQ = @SEQ and ID = @ID";

                                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });
                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                                        #endregion
                                    }
                                    else if (action_type == "delete")
                                    {
                                        #region delete
                                        sqlstr = "delete from EPHA_M_SECTIONS_GROUP where SEQ = @SEQ and ID = @ID";

                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                                        #endregion
                                    }

                                    if (!string.IsNullOrEmpty(sqlstr))
                                    {
                                        if (!string.IsNullOrEmpty(action_type))
                                        {
                                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                            {
                                                ret = "User is not authorized to perform this action.";
                                                break;
                                            }
                                            else
                                            {
                                                #region ExecuteNonQuerySQL Data
                                                var command = transaction.conn.CreateCommand();
                                                //command.CommandType = CommandType.StoredProcedure;
                                                command.CommandText = sqlstr;
                                                if (parameters != null && parameters?.Count > 0)
                                                {
                                                    foreach (var _param in parameters)
                                                    {
                                                        if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                                        {
                                                            command.Parameters.Add(_param);
                                                        }
                                                    }
                                                    //command.Parameters.AddRange(parameters?.ToArray());
                                                }
                                                ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                                #endregion  ExecuteNonQuerySQL Data
                                            }
                                            if (ret != "true") break;
                                        }
                                    }
                                }

                                if (ret == "true")
                                {
                                    if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                                        transaction.Commit();
                                    }
                                    else
                                    {
                                        transaction.Rollback();
                                    }
                                }
                                else
                                {
                                    transaction.Rollback();
                                }
                            }
                            catch (Exception ex_transaction)
                            {
                                transaction.Rollback();
                                ret = "error: " + ex_transaction.Message;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }

        #endregion HRA : Sections Group => Sub Area Group

        #region HRA : Group of Sub Area
        public string get_master_sub_area_group(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            #region Departments
            sqlstr = @"select distinct emp.departments as id,emp.departments as name
                         ,emp.functions, emp.departments
                         from vw_epha_person_details emp
                         where emp.departments is not null
                         order by emp.functions, emp.departments";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            dt.TableName = "departments";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Departments

            #region Sections
            sqlstr = @"select distinct emp.sections as id, emp.sections as name,
                         emp.functions, emp.departments, emp.sections
                         from vw_epha_person_details emp
                         where emp.departments is not null and emp.sections is not null
                         order by emp.functions, emp.departments, emp.sections";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "sections";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Sections

            #region Group of Sub Area 
            sqlstr = @" select  a.id, a.name, lower(a.name) as  field_check  
                        from epha_m_sections_group  a
                        where a.active_type = 1
                        order by a.name";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "sections_group";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Group of Sub Area


            int iMaxSeq = get_max("epha_m_sub_area");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_sub_area a 
                        order by id_sections, id_sections_group, seq";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;
                dt.Rows[0]["id_sections_group"] = null;
                dt.Rows[0]["id_sections"] = null;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_sub_area_group(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";

            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }


            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_sub_area");
                ret = set_sub_area_group(user_name, role_type, ref dsData, ref iMaxSeq);
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }

        public string set_sub_area_group(string user_name, string role_type, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "";

            #region update data
            try
            {
                dt = new DataTable();
                dt = ConvertDStoDT(dsData, "data");

                if (dt != null)
                {
                    if (dt?.Rows.Count > 0)
                    {
                        using (ClassConnectionDb transaction = new ClassConnectionDb())
                        {
                            transaction.OpenConnection();
                            transaction.BeginTransaction();

                            try
                            {

                                for (int i = 0; i < dt?.Rows.Count; i++)
                                {
                                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                                    string sqlstr = "";
                                    List<SqlParameter> parameters = new List<SqlParameter>();

                                    if (action_type == "insert")
                                    {
                                        #region insert
                                        sqlstr = "insert into EPHA_M_SUB_AREA " +
                                                 "(ID_SECTIONS, ID_SECTIONS_GROUP, SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                                 "values (@ID_SECTIONS, @ID_SECTIONS_GROUP, @SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                                        parameters.Add(new SqlParameter("@ID_SECTIONS", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["ID_SECTIONS"] });
                                        parameters.Add(new SqlParameter("@ID_SECTIONS_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_SECTIONS_GROUP"] });
                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });

                                        seq_now += 1;
                                        #endregion
                                    }
                                    else if (action_type == "update")
                                    {
                                        #region update
                                        sqlstr = "update EPHA_M_SUB_AREA set " +
                                                 "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                                 "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                                 "where SEQ = @SEQ and ID = @ID";

                                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });
                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                                        #endregion
                                    }
                                    else if (action_type == "delete")
                                    {
                                        #region delete
                                        sqlstr = "delete from EPHA_M_SUB_AREA where SEQ = @SEQ and ID = @ID";

                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                                        #endregion
                                    }

                                    if (!string.IsNullOrEmpty(sqlstr))
                                    {
                                        if (!string.IsNullOrEmpty(action_type))
                                        {
                                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                            {
                                                ret = "User is not authorized to perform this action.";
                                                break;
                                            }
                                            else
                                            {
                                                #region ExecuteNonQuerySQL Data
                                                var command = transaction.conn.CreateCommand();
                                                //command.CommandType = CommandType.StoredProcedure;
                                                command.CommandText = sqlstr;
                                                if (parameters != null && parameters?.Count > 0)
                                                {
                                                    foreach (var _param in parameters)
                                                    {
                                                        if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                                        {
                                                            command.Parameters.Add(_param);
                                                        }
                                                    }
                                                    //command.Parameters.AddRange(parameters?.ToArray());
                                                }
                                                ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                                #endregion  ExecuteNonQuerySQL Data
                                            }
                                            if (ret != "true") break;
                                        }
                                    }
                                }

                                if (ret == "true")
                                {
                                    if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                                        transaction.Commit();
                                    }
                                    else
                                    {
                                        transaction.Rollback();
                                    }
                                }
                                else
                                {
                                    transaction.Rollback();
                                }
                            }
                            catch (Exception ex_transaction)
                            {
                                transaction.Rollback();
                                ret = "error: " + ex_transaction.Message;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }

        #endregion HRA : Group of Sub Area

        #region HRA : Equipmet of Sub Area 
        public void get_master_data_of_area(ref DataSet dsData)
        {
            #region Plant
            sqlstr = @"  select distinct a.id_plant as id, a.plant as name, a.plant_check from vw_epha_data_of_area a order by a.plant ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "plant";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Plant

            #region Area Process Unit
            sqlstr = @"  select distinct a.id_area as id, a.area as name, a.plant_check, a.area_check from vw_epha_data_of_area a order by a.area ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "area";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Area Process Unit

            #region Complex

            sqlstr = @" select distinct a.id_toc as id, a.toc as name, a.plant_check, a.area_check, a.toc_check from vw_epha_data_of_area a order by a.toc ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "toc";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Complex

            #region master apu

            sqlstr = @"  select distinct a.id_unit as id, a.unit +'-'+ a.toc as name, a.plant_check, a.area_check, a.toc_check from vw_epha_data_of_area a order by a.unit +'-'+ a.toc  ";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            dt.TableName = "apu";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion master apu

        }
        public string get_master_sub_area_equipmet(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            get_master_data_of_area(ref dsData);

            #region Group of Sub Area 
            sqlstr = @" select  a.id, a.name, lower(a.name) as  field_check  
                        from epha_m_sections_group  a
                        where a.active_type = 1
                        order by a.name";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            var parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt != null)
            {
                dt.TableName = "sections_group";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            #endregion Group of Sub Area


            int iMaxSeq = get_max("epha_m_hazard_type");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_sub_area a 
                        order by id_sections, id_sections_group, seq";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;
                dt.Rows[0]["id_sections_group"] = null;
                dt.Rows[0]["id_sections"] = null;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_sub_area_equipmet(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";

            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_sub_area");
                ret = set_sub_area_equipmet(user_name, role_type, ref dsData, ref iMaxSeq);
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }

        public string set_sub_area_equipmet(string user_name, string role_type, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "";

            #region update data

            try
            {
                dt = new DataTable();
                dt = ConvertDStoDT(dsData, "data");

                if (dt != null)
                {
                    if (dt?.Rows.Count > 0)
                    {
                        using (ClassConnectionDb transaction = new ClassConnectionDb())
                        {
                            transaction.OpenConnection();
                            transaction.BeginTransaction();

                            try
                            {
                                for (int i = 0; i < dt?.Rows.Count; i++)
                                {
                                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                                    string sqlstr = "";
                                    List<SqlParameter> parameters = new List<SqlParameter>();

                                    if (action_type == "insert")
                                    {
                                        #region insert
                                        sqlstr = "insert into EPHA_M_SUB_AREA " +
                                                 "(ID_SECTIONS, ID_SECTIONS_GROUP, SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                                 "values (@ID_SECTIONS, @ID_SECTIONS_GROUP, @SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                                        parameters.Add(new SqlParameter("@ID_SECTIONS", SqlDbType.Int) { Value = dt.Rows[i]["ID_SECTIONS"] });
                                        parameters.Add(new SqlParameter("@ID_SECTIONS_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_SECTIONS_GROUP"] });
                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });

                                        seq_now += 1;
                                        #endregion
                                    }
                                    else if (action_type == "update")
                                    {
                                        #region update
                                        sqlstr = "update EPHA_M_SUB_AREA set " +
                                                 "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                                 "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                                 "where SEQ = @SEQ and ID = @ID";

                                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                                        parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });
                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                                        #endregion
                                    }
                                    else if (action_type == "delete")
                                    {
                                        #region delete
                                        sqlstr = "delete from EPHA_M_SUB_AREA where SEQ = @SEQ and ID = @ID";

                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                                        #endregion
                                    }

                                    if (!string.IsNullOrEmpty(sqlstr))
                                    {
                                        if (!string.IsNullOrEmpty(action_type))
                                        {
                                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                            {
                                                ret = "User is not authorized to perform this action.";
                                                break;
                                            }
                                            else
                                            {
                                                #region ExecuteNonQuerySQL Data
                                                var command = transaction.conn.CreateCommand();
                                                //command.CommandType = CommandType.StoredProcedure;
                                                command.CommandText = sqlstr;
                                                if (parameters != null && parameters?.Count > 0)
                                                {
                                                    foreach (var _param in parameters)
                                                    {
                                                        if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                                        {
                                                            command.Parameters.Add(_param);
                                                        }
                                                    }
                                                    //command.Parameters.AddRange(parameters?.ToArray());
                                                }
                                                ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                                #endregion  ExecuteNonQuerySQL Data
                                            }
                                            if (ret != "true") break;

                                        }
                                    }
                                }

                                if (ret == "true")
                                {
                                    if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                                        transaction.Commit();
                                    }
                                    else
                                    {
                                        transaction.Rollback();
                                    }
                                }
                                else
                                {
                                    transaction.Rollback();
                                }
                            }
                            catch (Exception ex_transaction)
                            {
                                transaction.Rollback();
                                ret = "error: " + ex_transaction.Message;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }
            #endregion update data

            return ret;
        }

        #endregion HRA : Equipmet of Sub Area

        #region HRA : Hazard Type & Hazard Riskfactors
        public string get_master_hazard_type(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }
            DataTable dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "");

            int iMaxSeq = get_max("epha_m_hazard_type");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_hazard_type a 
                        order by seq ";

            //
            DataTable dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            #region Execute to Datable
            try
            {
                var parameters = new List<SqlParameter>();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    dt.TableName = "sections";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_hazard_type(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";

            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_hazard_type");

                try
                {
                    ret = set_hazard_type(user_name, role_type, ref dsData, ref iMaxSeq);
                }
                catch (Exception ex)
                {
                    ret = "error: " + ex.Message;
                }
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, iMaxSeq.ToString()));
        }

        public string set_hazard_type(string user_name, string role_type, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "";
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            using (ClassConnectionDb transaction = new ClassConnectionDb())
            {
                transaction.OpenConnection();
                transaction.BeginTransaction();

                try
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string action_type = row["action_type"]?.ToString() ?? "";
                        List<SqlParameter> parameters = new List<SqlParameter>();
                        string sqlstr = "";

                        if (action_type == "insert")
                        {
                            sqlstr = "insert into EPHA_M_HAZARD_TYPE " +
                                     "(SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "values (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, getdate(), @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });

                            seq_now++;
                        }
                        else if (action_type == "update")
                        {
                            sqlstr = "update EPHA_M_HAZARD_TYPE set " +
                                     "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                     "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                     "where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });
                        }
                        else if (action_type == "delete")
                        {
                            sqlstr = "delete from EPHA_M_HAZARD_TYPE where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            if (!string.IsNullOrEmpty(action_type))
                            {
                                if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                {
                                    ret = "User is not authorized to perform this action.";
                                    break;
                                }
                                else
                                {

                                    #region ExecuteNonQuerySQL Data
                                    var command = transaction.conn.CreateCommand();
                                    //command.CommandType = CommandType.StoredProcedure;
                                    command.CommandText = sqlstr;
                                    if (parameters != null && parameters?.Count > 0)
                                    {
                                        foreach (var _param in parameters)
                                        {
                                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                            {
                                                command.Parameters.Add(_param);
                                            }
                                        }
                                        //command.Parameters.AddRange(parameters?.ToArray());
                                    }
                                    ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                    #endregion  ExecuteNonQuerySQL Data

                                }
                                if (ret != "true") break;

                            }
                        }
                    }
                    if (ret == "true")
                    {
                        if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                            transaction.Commit();
                            //// Mark the transaction scope as complete
                            //scope.Complete();
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                    }
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    ret = "error: " + ex.Message;
                }
            }

            return ret;
        }

        public string get_master_hazard_riskfactors(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name ?? "");

            int iMaxSeq = get_max("epha_m_hazard_riskfactors");

            sqlstr = @" select p.id, p.name 
                        from epha_m_hazard_type p 
                        where p.active_type = 1
                        order by p.id";


            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable
            try
            {
                var parameters = new List<SqlParameter>();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt != null)
            {
                dt.TableName = "hazard_type";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_hazard_riskfactors a 
                        order by seq ";

            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            #region Execute to Datable
            try
            {
                var parameters = new List<SqlParameter>();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["id_hazard_type"] = null;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_hazard_riskfactors(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";

            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }


            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, ""));

            iMaxSeq = get_max("epha_m_hazard_riskfactors");

            try
            {
                ret = set_hazard_riskfactors(user_name, role_type, ref dsData, ref iMaxSeq);
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, iMaxSeq.ToString()));
        }

        public string set_hazard_riskfactors(string user_name, string role_type, ref DataSet dsData, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }


            string ret = "";
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            using (ClassConnectionDb transaction = new ClassConnectionDb())
            {
                transaction.OpenConnection();
                transaction.BeginTransaction();

                try
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string action_type = row["action_type"]?.ToString() ?? "";
                        List<SqlParameter> parameters = new List<SqlParameter>();
                        string sqlstr = "";
                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (action_type == "insert")
                            {
                                sqlstr = "insert into EPHA_M_HAZARD_RISKFACTORS " +
                                         "(SEQ, ID, ID_HAZARD_TYPE, HEALTH_HAZARDS, HAZARDS_RATING, STANDARD_TYPE_TEXT, STANDARD_VALUE, STANDARD_UNIT, STANDARD_DESC, ACTIVE_TYPE, " +
                                         "CREATE_DATE, CREATE_BY, UPDATE_BY) " +
                                         "values (@SEQ, @ID, @ID_HAZARD_TYPE, @HEALTH_HAZARDS, @HAZARDS_RATING, @STANDARD_TYPE_TEXT, @STANDARD_VALUE, @STANDARD_UNIT, @STANDARD_DESC, @ACTIVE_TYPE, " +
                                         "getdate(), @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@ID_HAZARD_TYPE", SqlDbType.Int) { Value = row["ID_HAZARD_TYPE"] });
                                parameters.Add(new SqlParameter("@HEALTH_HAZARDS", SqlDbType.NVarChar, 4000) { Value = row["HEALTH_HAZARDS"] });
                                parameters.Add(new SqlParameter("@HAZARDS_RATING", SqlDbType.NVarChar, 4000) { Value = row["HAZARDS_RATING"] });
                                parameters.Add(new SqlParameter("@STANDARD_TYPE_TEXT", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_TYPE_TEXT"] });
                                parameters.Add(new SqlParameter("@STANDARD_VALUE", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_VALUE"] });
                                parameters.Add(new SqlParameter("@STANDARD_UNIT", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_UNIT"] });
                                parameters.Add(new SqlParameter("@STANDARD_DESC", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_DESC"] });
                                parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });

                                seq_now++;
                            }
                            else if (action_type == "update")
                            {
                                sqlstr = "update EPHA_M_HAZARD_RISKFACTORS set " +
                                         "ID_HAZARD_TYPE = @ID_HAZARD_TYPE, HEALTH_HAZARDS = @HEALTH_HAZARDS, HAZARDS_RATING = @HAZARDS_RATING, " +
                                         "STANDARD_TYPE_TEXT = @STANDARD_TYPE_TEXT, STANDARD_VALUE = @STANDARD_VALUE, STANDARD_UNIT = @STANDARD_UNIT, " +
                                         "STANDARD_DESC = @STANDARD_DESC, ACTIVE_TYPE = @ACTIVE_TYPE, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                         "where SEQ = @SEQ and ID = @ID";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                parameters.Add(new SqlParameter("@ID_HAZARD_TYPE", SqlDbType.Int) { Value = row["ID_HAZARD_TYPE"] });
                                parameters.Add(new SqlParameter("@HEALTH_HAZARDS", SqlDbType.NVarChar, 4000) { Value = row["HEALTH_HAZARDS"] });
                                parameters.Add(new SqlParameter("@HAZARDS_RATING", SqlDbType.NVarChar, 4000) { Value = row["HAZARDS_RATING"] });
                                parameters.Add(new SqlParameter("@STANDARD_TYPE_TEXT", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_TYPE_TEXT"] });
                                parameters.Add(new SqlParameter("@STANDARD_VALUE", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_VALUE"] });
                                parameters.Add(new SqlParameter("@STANDARD_UNIT", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_UNIT"] });
                                parameters.Add(new SqlParameter("@STANDARD_DESC", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_DESC"] });
                                parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });
                            }
                            else if (action_type == "delete")
                            {
                                sqlstr = "delete from EPHA_M_HAZARD_RISKFACTORS where SEQ = @SEQ and ID = @ID";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                            }
                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                if (!string.IsNullOrEmpty(action_type))
                                {
                                    if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                    {
                                        ret = "User is not authorized to perform this action.";
                                        break;
                                    }
                                    else
                                    {

                                        #region ExecuteNonQuerySQL Data
                                        var command = transaction.conn.CreateCommand();
                                        //command.CommandType = CommandType.StoredProcedure;
                                        command.CommandText = sqlstr;
                                        if (parameters != null && parameters?.Count > 0)
                                        {
                                            foreach (var _param in parameters)
                                            {
                                                if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                                {
                                                    command.Parameters.Add(_param);
                                                }
                                            }
                                            //command.Parameters.AddRange(parameters?.ToArray());
                                        }
                                        ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                                        #endregion  ExecuteNonQuerySQL Data

                                    }
                                    if (ret != "true") break;

                                }
                            }
                        }
                    }
                    if (ret == "true")
                    {
                        if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                            transaction.Commit();
                            //// Mark the transaction scope as complete
                            //scope.Complete();
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                    }

                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    ret = "error: " + ex.Message;
                }
            }

            return ret;
        }


        #endregion HRA : Hazard Type & Hazard Riskfactors


        #region Department, Sections
        private void getMasterDepartmentSections(ref DataSet ds, string sections_name = "", string departments_name = "")
        {
            // ตรวจสอบค่า เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (ds == null)
            {
                return;
            }

            List<SqlParameter> parameters = new List<SqlParameter>();
            dt = new DataTable();

            #region Sections 
            parameters = new List<SqlParameter>();
            if (sections_name != "") { parameters.Add(new SqlParameter("@sections_name", SqlDbType.VarChar) { Value = sections_name }); }
            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "sections", true);

            #region Execute to Datable
            try
            {
                _conn = new ClassConnectionDb();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_GetMasterSections";
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    dt.TableName = "sections";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt != null)
            {
                if (dt?.Rows.Count > 0) { departments_name = dt.Rows[0]["departments"]?.ToString() ?? ""; }
                dt.TableName = "sections";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            #endregion Sections

            #region Departments  
            parameters = new List<SqlParameter>();
            if (departments_name != "") { parameters.Add(new SqlParameter("@departments_name", SqlDbType.VarChar) { Value = departments_name }); }
            dt = new DataTable();
            #region Execute to Datable
            try
            {
                _conn = new ClassConnectionDb();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_GetMasterDepartments";
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    dt.TableName = "departments";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable
            if (dt != null)
            {
                dt.TableName = "departments";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            #endregion Departments

        }
        #endregion Department, Sections

        #region HRA : Group List
        public string get_master_group_list(LoadMasterPageBySectionModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            List<SqlParameter> parameters = new List<SqlParameter>();
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = param.user_name ?? "";  // ป้องกันการ Dereference null ของ user_name

            // ดึงค่าขั้นสูงสุดของ epha_m_group_list
            int iMaxSeq = get_max("epha_m_group_list");

            // เรียก Stored Procedure เพื่อดึงข้อมูล  
            #region Execute to Datable
            try
            {
                _conn = new ClassConnectionDb();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_Get_EPHA_M_Group_List";
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            // ตรวจสอบ DataTable ก่อนใช้งาน เพื่อป้องกัน Dereference null return value
            if (dt == null || dt.Rows.Count == 0)
            {
                if (dt == null)  // ถ้า dt เป็น null ให้สร้าง DataTable ใหม่
                {
                    dt = new DataTable();
                }

                // กรณีที่เป็นใบงานใหม่
                DataRow newRow = dt.NewRow();
                newRow["seq"] = iMaxSeq;
                newRow["id"] = iMaxSeq;
                newRow["create_by"] = user_name;
                newRow["action_type"] = "insert";
                newRow["action_change"] = 0;
                dt.Rows.Add(newRow);

                dt.AcceptChanges();

                // เพิ่มลำดับ iMaxSeq
                iMaxSeq += 1;
            }

            // ตั้งชื่อ Table ว่า data และเพิ่มลงใน dsData
            dt.AcceptChanges();
            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();

            // ตั้งค่า max_id สำหรับ seq
            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            // เพิ่ม DataTable ที่เก็บค่า max ลงใน dsData
            dsData.Tables.Add(dtma.Copy());
            dsData.AcceptChanges();

            // ตั้งชื่อ DataSet
            dsData.DataSetName = "dsData";
            dsData.AcceptChanges();

            // แปลง DataSet เป็น JSON
            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        public string set_master_group_list(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }
            iMaxSeq = get_max("epha_m_group_list");


            using (ClassConnectionDb transaction = new ClassConnectionDb())
            {
                transaction.OpenConnection();
                transaction.BeginTransaction();

                try
                {
                    // Data
                    ret = set_group_list(user_name, role_type, ref dsData, transaction, ref iMaxSeq);

                    if (ret == "true")
                    {
                        if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                            transaction.Commit();
                            //// Mark the transaction scope as complete
                            //scope.Complete();
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                    }

                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    ret = "error: " + ex.Message;
                }
            }

        Next_Line_Convert:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));

        }
        public string set_group_list(string user_name, string role_type, ref DataSet dsData, ClassConnectionDb transaction, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "";
            List<SqlParameter> parameters = new List<SqlParameter>();
            DataTable dt = ConvertDStoDT(dsData, "data");

            foreach (DataRow row in dt.Rows)
            {
                string action_type = row["action_type"]?.ToString() ?? "";

                parameters = new List<SqlParameter>();

                if (action_type == "insert")
                {
                    sqlstr = "INSERT INTO EPHA_M_GROUP_LIST " +
                             "(SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                             "VALUES (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.VarChar, 4000) { Value = row["NAME"] });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = row["DESCRIPTIONS"] });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = row["CREATE_BY"] });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = row["UPDATE_BY"] });

                    seq_now += 1;
                }
                else if (action_type == "update")
                {
                    sqlstr = "UPDATE EPHA_M_GROUP_LIST SET " +
                             "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE " +
                             "WHERE SEQ = @SEQ AND ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.VarChar, 4000) { Value = row["NAME"] });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = row["DESCRIPTIONS"] });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                }
                else if (action_type == "delete")
                {
                    sqlstr = "DELETE FROM EPHA_M_GROUP_LIST WHERE SEQ = @SEQ AND ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                }

                if (!string.IsNullOrEmpty(sqlstr))
                {
                    if (!string.IsNullOrEmpty(action_type))
                    {
                        if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            ret = "User is not authorized to perform this action.";
                            break;
                        }
                        else
                        {
                            #region ExecuteNonQuerySQL Data
                            var command = transaction.conn.CreateCommand();
                            //command.CommandType = CommandType.StoredProcedure;
                            command.CommandText = sqlstr;
                            if (parameters != null && parameters?.Count > 0)
                            {
                                foreach (var _param in parameters)
                                {
                                    if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                    {
                                        command.Parameters.Add(_param);
                                    }
                                }
                                //command.Parameters.AddRange(parameters?.ToArray());
                            }
                            ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                            #endregion  ExecuteNonQuerySQL Data

                        }
                        if (ret != "true") break;

                    }
                }
            }

            return ret;
        }

        #endregion HRA : Group List

        #region HRA : Worker Group
        public string get_master_worker_group(LoadMasterPageBySectionModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }
            List<SqlParameter> parameters = new List<SqlParameter>();
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name?.ToString() ?? "");
            string id_sections = (param.id_sections?.ToString() ?? "");
            string id_group_list = "";
            string key_group_list = "";

            getMasterDepartmentSections(ref dsData);

            if (dsData.Tables["sections"]?.Rows.Count > 0)
            {
                id_sections = dsData.Tables["sections"]?.Rows[0]["id"]?.ToString() ?? "";
            }

            parameters = new List<SqlParameter>();
            dt = new DataTable();

            #region Execute to Datable
            try
            {
                _conn = new ClassConnectionDb();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_Get_EPHA_M_Group_List";
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    dt.TableName = "group_list";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt?.Rows.Count > 0)
            {
                id_group_list = dt.Rows[0]["id"]?.ToString() ?? "";
                key_group_list = dt.Rows[0]["name"]?.ToString() ?? "";
            }
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            int iMaxSeq = get_max("epha_m_worker_group");

            sqlstr = @"  select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page 
                         from epha_m_worker_group a  
                         where a.id_sections is not null  
                         order by a.seq ";

            parameters = new List<SqlParameter>();
            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", false);

            #region Execute to Datable
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                else
                {
                    if (dsData.Tables["group_list"]?.Rows.Count > 0)
                    {
                        id_group_list = dsData.Tables["group_list"]?.Rows[0]["id"]?.ToString() ?? "";
                        key_group_list = dsData.Tables["group_list"]?.Rows[0]["name"]?.ToString() ?? "";
                        for (int i = 0; i < dsData.Tables["group_list"]?.Rows.Count; i++)
                        {
                            //กรณีที่เป็นใบงานใหม่
                            dt.Rows.Add(dt.NewRow());
                            dt.Rows[i]["seq"] = iMaxSeq;
                            dt.Rows[i]["id"] = iMaxSeq;
                            dt.Rows[i]["id_sections"] = id_sections;
                            dt.Rows[i]["id_group_list"] = id_group_list;
                            dt.Rows[i]["key_group_list"] = key_group_list;

                            dt.Rows[i]["create_by"] = user_name;
                            dt.Rows[i]["action_type"] = "insert";
                            dt.Rows[i]["action_change"] = 0;
                            dt.AcceptChanges();
                            iMaxSeq += 1;
                        }
                    }
                    else
                    {
                        //กรณีที่เป็นใบงานใหม่
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["seq"] = iMaxSeq;
                        dt.Rows[0]["id"] = iMaxSeq;
                        dt.Rows[0]["id_sections"] = id_sections;
                        dt.Rows[0]["id_group_list"] = id_group_list;
                        dt.Rows[0]["key_group_list"] = key_group_list;

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                        iMaxSeq += 1;
                    }
                }
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();


            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_worker_group(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }
            iMaxSeq = get_max("epha_m_worker_group");


            using (ClassConnectionDb transaction = new ClassConnectionDb())
            {
                transaction.OpenConnection();
                transaction.BeginTransaction();

                try
                {
                    // Data
                    ret = set_worker_group(user_name, role_type, ref dsData, transaction, ref iMaxSeq);
                    if (ret == "true")
                    {
                        if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                            transaction.Commit();
                            //// Mark the transaction scope as complete
                            //scope.Complete();
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                    }

                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    ret = "error: " + ex.Message;
                }
            }

        Next_Line_Convert:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));

        }
        public string set_worker_group(string user_name, string role_type, ref DataSet dsData, ClassConnectionDb transaction, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "";
            List<SqlParameter> parameters = new List<SqlParameter>();
            DataTable dt = ConvertDStoDT(dsData, "data");

            foreach (DataRow row in dt.Rows)
            {
                string action_type = row["action_type"]?.ToString() ?? "";

                parameters = new List<SqlParameter>();

                if (action_type == "insert")
                {
                    sqlstr = "INSERT INTO EPHA_M_WORKER_GROUP " +
                             "(ID_SECTIONS, ID_GROUP_LIST, KEY_GROUP_LIST, SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                             "VALUES (@ID_SECTIONS, @ID_GROUP_LIST, @KEY_GROUP_LIST, @SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                    parameters.Add(new SqlParameter("@ID_SECTIONS", SqlDbType.VarChar, 4000) { Value = row["ID_SECTIONS"] });
                    parameters.Add(new SqlParameter("@ID_GROUP_LIST", SqlDbType.VarChar, 4000) { Value = row["ID_GROUP_LIST"] });
                    parameters.Add(new SqlParameter("@KEY_GROUP_LIST", SqlDbType.VarChar, 4000) { Value = row["KEY_GROUP_LIST"] });
                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.VarChar, 4000) { Value = row["NAME"] });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = row["DESCRIPTIONS"] });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = row["CREATE_BY"] });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = row["UPDATE_BY"] });

                    seq_now += 1;
                }
                else if (action_type == "update")
                {
                    sqlstr = "UPDATE EPHA_M_WORKER_GROUP SET " +
                             "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE , ID_GROUP_LIST = @ID_GROUP_LIST , KEY_GROUP_LIST = @KEY_GROUP_LIST " +
                             "WHERE SEQ = @SEQ AND ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                    parameters.Add(new SqlParameter("@ID_GROUP_LIST", SqlDbType.VarChar, 4000) { Value = row["ID_GROUP_LIST"] });
                    parameters.Add(new SqlParameter("@KEY_GROUP_LIST", SqlDbType.VarChar, 4000) { Value = row["KEY_GROUP_LIST"] });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.VarChar, 4000) { Value = row["NAME"] });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = row["DESCRIPTIONS"] });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                }
                else if (action_type == "delete")
                {
                    sqlstr = "DELETE FROM EPHA_M_WORKER_GROUP WHERE SEQ = @SEQ AND ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                }

                if (!string.IsNullOrEmpty(sqlstr))
                {
                    if (!string.IsNullOrEmpty(action_type))
                    {
                        if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            ret = "User is not authorized to perform this action.";
                            break;
                        }
                        else
                        {
                            #region ExecuteNonQuerySQL Data
                            var command = transaction.conn.CreateCommand();
                            //command.CommandType = CommandType.StoredProcedure;
                            command.CommandText = sqlstr;
                            if (parameters != null && parameters?.Count > 0)
                            {
                                foreach (var _param in parameters)
                                {
                                    if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                    {
                                        command.Parameters.Add(_param);
                                    }
                                }
                                //command.Parameters.AddRange(parameters?.ToArray());
                            }
                            ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                            #endregion  ExecuteNonQuerySQL Data
                        }
                        if (ret != "true") break;

                    }
                }
            }

            return ret;
        }

        #endregion HRA : Worker Group

        #region HRA : Worker List
        public string get_master_worker_list(LoadMasterPageModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }
            List<SqlParameter> parameters = new List<SqlParameter>();
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name?.ToString() ?? "");
            string id_worker_group = "";
            string id_sections = "";

            getMasterDepartmentSections(ref dsData);

            if (dsData.Tables["sections"]?.Rows.Count > 0)
            {
                id_sections = dsData.Tables["sections"]?.Rows[0]["id"]?.ToString() ?? "";
            }

            parameters = new List<SqlParameter>();
            dt = new DataTable();
            #region Execute to Datable
            try
            {
                _conn = new ClassConnectionDb();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_get_epha_m_worker_group";
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    dt.TableName = "worker_group";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dt?.Rows.Count > 0)
            {
                id_worker_group = dt.Rows[0]["id"]?.ToString() ?? "";
            }
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            int iMaxSeq = get_max("epha_m_worker_list");
            sqlstr = @"  select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                         , b.id_sections
                         from epha_m_worker_list a 
                         inner join epha_m_worker_group b on a.id_worker_group = b.id
                         where a.id_worker_group is not null and b.id_sections is not null
                         order by a.seq ";

            parameters = new List<SqlParameter>();
            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", false);

            #region Execute to Datable
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    //command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    if (parameters != null && parameters?.Count > 0)
                    {
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                    }
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
            }

            if (dt != null)
            {
                if (dsData.Tables["worker_group"]?.Rows.Count > 0)
                {
                    Boolean bOldData = true;
                    if (dt?.Rows.Count == 0) { bOldData = false; } else { bOldData = true; }

                    for (int i = 0; i < dsData.Tables["worker_group"]?.Rows.Count; i++)
                    {
                        Boolean bNewRow = true;
                        string id_worker_group_select = dsData.Tables["worker_group"]?.Rows[i]["id"]?.ToString() ?? "";

                        if (bOldData)
                        {
                            if (dt?.Rows.Count > 0)
                            {
                                var filterParameters = new Dictionary<string, object>();
                                filterParameters.Add("id_worker_group", id_worker_group_select);
                                var (drEmpActive, iEmpActive) = FilterDataTable(dt, filterParameters);
                                if (drEmpActive != null)
                                {
                                    if (drEmpActive?.Length > 0)
                                    {
                                        bNewRow = false;
                                    }
                                }
                            }
                        }

                        if (bNewRow)
                        {
                            //กรณีที่เป็นใบงานใหม่
                            dt.Rows.Add(dt.NewRow());
                            dt.Rows[i]["seq"] = iMaxSeq;
                            dt.Rows[i]["id"] = iMaxSeq;
                            dt.Rows[i]["id_worker_group"] = ConvertToIntOrDBNull(id_worker_group_select);
                            dt.Rows[i]["id_sections"] = ConvertToIntOrDBNull(id_sections);

                            dt.Rows[i]["create_by"] = user_name;
                            dt.Rows[i]["action_type"] = "insert";
                            dt.Rows[i]["action_change"] = 0;
                            dt.AcceptChanges();
                            iMaxSeq += 1;
                        }
                    }
                }
                else
                {
                    //กรณีที่เป็นใบงานใหม่
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["seq"] = iMaxSeq;
                    dt.Rows[0]["id"] = iMaxSeq;
                    dt.Rows[0]["id_worker_group"] = ConvertToIntOrDBNull(id_worker_group);
                    dt.Rows[0]["id_sections"] = ConvertToIntOrDBNull(id_sections);

                    dt.Rows[0]["create_by"] = user_name;
                    dt.Rows[0]["action_type"] = "insert";
                    dt.Rows[0]["action_change"] = 0;
                    dt.AcceptChanges();
                    iMaxSeq += 1;
                }
            }

            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_worker_list(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string json_data = param.json_data?.ToString() ?? "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToData(user_name, role_type, ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }
            iMaxSeq = get_max("epha_m_worker_group");

            using (ClassConnectionDb transaction = new ClassConnectionDb())
            {
                transaction.OpenConnection();
                transaction.BeginTransaction();

                try
                {
                    // Data
                    ret = set_worker_list(user_name, role_type, ref dsData, transaction, ref iMaxSeq);
                    if (ret == "true")
                    {
                        if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
                            transaction.Commit();
                            //// Mark the transaction scope as complete
                            //scope.Complete();
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                    }

                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    ret = "error: " + ex.Message;
                }
            }

        Next_Line_Convert:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));

        }
        public string set_worker_list(string user_name, string role_type, ref DataSet dsData, ClassConnectionDb transaction, ref int seq_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string ret = "";
            List<SqlParameter> parameters = new List<SqlParameter>();
            DataTable dt = ConvertDStoDT(dsData, "data");

            foreach (DataRow row in dt.Rows)
            {
                string action_type = row["action_type"]?.ToString() ?? "";

                parameters = new List<SqlParameter>();

                if (action_type == "insert")
                {
                    sqlstr = "INSERT INTO EPHA_M_WORKER_LIST " +
                             "(ID_WORKER_GROUP,SEQ,ID,NO,USER_NAME,USER_DISPLAYNAME,USER_TITLE,USER_TYPE,ACTIVE_TYPE,CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY) " +
                             "VALUES (@ID_WORKER_GROUP, @SEQ, @ID, @NO, @USER_NAME, @USER_DISPLAYNAME, @USER_TITLE, @USER_TYPE,  @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                    parameters.Add(new SqlParameter("@ID_WORKER_GROUP", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_WORKER_GROUP"]) });
                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_now) });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_now) });
                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                    parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["USER_NAME"]) });
                    parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["USER_DISPLAYNAME"]) });
                    parameters.Add(new SqlParameter("@USER_TITLE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["USER_TITLE"]) });
                    parameters.Add(new SqlParameter("@USER_TYPE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["USER_TYPE"]) });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ACTIVE_TYPE"]) });
                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = row["CREATE_BY"] });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = row["UPDATE_BY"] });

                    seq_now += 1;
                }
                else if (action_type == "update")
                {
                    sqlstr = "UPDATE EPHA_M_WORKER_LIST SET " +
                             "NO = @NO, USER_NAME = @USER_NAME, USER_DISPLAYNAME = @USER_DISPLAYNAME, USER_TITLE = @USER_TITLE, USER_TYPE = @USER_TYPE, ACTIVE_TYPE = @ACTIVE_TYPE " +
                             "WHERE SEQ = @SEQ AND ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                    parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["USER_NAME"]) });
                    parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["USER_DISPLAYNAME"]) });
                    parameters.Add(new SqlParameter("@USER_TITLE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["USER_TITLE"]) });
                    parameters.Add(new SqlParameter("@USER_TYPE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["USER_TYPE"]) });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ACTIVE_TYPE"]) });
                }
                else if (action_type == "delete")
                {
                    sqlstr = "DELETE FROM EPHA_M_WORKER_LIST WHERE SEQ = @SEQ AND ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                }

                if (!string.IsNullOrEmpty(sqlstr))
                {
                    if (!string.IsNullOrEmpty(action_type))
                    {
                        if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                        {
                            ret = "User is not authorized to perform this action.";
                            break;
                        }
                        else
                        {
                            #region ExecuteNonQuerySQL Data
                            var command = transaction.conn.CreateCommand();
                            //command.CommandType = CommandType.StoredProcedure;
                            command.CommandText = sqlstr;
                            if (parameters != null && parameters?.Count > 0)
                            {
                                foreach (var _param in parameters)
                                {
                                    if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                                    {
                                        command.Parameters.Add(_param);
                                    }
                                }
                                //command.Parameters.AddRange(parameters?.ToArray());
                            }
                            ret = transaction.ExecuteNonQuerySQL(command, user_name, role_type);
                            #endregion  ExecuteNonQuerySQL Data
                        }
                        if (ret != "true") break;

                    }
                }
            }

            return ret;
        }

        #endregion HRA : Worker List

    }
}
