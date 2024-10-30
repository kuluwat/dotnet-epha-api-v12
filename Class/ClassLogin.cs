
using dotnet_epha_api.Class;
using dotnet6_epha_api.Class;
using Model;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;


namespace Class
{
    public class ClassLogin
    {
        string sqlstr = "";
        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb cls_conn = new ClassConnectionDb();
        ClassConnectionDb _conn = new ClassConnectionDb();

        static public bool IsAuthorizedRole(string userName, string roleType)
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

        static public bool IsAuthorized(string userName)
        {
            // ตรวจสอบว่ามี userName หรือไม่
            return !string.IsNullOrEmpty(userName);
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
        static public string GetUserRoleFromDb(string user_name)
        {
            string role_type = "";

            if (string.IsNullOrEmpty(user_name)) { return role_type; }

            //กรณีที่เป็น Employee ที่กำหนดสิทธิ์ในระบบ
            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();
             
            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_name))
            {
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }

            ClassConnectionDb _conn = new ClassConnectionDb();
            dt = new DataTable(); 
            #region Execute to Datable
            //parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_GetQueryUserRole";
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

            if (dt != null)
            {
                if (dt?.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        role_type = dt.Rows[i]["role_type"]?.ToString() ?? "";
                        if (role_type == "admin") { break; }
                    }

                }
            }

            return role_type;
        }
        public DataTable dataUserRole(string user_name)
        {
            user_name = user_name ?? "";

            //กรณีที่เป็น Employee ที่กำหนดสิทธิ์ในระบบ
            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();
             
            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_name))
            {
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }

            dt = new DataTable(); 
            #region Execute to Datable
            //parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText ="usp_GetQueryUserRole"  ;
                    //command.Parameters.Add(":costcenter", costcenter); 
                    command.Parameters.Clear();
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

            return dt;
        }
        public DataTable dataEmployeeRole(string user_name)
        {
            user_name = user_name ?? "";

            //กรณีที่เป็น Employee ทั่วไปเข้าใช้งานระบบ  
            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();
             
            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_name))
            {
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }

            dt = new DataTable(); 
            #region Execute to Datable
            //parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText =  "usp_GetQueryEmployeeRole" ;
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

            return dt;
        }
        private DataTable _dataUser_Role(LoginUserModel param)
        {
            string userName = (param?.user_name ?? string.Empty);  // ตรวจสอบว่า param เป็น null หรือไม่
            try
            {
                // ตรวจสอบว่ามี '@' ในชื่อผู้ใช้หรือไม่ และถ้ามีก็แยกออกจากกัน
                if (userName.Contains("@"))
                {
                    string[] userNameParts = userName.Split('@');
                    if (userNameParts.Length > 1)
                    {
                        userName = userNameParts[0];  // ใช้ส่วนแรกของอีเมลเป็นชื่อผู้ใช้
                    }
                }
            }
            catch (Exception ex)
            {
                // คุณสามารถเพิ่มการจัดการข้อผิดพลาดที่นี่ เช่น การเขียน log หรือแสดงข้อความแจ้งเตือน
                // LogError(ex);
            }

            DataTable dt = new DataTable();
            cls = new ClassFunctions();

            // ดึงข้อมูลบทบาทของผู้ใช้จากฟังก์ชัน dataUserRole
            dt = dataUserRole(userName);

            // ถ้าไม่มีข้อมูลใน dt หรือจำนวนแถวเป็น 0 ให้ดึงข้อมูลจาก dataEmployeeRole
            if (dt == null || dt.Rows.Count == 0)
            {
                cls_conn = new ClassConnectionDb();
                dt = dataEmployeeRole(userName);
            }

            // ถ้าผู้ใช้คือ "admin" ให้แก้ไขข้อมูลแถวแรกใน DataTable
            if (userName.Equals("admin", StringComparison.OrdinalIgnoreCase) && dt.Rows.Count > 0)
            {
                DataRow adminRow = dt.Rows[0];
                adminRow["role_type"] = "admin";
                adminRow["user_name"] = "admin";
                adminRow["user_id"] = "00000000";
                adminRow["user_email"] = "admin-epha@thaioilgroup.com";
                adminRow["user_display"] = userName + " (Admin)";
                adminRow["user_img"] = "images/user-avatar.png";
                dt.AcceptChanges();
            }

            return dt;
        }

        public string login(LoginUserModel param, ref DataTable dtRef)
        {
            DataTable dt = new DataTable();
            string user_name = (param.user_name ?? "");

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            try
            {
                dt = new DataTable();
                dt = _dataUser_Role(param);
                if (dt != null)
                {
                    dtRef = dt; dtRef.AcceptChanges();
                }
                return cls_json.SetJSONresult(dt);
            }
            catch (Exception ex)
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", ex.Message.ToString()));
            }
        }

        public string authorization_page(PageRoleListModel param)
        {
            string page_controller = param.page_controller ?? "";
            string user_name = param.user_name?.ToString() ?? "";
            try
            {
                if (user_name.IndexOf("@") > -1)
                {
                    string[] x = user_name.Split('@');
                    if (x.Length > 1)
                    {
                        user_name = x[0];
                    }
                }
            }
            catch { }

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            try
            {
                DataTable dt = new DataTable();
                dt = _dtAuthorization_Page(user_name, page_controller);

                return cls_json.SetJSONresult(dt);
            }
            catch (Exception ex)
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", ex.Message.ToString()));
            }
        }

        public string check_authorization_page_fix(PageRoleListModel param)
        {
            if (param == null)
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            DataTable dt = new DataTable();
            string page_controller = (param.page_controller ?? "");
            string user_name = (param.user_name ?? "");
            try
            {
                if (!string.IsNullOrEmpty(user_name))
                {
                    if (user_name.IndexOf("@") > -1)
                    {
                        string[] x = user_name.Split('@');
                        if (x.Length > 1)
                        {
                            user_name = x[0];
                        }
                    }
                }
            }
            catch { }

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            try
            {
                string role_type = _dtAuthorization_RoleType(user_name);

                dt = _dtAuthorization_Page(user_name, page_controller);
                dt.Columns.Add("followup_page", typeof(int));
                dt.AcceptChanges();

                ClassHazop cls = new ClassHazop();
                DataTable dtFollow = new DataTable();
                dtFollow = cls.DataHomeTask((role_type == "admin" ? "" : user_name), role_type, "", false, true, "13");
                if (dtFollow?.Rows.Count == 0)
                {
                    dtFollow = cls.DataHomeTask((role_type == "admin" ? "" : user_name), role_type, "", false, true, "14");
                }
                else
                {
                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        dt.Rows[i]["followup_page"] = 0;
                        string pha_type = (dt.Rows[i]["page_controller"]?.ToString() ?? "").ToUpper() + "'";
                        var filterParameters = new Dictionary<string, object>();
                        filterParameters.Add("pha_type", pha_type);
                        var (drWorksheet, iMerge) = FilterDataTable(dtFollow, filterParameters);
                        if (drWorksheet != null)
                        {
                            if (drWorksheet?.Length > 0)
                            {
                                dt.Rows[i]["followup_page"] = 1;
                                dt.AcceptChanges();
                            }
                        }
                    }

                    dt.AcceptChanges();
                }

                return cls_json.SetJSONresult(dt);
            }
            catch (Exception ex) { return cls_json.SetJSONresult(ClassFile.refMsg("Error", ex.Message.ToString())); }

        }
        public string _dtAuthorization_RoleType(string user_name)
        {
            user_name = (user_name ?? "").ToLower();

            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();
             
            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_name))
            {
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }

            dt = new DataTable(); 
            #region Execute to Datable
            //parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_GetQueryUserRole"  ;
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

            string role_type = "employee"; // ค่าเริ่มต้นเป็น employee
            if (dt?.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    role_type = (dt.Rows[i]["role_type"]?.ToString() ?? "").ToLower();
                    if (role_type == "admin")
                    {
                        role_type = "admin";
                        break; // ออกจากลูปทันทีเมื่อเจอ admin
                    }
                }
            }

            // คืนค่า role_type ตามที่ตรวจสอบได้ ไม่คืน employee เสมอ
            return role_type;
        }
        public DataTable _dtAuthorization_Page(string user_name, string page_controller)
        {
            user_name = (user_name ?? "").ToLower();
            user_name = (user_name == "admin" ? "" : user_name);
            page_controller = (page_controller ?? "").ToLower();


            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();
             
            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_name))
            {
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }
            if (!string.IsNullOrEmpty(page_controller))
            {
                parameters.Add(new SqlParameter("@page_controller", SqlDbType.VarChar, 1000) { Value = page_controller });
            }

            dt = new DataTable(); 
            #region Execute to Datable
            //parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb();
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_GetQueryPageRole" ;
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

            if (dt?.Rows.Count == 0)
            {
                if (!string.IsNullOrEmpty(user_name))
                { 
                    parameters = new List<SqlParameter>();
                    if (!string.IsNullOrEmpty(user_name))
                    {
                        parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
                    }
                    dt = new DataTable(); 
                    #region Execute to Datable
                    //parameters = new List<SqlParameter>();
                    try
                    {
                        _conn = new ClassConnectionDb();
                        _conn = new ClassConnectionDb(); _conn.OpenConnection();
                        try
                        {
                            var command = _conn.conn.CreateCommand();
                            command.CommandType = CommandType.StoredProcedure;
                            command.CommandText = "usp_GetQueryPageRoleEmployee";
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
                    if (dt?.Rows.Count > 0)
                    {
                        return dt;
                    }
                }
                if (dt?.Rows.Count == 0)
                {
                    //กรณีที่ไม่มี menu?? ถึงขั้นตอนนี้น่าจะต้องมี --> create new row 
                    dt.NewRow(); dt.AcceptChanges();
                }
            }

            return dt;
        }

        public DataTable _dtAuthorization_Page_By_Doc(string user_name, string role_type, string page_controller)
        {
            // ตรวจสอบค่า user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (string.IsNullOrEmpty(user_name))
            {
                return new DataTable();
            }

            user_name = (user_name ?? "").ToLower();
            user_name = (user_name == "admin" ? "" : user_name);
            page_controller = (page_controller ?? "").ToLower();


            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            sqlstr = @" select distinct a.user_name, a.pha_sub_software as page_controller from ( 
                             select ht.pha_sub_software, dt.*
                             from epha_t_header ht 
                             inner join (
                             select distinct a.responder_user_name as user_name, a.id_pha 
                             from epha_t_node_worksheet a inner join vw_epha_max_seq_by_pha_no b on a.id_pha = b.id_pha
                             where a.responder_user_name is not null and isnull(a.action_project_team,0) = 0 
                             union
                             select distinct a.user_name, a.id_pha 
                             from epha_t_member_team a inner join vw_epha_max_seq_by_pha_no b on a.id_pha = b.id_pha
                             where a.user_name is not null
                             union
                             select distinct a.user_name, a.id_pha 
                             from epha_t_approver a inner join vw_epha_max_seq_by_pha_no b on a.id_pha = b.id_pha
                             where a.user_name is not null
                             union
                             select distinct a.user_name, a.id_pha 
                             from (select ta3.user_name, ta3.id_pha from epha_t_approver_ta3 ta3 inner join epha_t_approver ta2 on ta3.id_approver = ta2.id)a inner join vw_epha_max_seq_by_pha_no b on a.id_pha = b.id_pha
                             where a.user_name is not null
                             union
                             select distinct a.request_user_name as user_name, a.id as id_pha 
                             from epha_t_header a inner join vw_epha_max_seq_by_pha_no b on a.id = b.id_pha
                             where a.request_user_name is not null
                             )dt on ht.id = dt.id_pha 
                        )a where a.pha_sub_software is not null";

            if (!string.IsNullOrEmpty(user_name))
            {
                sqlstr += " and lower(a.user_name)  = lower(@user_name)  ";
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
            }
            if (!string.IsNullOrEmpty(page_controller))
            {
                sqlstr += " and lower(a.page_controller)  = lower(@page_controller)  ";
                parameters.Add(new SqlParameter("@page_controller", SqlDbType.VarChar, 1000) { Value = page_controller ?? "" });
            }
            sqlstr += " order by a.pha_sub_software ";

            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
            #region Execute to Datable
            //parameters = new List<SqlParameter>();
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

            if (dt?.Rows.Count == 0)
            {
                // กรณีที่ไม่มี data?? ถึงขั้นตอนนี้น่าจะต้องมี --> create new row
                DataRow newRow = dt.NewRow();
                dt.Rows.Add(newRow);
                dt.AcceptChanges();
            }

            return dt;
        }
        private int get_max(string table_name)
        {
            if (string.IsNullOrEmpty(table_name))
            {
                return 0;
            }

            // ตรวจสอบชื่อ table ให้มีเฉพาะอักขระที่ปลอดภัย
            if (!Regex.IsMatch(table_name, @"^[a-zA-Z0-9_]+$"))
            {
                throw new ArgumentException("Invalid table name format.");
            }

            // เรียกใช้ stored procedure โดยใช้ชื่อ table
            List<SqlParameter> parameters = new List<SqlParameter>();

            parameters.Add(new SqlParameter("@TableName", SqlDbType.NVarChar) { Value = table_name });
            parameters.Add(new SqlParameter("@NextId", SqlDbType.Int) { Direction = ParameterDirection.Output });
             
            DataTable dt = new DataTable();
            #region Execute to Datable
            //parameters = new List<SqlParameter>();
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
                    //dt.TableName = "data";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            return Convert.ToInt32(parameters[1].Value);  // ค่า @NextId ที่ถูกส่งออกมา
        }

        public string register_account(RegisterAccountModel param)
        {
            string user_displayname = (param.user_displayname ?? "");
            string user_email = (param.user_email ?? "");
            string user_password = (param.user_password ?? "");
            string user_password_confirm = (param.user_password_confirm ?? "");

            //string user_name = user_email ?? ""; 
            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            try
            {
                ClassEmail clsemail = new ClassEmail();
                string hashedPassword = clsemail.EncryptString(user_password);
                string hashedPasswordConfirm = clsemail.EncryptString(user_password_confirm);

                string ret = "";
                string msg = "";

                DataTable dt = new DataTable();
                cls = new ClassFunctions();

                string sqlstr = @"
                      SELECT a.user_name, a.user_id, a.user_email, a.user_displayname
                      ,LOWER(COALESCE(c.name, 'employee')) AS role_type
                      FROM VW_EPHA_PERSON_DETAILS a
                      INNER JOIN EPHA_M_ROLE_SETTING b ON LOWER(a.user_name) = LOWER(b.user_name) AND b.active_type = 1
                      INNER JOIN EPHA_M_ROLE_TYPE c ON LOWER(c.id) = LOWER(b.id_role_group) AND c.active_type = 1
                      WHERE a.active_type = 1";

                var parameters = new List<SqlParameter>();
                if (!string.IsNullOrEmpty(user_email))
                {
                    sqlstr += " AND LOWER(a.user_email) = LOWER(@user_email)";
                    parameters.Add(new SqlParameter("@user_email", SqlDbType.VarChar, 100) { Value = user_email });
                }
                sqlstr += " ORDER BY a.user_name, c.name";

                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", false);
                #region Execute to Datable
                //parameters = new List<SqlParameter>();
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


                if (dt?.Rows.Count > 0)
                {
                    ret = "true";
                    msg = "User already has data in the system.";
                }
                else
                {
                    int seq = get_max("EPHA_REGISTER_ACCOUNT");

                    using (ClassConnectionDb transaction = new ClassConnectionDb())
                    {
                        transaction.OpenConnection();
                        transaction.BeginTransaction();
                        try
                        {
                            sqlstr = @"INSERT INTO EPHA_REGISTER_ACCOUNT
                               (SEQ, ID, REGISTER_TYPE, USER_DISPLAYNAME, USER_EMAIL, USER_PASSWORD, USER_PASSWORD_CONFIRM, ACCEPT_STATUS, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY)
                               VALUES (@SEQ, @ID, @REGISTER_TYPE, @USER_DISPLAYNAME, @USER_EMAIL, @USER_PASSWORD, @USER_PASSWORD_CONFIRM, @ACCEPT_STATUS, GETDATE(), NULL, @CREATE_BY, NULL)";

                            parameters = new List<SqlParameter>
                            {
                                new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq },
                                new SqlParameter("@ID", SqlDbType.Int) { Value = seq },
                                new SqlParameter("@REGISTER_TYPE", SqlDbType.Int) { Value = 1 },
                                new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = user_displayname },
                                new SqlParameter("@USER_EMAIL", SqlDbType.VarChar, 100) { Value = user_email },
                                new SqlParameter("@USER_PASSWORD", SqlDbType.VarChar, 50) { Value = hashedPassword },
                                new SqlParameter("@USER_PASSWORD_CONFIRM", SqlDbType.VarChar, 50) { Value = hashedPasswordConfirm },
                                new SqlParameter("@ACCEPT_STATUS", SqlDbType.VarChar, 50) { Value = DBNull.Value },
                                new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = "system" }
                            };

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                {
                                    ret = "User is not authorized to perform this action.";
                                    //break;
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
                                //if (ret != "true") break;
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

                    if (ret.ToLower() == "true")
                    {
                        ret = "true";
                        msg = "User registration is complete. Please wait for the login credentials from the system administrator.";
                    }
                    else
                    {
                        ret = "error";
                        msg = ret;
                    }
                }

                if (ret.ToLower() == "true")
                {
                    // Email แจ้ง admin ให้ accept การ register
                    ClassEmail clsmail = new ClassEmail();
                    //clsmail.MailToAdminRegisterAccount(user_displayname, user_email, user_password, user_password_confirm);
                    clsmail.MailToAdminRegisterAccount(user_displayname, user_email);
                }

                dt = new DataTable();
                dt.Columns.Add("status");
                dt.Columns.Add("msg");
                dt.AcceptChanges();

                dt.Rows.Add(dt.NewRow());
                dt.AcceptChanges();
                dt.Rows[0]["status"] = ret;
                dt.Rows[0]["msg"] = msg;

                return cls_json.SetJSONresult(dt);
            }
            catch (Exception ex)
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", ex.Message.ToString()));
            }
        }

        public string update_register_account(RegisterAccountModel param)
        {
            string user_active = (param.user_active ?? "");
            string user_email = (param.user_email ?? "");
            string accept_status = (param.accept_status ?? "");

            //string user_name = user_email ?? ""; 
            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            try
            {
                string ret = "";
                string msg = "";

                DataTable dt = new DataTable();
                cls = new ClassFunctions();

                string sqlstr = @"SELECT a.user_name, a.user_id, a.user_email, a.user_displayname
                      ,LOWER(COALESCE(c.name, 'employee')) AS role_type
                      FROM VW_EPHA_PERSON_DETAILS a
                      INNER JOIN EPHA_M_ROLE_SETTING b ON LOWER(a.user_name) = LOWER(b.user_name) AND b.active_type = 1
                      INNER JOIN EPHA_M_ROLE_TYPE c ON LOWER(c.id) = LOWER(b.id_role_group) AND c.active_type = 1
                      WHERE a.active_type = 1";

                var parameters = new List<SqlParameter>();
                if (!string.IsNullOrEmpty(user_email))
                {
                    sqlstr += " AND LOWER(a.user_email) = LOWER(@user_email)";
                    parameters.Add(new SqlParameter("@user_email", SqlDbType.VarChar, 100) { Value = user_email });
                }
                sqlstr += " ORDER BY a.user_name, c.name";

                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", false);
                #region Execute to Datable
                //parameters = new List<SqlParameter>();
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

                #region insert/update  
                int seq = get_max("EPHA_REGISTER_ACCOUNT");
                 
                using (ClassConnectionDb transaction = new ClassConnectionDb())
                {
                    transaction.OpenConnection();
                    transaction.BeginTransaction();
                    try
                    {
                        sqlstr = @"UPDATE EPHA_REGISTER_ACCOUNT SET
                           ACCEPT_STATUS = @ACCEPT_STATUS,
                           UPDATE_DATE = GETDATE(),
                           UPDATE_BY = @UPDATE_BY
                           WHERE USER_EMAIL = @USER_EMAIL";

                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@ACCEPT_STATUS", SqlDbType.VarChar, 50) { Value = accept_status });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 400) { Value = user_active });
                        parameters.Add(new SqlParameter("@USER_EMAIL", SqlDbType.VarChar, 100) { Value = user_email });


                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                ret = "User is not authorized to perform this action.";
                                //break;
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
                            //if (ret != "true") break;
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
                #endregion insert/update    

                if (ret.ToLower() == "true")
                {
                    ret = "true";
                    msg = "User registration update is complete.";
                }
                else
                {
                    ret = "error";
                    msg = ret;
                }

                if (ret.ToLower() == "true")
                {
                    if (dt?.Rows.Count > 0)
                    {
                        string user_displayname = (dt.Rows[0]["user_displayname"] + "");
                        string user_password = (dt.Rows[0]["user_password"] + "");
                        string user_password_confirm = (dt.Rows[0]["user_password_confirm"] + "");

                        // email แจ้งผู้ใช้งานว่าการลงทะเบียนสำเร็จ
                        ClassEmail clsmail = new ClassEmail();
                        //clsmail.MailToUserRegisterAccount(user_displayname, user_email, user_password, user_password_confirm, accept_status);
                        clsmail.MailToUserRegisterAccount(user_displayname, user_email, accept_status);
                    }
                }

                dt = new DataTable();
                dt.Columns.Add("status");
                dt.Columns.Add("msg");
                dt.AcceptChanges();

                dt.Rows.Add(dt.NewRow());
                dt.AcceptChanges();
                dt.Rows[0]["status"] = ret;
                dt.Rows[0]["msg"] = msg;

                return cls_json.SetJSONresult(dt);
            }
            catch (Exception ex)
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", ex.Message.ToString()));
            }

        }
    }
}
