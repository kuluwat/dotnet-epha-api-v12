

using dotnet_epha_api.Class;
using dotnet6_epha_api.Class;
using Model;
using Newtonsoft.Json;

using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
namespace Class
{
    public class ClassHazop
    {
        string sqlstr = "";
        string jsper = "";
        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb _conn = new ClassConnectionDb();
        ClassConnectionDb cls_conn = new ClassConnectionDb();

        #region function
        string[] sMonth = ("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC").Split(',');
        public string convert_revision_text(string _index)
        {
            string[] characters = ("a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z").Split(',');

            // Validate if the version is a valid integer
            if (int.TryParse(_index, out int versionIndex) && versionIndex >= 0 && versionIndex < characters.Length)
            {
                // Return the character at the specified index
                return characters[versionIndex];
            }
            else
            {
                // Handle invalid version
                return "";
            }
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

        #region function master / all

        public string get_pha_no(string sub_software, string year)
        {
            if (string.IsNullOrWhiteSpace(sub_software) || string.IsNullOrWhiteSpace(year))
            {
                throw new ArgumentException("sub_software and year cannot be null or empty.");
            }

            DataTable _dt = new DataTable();
            cls = new ClassFunctions();

            string sqlstr = @"
            select @subSoftware + '-' + @year + '-' + right('0000000' + trim(str(coalesce(max(replace(upper(pha_no), @subSoftware + '-' + @year + '-', '') + 1), 1))), 7) as pha_no
            from epha_t_header
            where lower(pha_sub_software) = lower(@subSoftware) and year = @year";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@subSoftware", SqlDbType.VarChar) { Value = sub_software.ToUpper() });
            parameters.Add(new SqlParameter("@year", SqlDbType.VarChar) { Value = year.ToUpper() });

            //_dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
            #region Execute to Datable
            //var parameters = new List<SqlParameter>();
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
                    _dt = new DataTable();
                    _dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    _dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            return _dt.Rows[0]["pha_no"]?.ToString() ?? "";
        }

        public void set_max_id(ref DataTable dtmax, string name, string values)
        {
            if (dtmax == null)
            {
                dtmax = new DataTable();
            }

            if (!dtmax.Columns.Contains("name"))
            {
                dtmax.Columns.Add("name");
            }

            if (!dtmax.Columns.Contains("values"))
            {
                dtmax.Columns.Add("values");
            }

            dtmax.AcceptChanges();

            DataRow newRow = dtmax.NewRow();
            newRow["name"] = name;
            newRow["values"] = values;
            dtmax.Rows.Add(newRow);
            dtmax.AcceptChanges();
        }
        public int get_max(string table_name, string id_pha = "")
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
                DataTable dt = new DataTable();
                // เรียกใช้ stored procedure โดยใช้ชื่อ table
                List<SqlParameter> parameters = new List<SqlParameter>();

                parameters.Add(new SqlParameter("@TableName", SqlDbType.NVarChar) { Value = table_name });
                parameters.Add(new SqlParameter("@NextId", SqlDbType.Int) { Direction = ParameterDirection.Output });
                if (!string.IsNullOrEmpty(id_pha))
                {
                    parameters.Add(new SqlParameter("@IdPha", SqlDbType.NVarChar) { Value = id_pha });
                }
                #region Execute to Datable
                //var parameters = new List<SqlParameter>();
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
            catch
            {
                return 0;
            }
        }

        private void get_history_doc(ref DataSet _dsData, string sub_software)
        {
            var parameters = new List<SqlParameter>();
            DataTable dt = new DataTable();

            string sqlstr = @"select * from(
                        select distinct b.reference_moc as name
                        from epha_t_header a 
                        inner join EPHA_T_GENERAL b on a.id = b.id_pha 
                        where a.pha_status not in (81)
                      ) t 
                      where t.name is not null 
                      order by t.name";
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
            dt.TableName = "his_reference_moc";
            _dsData.Tables.Add(dt.Copy());
            _dsData.AcceptChanges();

            sqlstr = @"select * from(
                select distinct b.pha_request_name as name
                from epha_t_header a 
                inner join EPHA_T_GENERAL b on a.id = b.id_pha 
                where a.pha_status not in (81)
               ) t 
               where t.name is not null 
               order by t.name";
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
            dt.TableName = "his_pha_request_name";
            _dsData.Tables.Add(dt.Copy());
            _dsData.AcceptChanges();

            sqlstr = @"select * from (
                select distinct c.document_name as name
                from epha_t_header a 
                inner join EPHA_T_GENERAL b on a.id = b.id_pha 
                inner join EPHA_T_DRAWING c on a.id = c.id_pha 
                where a.pha_status not in (81)
               ) t 
               where t.name is not null 
               order by t.name";
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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

            dt.TableName = "his_document_name";
            _dsData.Tables.Add(dt.Copy());
            _dsData.AcceptChanges();

            sqlstr = @"select * from (
                select distinct c.document_no as name
                from epha_t_header a 
                inner join EPHA_T_GENERAL b on a.id = b.id_pha 
                inner join EPHA_T_DRAWING c on a.id = c.id_pha 
                where a.pha_status not in (81)
               ) t 
               where t.name is not null 
               order by t.name";
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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

            dt.TableName = "his_document_no";
            _dsData.Tables.Add(dt.Copy());
            _dsData.AcceptChanges();

            //switch (sub_software.ToLower())
            //{
            //    case "hazop":
            //        _history_hazop(ref _dsData);
            //        break;
            //    case "jsea":
            //        _history_jsea(ref _dsData);
            //        break;
            //    case "whatif":
            //        _history_whatif(ref _dsData);
            //        break;
            //}
        }

        //private void _history_hazop(ref DataSet _dsData)
        //{
        //    #region Node History
        //    {
        //        string sqlstr = @"select * from (select distinct c.node as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_node";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion

        //    #region Design Intent History
        //    {
        //        string sqlstr = @"select * from (select distinct c.design_intent as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_design_intent";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion

        //    #region Design Conditions History
        //    {
        //        string sqlstr = @"select * from (select distinct c.design_conditions as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_design_conditions";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion

        //    #region Operating Conditions History
        //    {
        //        string sqlstr = @"select * from (select distinct c.operating_conditions as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_operating_conditions";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion

        //    #region Node Boundary History
        //    {
        //        string sqlstr = @"select * from (select distinct c.node_boundary as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_node_boundary";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion

        //    #region Causes History
        //    {
        //        string sqlstr = @"select * from (select distinct c.causes as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE_WORKSHEET c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_causes";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion

        //    #region Consequences History
        //    {
        //        string sqlstr = @"select * from (select distinct c.consequences as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE_WORKSHEET c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_consequences";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion

        //    #region Existing Safeguards History
        //    {
        //        string sqlstr = @"select * from (select distinct c.existing_safeguards as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE_WORKSHEET c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_existing_safeguards";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion

        //    #region Safety Critical Equipment Tag History
        //    {
        //        string sqlstr = @"select * from (select distinct c.safety_critical_equipment_tag as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE_WORKSHEET c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_safety_critical_equipment_tag";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion

        //    #region Recommendations History
        //    {
        //        string sqlstr = @"select * from (select distinct c.recommendations as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE_WORKSHEET c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_recommendations";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion

        //    #region Note History
        //    {
        //        string sqlstr = @"select * from (select distinct c.note as name
        //                  from epha_t_header a 
        //                  inner join EPHA_T_GENERAL b on a.id = b.id_pha 
        //                  inner join EPHA_T_NODE c on a.id = c.id_pha 
        //                  where a.pha_status not in (81) )t where t.name is not null order by t.name";
        //        DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
        //        dt.TableName = "his_note";
        //        _dsData.Tables.Add(dt.Copy());
        //        _dsData.AcceptChanges();
        //    }
        //    #endregion
        //}

        public void get_master_ram(ref DataSet _dsData)
        {
            // ตรวจสอบค่า  เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (_dsData == null)
            {
                return;
            }
            DataTable dt = new DataTable();

            #region RAM Data
            {
                string sqlstr = @" select seq, id, name, 0 as selected_type, category_type, document_file_size, document_file_name, document_file_path, a.rows_level, a.columns_level
                           , document_definition_file_path
                           , 'update' as action_type, 0 as action_change
                           from EPHA_M_RAM a where active_type = 1
                           order by seq ";

                //DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
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
                    dt.TableName = "ram";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

            }
            #endregion

            #region Security Level Data
            {
                string sqlstr = @" select a.category_type, b.id_ram, b.security_level, b.security_text
                           , people as people_text, assets as assets_text, enhancement as enhancement_text, reputation as reputation_text, product_quality as product_quality_text 
                           from EPHA_M_RAM a 
                           inner join EPHA_M_RAM_LEVEL b on a.id = b.id_ram 
                           order by b.id_ram, b.sort_by ";

                //DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
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
                    dt.TableName = "security_level";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

            }
            #endregion

            #region RAM Level Data
            {
                string sqlstr = @" select  b.*, 0 as selected_type ,a.category_type
                           , b.security_text
                           , people as people_text, assets as assets_text, enhancement as enhancement_text, reputation as reputation_text, product_quality as product_quality_text
                           , a.rows_level, a.columns_level
                           , 'update' as action_type, 0 as action_change
                           from  EPHA_M_RAM a
                           inner join EPHA_M_RAM_LEVEL b on a.id = b.id_ram 
                           order by a.id , b.sort_by ";

                //DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
                dt = new DataTable();
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
                    dt.TableName = "ram_level";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();

                    if (dt?.Rows.Count > 0)
                    {
                        DataTable dtNew = new DataTable();
                        dtNew.Columns.Add("id_ram", typeof(int));
                        dtNew.Columns.Add("selected_type", typeof(int));
                        dtNew.Columns.Add("rows_level", typeof(int));
                        dtNew.Columns.Add("columns_level", typeof(int));
                        dtNew.Columns.Add("likelihood_level");
                        dtNew.Columns.Add("likelihood_show");
                        dtNew.Columns.Add("likelihood_text");
                        dtNew.Columns.Add("likelihood_desc");
                        dtNew.Columns.Add("likelihood_criterion");
                        dtNew.AcceptChanges();

                        if (_dsData?.Tables["ram"] != null)
                        {
                            DataTable dtCopy = _dsData.Tables["ram"].Copy();
                            dtCopy.AcceptChanges();
                            if (dtCopy != null)
                            {
                                if (dtCopy?.Rows.Count > 0)
                                {
                                    foreach (DataRow row in dtCopy.Rows)
                                    {
                                        int id_ram = Convert.ToInt32(row["id"]);
                                        int rows_level = Convert.ToInt32(row["rows_level"]);
                                        int columns_level = Convert.ToInt32(row["columns_level"]);

                                        DataRow[] dr = (_dsData.Tables["ram_level"]).Select("id_ram=" + id_ram);
                                        if (dr.Length > 0)
                                        {
                                            foreach (DataRow rl in dr)
                                            {
                                                for (int j = 1; j <= 7; j++)
                                                {
                                                    if (string.IsNullOrEmpty(rl["likelihood" + j + "_level"]?.ToString())) { break; }

                                                    DataRow newRow = dtNew.NewRow();
                                                    newRow["id_ram"] = id_ram;
                                                    newRow["selected_type"] = 0;
                                                    newRow["rows_level"] = rows_level;
                                                    newRow["columns_level"] = columns_level;
                                                    newRow["likelihood_level"] = rl["likelihood" + j + "_level"];
                                                    newRow["likelihood_show"] = rl["likelihood" + j + "_text"];

                                                    if (columns_level == 5)
                                                    {
                                                        newRow["likelihood_text"] = rl["likelihood" + j + "_text"];
                                                        newRow["likelihood_desc"] = rl["likelihood" + j + "_desc"];
                                                        newRow["likelihood_criterion"] = rl["likelihood" + j + "_criterion"];
                                                    }

                                                    dtNew.Rows.Add(newRow);
                                                    dtNew.AcceptChanges();

                                                    if (j == columns_level) { break; }
                                                }
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        dtNew.TableName = "likelihood_level";
                        _dsData.Tables.Add(dtNew.Copy());
                        _dsData.AcceptChanges();
                    }
                }
            }
            #endregion

            #region RAM Color Data
            {
                string sqlstr = @" select seq,name,descriptions from  EPHA_M_RAM_COLOR a where active_type = 1 order by sort_by ";

                //DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
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
                    dt.TableName = "ram_color";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

            }
            #endregion
        }
        public string employees_list(EmployeeListModel param)
        {
            if (param == null)
            {
                return JsonConvert.SerializeObject(new { message = "Invalid parameter." }, Formatting.Indented);
            }

            try
            {
                DataSet dsData = new DataSet();

                // ถ้าไม่มีชื่อใน list, ตั้งค่า maxRows เป็น 100
                int maxRows = param.user_name_list.Count == 0 ? 100 : param.user_name_list.Count;

                // จำกัดค่า maxRows ไม่ให้มากเกินไป เช่น ไม่เกิน 1000
                maxRows = Math.Min(maxRows, 1000);

                // สร้าง DataTable สำหรับ UserNameList (Table-Valued Parameter)
                DataTable userNameListTable = new DataTable();
                userNameListTable.Columns.Add("UserName", typeof(string));

                // ใส่ค่ารายชื่อผู้ใช้ลงใน DataTable
                foreach (var name in param.user_name_list)
                {
                    userNameListTable.Rows.Add(name);
                }

                // สร้างพารามิเตอร์
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@maxRows", SqlDbType.Int) { Value = maxRows });
                parameters.Add(new SqlParameter("@userNameList", SqlDbType.Structured) { Value = userNameListTable, TypeName = "EPHA_M_USERNAMELIST" });

                // เรียกใช้ Stored Procedure
                DataTable dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable("usp_GetPersonDetails", parameters, isStoredProcedure: true);
                #region Execute to Datable
                //var parameters = new List<SqlParameter>();
                try
                {
                    _conn = new ClassConnectionDb();
                    _conn = new ClassConnectionDb(); _conn.OpenConnection();
                    try
                    {
                        var command = _conn.conn.CreateCommand();
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = "usp_GetPersonDetails";
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

                if (dt == null)
                {
                    return JsonConvert.SerializeObject(new { message = "No data found." }, Formatting.Indented);
                }
                else
                {
                    if (dt?.Rows.Count == 0)
                    {
                        dt.Rows.Add(dt.NewRow());
                        dt.AcceptChanges();
                    }
                    if (dt?.Rows.Count > 0)
                    {
                        dt.TableName = "employee";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                    }
                    string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);
                    return json;
                }
            }
            catch (Exception ex_error) { return ex_error.Message.ToString(); }
        }

        public string employees_search(EmployeeModel param)
        {
            if (param == null)
            {
                return JsonConvert.SerializeObject(new { message = "Invalid parameter." }, Formatting.Indented);
            }

            try
            {
                DataSet dsData = new DataSet();
                int maxRows = string.IsNullOrEmpty(param.max_rows) ? 100 : Math.Min(int.Parse(param.max_rows), 1000); // จำกัดค่า maxRows
                string userIndicator = (param.user_indicator ?? "");
                string userFilterText = (param.user_filter_text ?? "");

                var parameters = new List<SqlParameter>();

                //if (!string.IsNullOrEmpty(userIndicator))
                //{
                //    cls = new ClassFunctions();
                //    userIndicator = cls.ChkSqlStr(userIndicator, 4000);
                //}
                //if (!string.IsNullOrEmpty(userFilterText))
                //{
                //    cls = new ClassFunctions();
                //    userFilterText = cls.ChkSqlStr(userFilterText, 4000);
                //}

                sqlstr = "select top (@maxRows) seq, seq as id, user_id as employee_id, user_name as employee_name, user_displayname as employee_displayname, " +
                         "user_email as employee_email, t.user_title as employee_position, 'assets/img/team/avatar.webp' as employee_img, user_type as employee_type, 0 as selected_type " +
                         "from VW_EPHA_PERSON_DETAILS t where t.user_displayname is not null";

                parameters.Add(new SqlParameter("@maxRows", SqlDbType.Int) { Value = maxRows });

                if (!string.IsNullOrEmpty(userIndicator) && !string.IsNullOrEmpty(userFilterText))
                {
                    sqlstr += @" and (
                                    trim(lower(t.user_displayname)) like '%' + replace(lower(@userIndicator), ' ', '') + '%' or trim(lower(t.user_title)) like '%' + replace(lower(@userIndicator), ' ', '') + '%' or
                                    trim(lower(t.user_displayname)) like '%' + replace(lower(@userFilterText), ' ', '') + '%' or trim(lower(t.user_title)) like '%' + replace(lower(@userFilterText), ' ', '') + '%' or
                                    trim(lower(t.user_title + t.user_displayname)) like '%' + replace(lower(@userIndicator), ' ', '') + '%' or trim(lower(t.user_title + t.user_displayname)) like '%' + replace(lower(@userFilterText), ' ', '') + '%'
                                )";
                    parameters.Add(new SqlParameter("@userIndicator", SqlDbType.VarChar, 4000) { Value = userIndicator });
                    parameters.Add(new SqlParameter("@userFilterText", SqlDbType.VarChar, 4000) { Value = userFilterText });
                }
                else if (!string.IsNullOrEmpty(userIndicator))
                {
                    sqlstr += @" and (
                                    trim(lower(t.user_displayname)) like '%' + replace(lower(@userIndicator), ' ', '') + '%' or trim(lower(t.user_title)) like '%' + replace(lower(@userIndicator), ' ', '') + '%' or
                                    trim(lower(t.user_title + t.user_displayname)) like '%' + replace(lower(@userIndicator), ' ', '') + '%' or
                                    trim(lower(t.user_title + t.user_name)) like '%' + replace(lower(@userIndicator), ' ', '') + '%'or
                                    trim(lower(t.user_name)) like '%' + replace(lower(@userIndicator), ' ', '') + '%'
                                )";
                    parameters.Add(new SqlParameter("@userIndicator", SqlDbType.VarChar, 4000) { Value = userIndicator });
                }
                else if (!string.IsNullOrEmpty(userFilterText))
                {
                    sqlstr += @" and (
                                     trim(lower(t.user_displayname)) like '%' + replace(lower(@userFilterText), ' ', '') + '%' or trim(lower(t.user_title)) like '%' + replace(lower(@userFilterText), ' ', '') + '%' or
                                     trim(lower(t.user_title + t.user_displayname)) like '%' + replace(lower(@userFilterText), ' ', '') + '%' or
                                     trim(lower(t.user_name)) like '%' + replace(lower(@userFilterText), ' ', '') + '%'
                                 )";
                    parameters.Add(new SqlParameter("@userFilterText", SqlDbType.VarChar, 4000) { Value = userFilterText });
                }

                sqlstr += " order by user_name";

                DataTable dt = new DataTable();
                //DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                #region Execute to Datable
                //var parameters = new List<SqlParameter>();
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

                if (dt == null)
                {
                    return JsonConvert.SerializeObject(new { message = "No data found." }, Formatting.Indented);
                }
                else
                {
                    if (dt?.Rows.Count == 0)
                    {
                        dt.Rows.Add(dt.NewRow());
                        dt.AcceptChanges();
                    }
                    if (dt?.Rows.Count > 0)
                    {
                        dt.TableName = "employee";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                    }
                    string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);
                    return json;
                }
            }
            catch (Exception ex_error) { return ex_error.Message.ToString(); }
        }

        private void get_authorization_page(ref DataSet _dsData, string user_name, string role_type)
        {
            // ตรวจสอบค่า เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (_dsData == null || string.IsNullOrEmpty(user_name))
            {
                return;
            }

            ClassLogin clsLogin = new ClassLogin();
            DataTable dtAuthPage = clsLogin._dtAuthorization_Page(user_name, "");
            if (dtAuthPage != null)
            {
                dtAuthPage.TableName = "authorization_page";
                _dsData.Tables.Add(dtAuthPage.Copy()); _dsData.AcceptChanges();
            }
        }
        private void authorization_page_by_doc(ref DataSet _dsData, string user_name, string role_type)
        {
            // ตรวจสอบค่า  เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (_dsData == null || string.IsNullOrEmpty(user_name))
            {
                return;
            }

            ClassLogin clsLogin = new ClassLogin();
            DataTable dtAuthPage = clsLogin._dtAuthorization_Page_By_Doc(user_name, role_type, "");
            if (dtAuthPage != null)
            {
                dtAuthPage.TableName = "authorization_page_by_doc";
                _dsData.Tables.Add(dtAuthPage.Copy()); _dsData.AcceptChanges();
            }
        }
        public void get_employee_list(Boolean bCleare, ref DataSet _dsData)
        {
            string sqlstr = @"
               select seq, seq as id, user_id as employee_id, user_name as employee_name, user_displayname as employee_displayname, user_email as employee_email,
               t.user_title as employee_position, 
               'assets/img/team/avatar.webp' as employee_img, user_type as employee_type, 
               0 as selected_type
               from VW_EPHA_PERSON_DETAILS t 
               where t.user_displayname is not null 
               order by user_name";

            //DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            DataTable dt = new DataTable();
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


            if (bCleare)
            {
                if (dt == null || dt?.Rows.Count == 0)
                {
                    dt.Rows.Add(dt.NewRow());
                    dt.AcceptChanges();
                }
                else { dt.Rows.Clear(); }
            }
            if (dt != null)
            {
                dt.TableName = "employee";
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
            }
        }

        private void get_master(ref DataSet _dsData, string sub_software, string page_name)
        {
            // ตรวจสอบค่า  เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (_dsData == null || string.IsNullOrEmpty(sub_software))
            {
                return;
            }

            var parameters = new List<SqlParameter>();
            DataTable dt = new DataTable();
            get_employee_list(true, ref _dsData);

            #region company
            sqlstr = @"SELECT seq AS id, name FROM EPHA_M_COMPANY ORDER BY id";
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            #region Execute to Datable
            parameters = new List<SqlParameter>();
            try
            {
                _conn = new ClassConnectionDb();
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
                dt.TableName = "company";
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
            }

            #endregion company

            #region master ram
            if (!(sub_software == "hra"))
            {
                get_master_ram(ref _dsData);
            }
            #endregion ram

            #region master apu
            sqlstr = page_name == "followup"
                ? @"SELECT DISTINCT a.id, a.name, a.id AS id_area, a.id AS id_apu, LOWER(a.name) AS area_check, c.seq AS id_company FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY a.name"
                : @"SELECT DISTINCT a.id, a.name, a.id AS id_area, a.id AS id_apu, LOWER(a.name) AS area_check FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY a.name";
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
                dt.TableName = "apu";
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
            }

            sqlstr = @"SELECT DISTINCT a.name AS id, a.name AS name, c.seq AS id_company FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY a.name";
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
                dt.TableName = "apu_filter";
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
            }

            #endregion apu

            #region master unit no 
            string unitNoSql = sub_software.ToLower() switch
            {
                "hazop" => @"SELECT DISTINCT a.id AS id_area, a.id AS id_apu, t.id, t.id, t.descriptions AS name, LOWER(a.name) AS area_check, t.id_company FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY t.descriptions",
                "whatif" => @"SELECT DISTINCT a.id AS id_area, a.id AS id_apu, t.id, t.id, t.descriptions AS name, LOWER(a.name) AS area_check, t.id_company FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY t.descriptions",
                "jsea" => @"SELECT DISTINCT a.id AS id_area, a.id AS id_apu, t.id, t.id, t.descriptions AS name, LOWER(a.name) AS area_check, t.id_company, t.id_plant_area FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY t.descriptions",
                "hra" => @"SELECT DISTINCT a.id AS id_area, a.id AS id_apu, t.id, t.id, (t.descriptions + '-' + pa.name) AS name, LOWER(a.name) AS area_check, t.id_company, t.id_plant_area FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY (t.descriptions + '-' + pa.name)",
                _ => @"SELECT DISTINCT a.id AS id_area, a.id AS id_apu, t.id, t.id, t.descriptions AS name, LOWER(a.name) AS area_check, t.id_company, t.id_plant_area FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY t.descriptions"
            };
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(unitNoSql, null); 
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
                dt.TableName = "unit_no";
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
            }

            #endregion unit no

            #region master request_type
            sqlstr = @"SELECT DISTINCT a.id, a.name FROM EPHA_M_REQUEST_TYPE a WHERE lower(pha_sub_software) = lower(@subSoftware) ORDER BY a.id";
            parameters = new List<SqlParameter> { new SqlParameter("@subSoftware", SqlDbType.VarChar, 100) { Value = sub_software.ToLower() } };
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters); 
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
                dt.TableName = "request_type";
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
            }

            #endregion master request_type

            if (sub_software.ToLower() == "hazop" || sub_software.ToLower() == "whatif")
            {
                #region master functional location
                sqlstr = @"SELECT *, a.functional_location AS id, a.functional_location AS name, 0 AS selected_type FROM EPHA_M_FUNCTIONAL_LOCATION a WHERE active_type = 1 ORDER BY seq";
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
                    dt.TableName = "functional";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }
                #endregion functional location  

                #region master business unit
                sqlstr = @"SELECT DISTINCT a.id AS id_area, a.id AS id_apu, t.id, t.name + ' : ' + t.descriptions + '(' + CONVERT(VARCHAR, a.name) + '-' + CONVERT(VARCHAR, pa.name) + ')' AS name, LOWER(a.name) AS area_check FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY a.id, t.id";
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
                    dt.TableName = "business_unit";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion business unit

                #region master guidwords  
                sqlstr = @"SELECT seq, parameter, deviations, guide_words, guide_words AS guidewords, process_deviation, area_application, 0 AS selected_type, 0 AS main_parameter, def_selected, no, g.no_guide_words AS guidewords_no, g.no_deviations AS deviations_no FROM EPHA_M_GUIDE_WORDS g WHERE active_type = 1 ORDER BY seq, parameter, deviations, guide_words, process_deviation, area_application";
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
                    else
                    {
                        //sort data 
                        string beforeParameter = "";
                        string afterParameter = "";
                        for (int i = 0; i < dt?.Rows.Count; i++)
                        {
                            beforeParameter = dt.Rows[i]["parameter"]?.ToString() ?? "";
                            if (beforeParameter != afterParameter)
                            {
                                afterParameter = beforeParameter;
                                dt.Rows[i]["main_parameter"] = 1;
                                dt.AcceptChanges();
                            }
                        }
                        if (beforeParameter != afterParameter)
                        {
                            int icount_row = dt?.Rows.Count ?? 0;
                            afterParameter = beforeParameter;
                            dt.Rows[icount_row]["main_parameter"] = 1;
                            dt.AcceptChanges();
                        }
                    }
                }
                if (dt != null)
                {
                    dt.TableName = "guidwords";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion guidwords  
            }
            else if (sub_software.ToLower() == "jsea")
            {
                #region master business unit
                sqlstr = @"SELECT DISTINCT a.id AS id_area, a.id AS id_apu, t.id, t.name + ' : ' + t.descriptions + '(' + CONVERT(VARCHAR, a.name) + '-' + CONVERT(VARCHAR, pa.name) + ')' AS name, LOWER(a.name) AS area_check FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY a.id, t.id";
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
                    dt.TableName = "business_unit";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }
                #endregion business unit

                #region Complex
                sqlstr = @"SELECT t.id_company, t.id, t.name, a.id AS id_area, a.id AS id_apu, LOWER(a.name) AS area_check FROM epha_m_area_complex t INNER JOIN EPHA_M_AREA a ON t.id_area = a.id ORDER BY t.id_company, t.id";
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
                    dt.TableName = "toc";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }
                #endregion Complex

                #region Tag ID
                sqlstr = @"SELECT id_company, id_apu, id_area, id, name, LOWER(t.name) AS area_check FROM EPHA_M_TAGID t ORDER BY id_company, id_apu, id";
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
                    dt.TableName = "tagid";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion Tag ID 

                #region Departments
                sqlstr = @"SELECT DISTINCT departments AS id, functions + '-' + departments AS name, LOWER(a.departments) AS text_check FROM VW_EPHA_PERSON_DETAILS a WHERE ISNULL(functions, '') <> '' AND ISNULL(departments, '') <> '' ORDER BY functions + '-' + departments";
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
                    dt.TableName = "departments";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion Departments

                #region Mandatory Note
                sqlstr = @"SELECT DISTINCT a.id, a.name, LOWER(a.name) AS text_check, a.active_def FROM EPHA_M_MANDATORY_NOTE a WHERE a.active_type = 1 ORDER BY a.active_def DESC, a.id";
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
                    dt.TableName = "mandatory_note";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }
                #endregion Mandatory Note
            }
            else if (sub_software.ToLower() == "hra")
            {
                #region Departments
                sqlstr = @"SELECT DISTINCT emp.departments AS id, emp.departments AS name, emp.functions, emp.departments FROM vw_epha_person_details emp WHERE emp.departments IS NOT NULL ORDER BY emp.functions, emp.departments";
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
                    dt.TableName = "departments";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }
                #endregion Departments

                #region Sections
                sqlstr = @"SELECT DISTINCT emp.sections AS id, emp.sections AS name, emp.functions, emp.departments, emp.sections FROM vw_epha_person_details emp WHERE emp.departments IS NOT NULL AND emp.sections IS NOT NULL ORDER BY emp.functions, emp.departments, emp.sections";
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
                    dt.TableName = "sections";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion Sections

                #region Complex
                sqlstr = @"SELECT t.id_company, t.id, t.name, a.id AS id_area, a.id AS id_apu, LOWER(a.name) AS area_check FROM epha_m_area_complex t INNER JOIN EPHA_M_AREA a ON t.id_area = a.id ORDER BY t.id_company, t.id";
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
                    dt.TableName = "toc";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion Complex

                #region subarea
                sqlstr = @"SELECT a.id, a.name, LOWER(a.name) AS field_check FROM epha_m_sections_group a WHERE a.active_type = 1 ORDER BY a.id";
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
                    dt.TableName = "subarea";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }


                sqlstr = @"SELECT a.id, a.name, LOWER(a.name) AS field_check FROM epha_m_sections_group a WHERE a.active_type = 1 ORDER BY a.id";
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
                    dt.TableName = "sections_group";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                sqlstr = @"SELECT a.id, b.name, LOWER(b.name) AS field_check, a.descriptions, a.id_sections FROM epha_m_sub_area a INNER JOIN epha_m_sections_group b ON a.id_sections_group = b.id WHERE a.active_type = 1 ORDER BY a.name";
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
                    dt.TableName = "subarea_location";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion subarea

                #region standard_type
                sqlstr = @"SELECT DISTINCT a.id_hazard_type, NULL AS id, a.standard_type_text AS name, LOWER(a.standard_type_text) AS field_check FROM epha_m_hazard_riskfactors a WHERE a.active_type = 1 ORDER BY a.standard_type_text";
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
                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    dt.Rows[i]["id"] = (i + 1);
                    dt.AcceptChanges();
                }
                if (dt != null)
                {
                    dt.TableName = "standard_type";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion standard_type

                #region hazard_type
                sqlstr = @"SELECT a.id, a.name, LOWER(a.name) AS field_check, a.descriptions FROM epha_m_hazard_type a WHERE a.active_type = 1 ORDER BY a.id";
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
                    dt.TableName = "hazard_type";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }
                #endregion hazard_riskfactors

                #region hazard_riskfactors
                sqlstr = @"SELECT DISTINCT a.id_hazard_type, NULL AS id, a.health_hazards AS name, LOWER(a.health_hazards) AS field_check FROM epha_m_hazard_riskfactors a WHERE a.active_type = 1 ORDER BY a.health_hazards";
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
                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        dt.Rows[i]["id"] = i;
                    }
                    dt.TableName = "hazard_riskfactors";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                sqlstr = @"SELECT a.id_hazard_type, a.id, a.health_hazards AS name, LOWER(a.health_hazards) AS field_check, a.hazards_rating, a.standard_value, a.standard_unit, a.standard_desc, a.standard_type_text FROM epha_m_hazard_riskfactors a WHERE a.active_type = 1 ORDER BY a.id_hazard_type, a.health_hazards";
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
                    dt.TableName = "hazard_standard";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion hazard_riskfactors

                #region worker_group
                sqlstr = @"SELECT a.id_business_unit, a.id, a.name, LOWER(a.name) AS field_check FROM epha_m_worker_group a WHERE a.active_type = 1 ORDER BY a.id_business_unit, a.id";
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
                    dt.TableName = "worker_group";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion worker_group

                #region worker_list
                sqlstr = @"SELECT a.id_worker_group, a.id, a.user_displayname AS name, a.user_name, a.user_type, a.user_displayname FROM epha_m_worker_list a WHERE a.active_type = 1 ORDER BY a.id_worker_group, a.user_displayname";
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
                    dt.TableName = "worker_list";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion worker_list

                #region worker_task
                sqlstr = @"SELECT a.id_unit_no, a.id, a.name, LOWER(a.name) AS field_check FROM epha_m_worker_task a WHERE a.active_type = 1 ORDER BY a.id";
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
                    dt.TableName = "worker_task";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion worker_task

                #region activities
                sqlstr = @"SELECT a.id, a.name, LOWER(a.name) AS area_check FROM epha_m_activities a ORDER BY a.name";
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
                    dt.TableName = "activities";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }


                sqlstr = @"SELECT a.id, a.name, LOWER(a.name) AS text_check FROM epha_m_frequency_level a ORDER BY a.name";
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
                    dt.TableName = "frequency_level";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }


                sqlstr = @"SELECT a.id, a.name, LOWER(a.name) AS text_check FROM epha_m_exposure_level a ORDER BY a.name";
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
                    dt.TableName = "exposure_level";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }


                sqlstr = @"SELECT a.frequency_level, a.exposure_level, a.results, a.results_desc FROM epha_m_compare_exposure_rating a ORDER BY a.results";
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
                    dt.TableName = "compare_exposure_rating";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                sqlstr = @"SELECT a.health_effect_rating, a.exposure_rating, a.results, a.results_desc FROM epha_m_compare_initial_risk_rating a ORDER BY a.results";
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
                    dt.TableName = "compare_initial_risk_rating";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion activities

                #region recommendations follow
                sqlstr = @"SELECT a.id, a.name, LOWER(a.name) AS text_check FROM epha_m_rangtype a ORDER BY a.name";
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
                    dt.TableName = "rangtype";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion recommendations follow
            }
            else if (sub_software.ToLower() == "followup" || sub_software.ToLower() == "search")
            {
                #region master functional location
                sqlstr = @"SELECT *, a.functional_location AS id, a.functional_location AS name, 0 AS selected_type FROM EPHA_M_FUNCTIONAL_LOCATION a WHERE active_type = 1 ORDER BY seq";
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
                    dt.TableName = "functional";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }

                #endregion functional location  

                #region master business unit
                sqlstr = @"SELECT DISTINCT a.id AS id_area, a.id AS id_apu, t.id, t.name + ' : ' + t.descriptions + '(' + CONVERT(VARCHAR, a.name) + '-' + CONVERT(VARCHAR, c.plant) + ')' AS name, LOWER(a.name) AS area_check FROM  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area  ORDER BY a.id, t.id";
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
                    dt.TableName = "business_unit";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }
                #endregion business unit

                #region master guidwords  
                sqlstr = @"SELECT seq, parameter, deviations, guide_words, guide_words AS guidewords, process_deviation, area_application, 0 AS selected_type, 0 AS main_parameter, def_selected, no, g.no_guide_words AS guidewords_no, g.no_deviations AS deviations_no FROM EPHA_M_GUIDE_WORDS g WHERE active_type = 1 ORDER BY seq, parameter, deviations, guide_words, process_deviation, area_application";
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
                    else
                    {
                        //sort data 
                        string beforeParameter = "";
                        string afterParameter = "";
                        for (int i = 0; i < dt?.Rows.Count; i++)
                        {
                            beforeParameter = dt.Rows[i]["parameter"]?.ToString() ?? "";
                            if (beforeParameter != afterParameter)
                            {
                                afterParameter = beforeParameter;
                                dt.Rows[i]["main_parameter"] = 1;
                                dt.AcceptChanges();
                            }
                        }
                        if (beforeParameter != afterParameter)
                        {
                            afterParameter = beforeParameter;
                            int icount_row = dt?.Rows.Count ?? 0;

                            dt.Rows[icount_row]["main_parameter"] = 1;
                            dt.AcceptChanges();
                        }
                    }
                }
                if (dt != null)
                {
                    dt.TableName = "guidwords";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                }
                #endregion guidwords  
            }
        }

        private void get_master_search(ref DataSet _dsData, string sub_software, string user_name)
        {

            if (string.IsNullOrEmpty(sub_software))
            {
                throw new ArgumentException("Invalid sub_software value.");
            }
            // กำหนด whitelist ของ software ที่อนุญาต
            var allowedSoftwares = new List<string> { "hazop", "jsea", "whatif", "hra" };

            if (!allowedSoftwares.Contains(sub_software.ToLower()))
            {
                throw new ArgumentException("Invalid sub_software value.");
            }

            string role_type = "";
            check_role_user_active(user_name, ref role_type);

            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            #region master status 
            sqlstr = @"select distinct a.id, a.descriptions as name 
               from EPHA_M_STATUS a where active_type = 1  
               order by a.id";

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

            dt.TableName = "status";
            _dsData.Tables.Add(dt.Copy());
            _dsData.AcceptChanges();
            #endregion master status

            #region master apu 
            sqlstr = @"select distinct a.id, a.name as name, a.id as id_area, a.id as id_apu, lower(a.name) as area_check 
               from EPHA_M_AREA a 
               inner join EPHA_T_GENERAL g on g.id_apu = a.id
               inner join EPHA_T_HEADER h on h.id = g.id_pha 
               inner join VW_EPHA_DATA_DOC_BY_USER du on h.id = du.id_pha 
               where lower(h.pha_sub_software) = lower(@sub_software)";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });

            if (role_type != "admin")
            {
                sqlstr += @" and lower(du.user_name)  = lower(@user_name)";
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
            }
            sqlstr += @" order by a.name";

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

            dt.TableName = "apu";
            _dsData.Tables.Add(dt.Copy());
            _dsData.AcceptChanges();
            #endregion master apu

            #region master unit no  

            sqlstr = @"select distinct t.id, t.descriptions as name, a.id as id_area, a.id as id_apu, lower(a.name) as area_check, t.id_company 
               from  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area 
               inner join EPHA_T_GENERAL g on g.id_apu = a.id and g.id_unit_no = t.id 
               inner join EPHA_T_HEADER h on h.id = g.id_pha 
               inner join VW_EPHA_DATA_DOC_BY_USER du on h.id = du.id_pha 
               where lower(h.pha_sub_software) = lower(@sub_software)";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });

            if (role_type != "admin")
            {
                sqlstr += @" and lower(du.user_name)  = lower(@user_name)";
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
            }
            sqlstr += @" order by t.descriptions";

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

            dt.TableName = "unit_no";
            _dsData.Tables.Add(dt.Copy());
            _dsData.AcceptChanges();
            #endregion master unit no

            if ((sub_software ?? "").ToLower() == "hazop" || (sub_software ?? "").ToLower() == "whatif")
            {
                #region master functional location

                sqlstr = @"select distinct g.functional_location as id, g.functional_location as name, 0 as selected_type
                   from  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area 
                   inner join EPHA_T_GENERAL g on g.id_apu = a.id and g.functional_location is not null
                   inner join EPHA_T_HEADER h on h.id = g.id_pha 
                   inner join VW_EPHA_DATA_DOC_BY_USER du on h.id = du.id_pha 
                   where lower(h.pha_sub_software) = lower(@sub_software)";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });

                if (role_type != "admin")
                {
                    sqlstr += @" and lower(du.user_name)  = lower(@user_name)";
                    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
                }
                sqlstr += @" order by g.functional_location ";

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


                dt.TableName = "functional";
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
                #endregion functional location  

                #region master business unit 
                sqlstr = @"select distinct a.id, t.name as name, a.id as id_area, a.id as id_apu, lower(a.name) as area_check 
                   from  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area 
                   inner join EPHA_T_GENERAL g on g.id_apu = a.id and g.id_business_unit = t.id
                   inner join EPHA_T_HEADER h on h.id = g.id_pha 
                   inner join VW_EPHA_DATA_DOC_BY_USER du on h.id = du.id_pha 
                   where lower(h.pha_sub_software) = lower(@sub_software)";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });

                if (role_type != "admin")
                {
                    sqlstr += @" and lower(du.user_name)  = lower(@user_name)";
                    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
                }
                sqlstr += @" order by t.name ";

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


                dt.TableName = "business_unit";
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
                #endregion business unit

                #region approver
                sqlstr = @"select distinct h.approver_user_name as id, h.approver_user_name as name 
                   from  EPHA_M_COMPANY c inner join EPHA_M_BUSINESS_UNIT t on c.seq = t.id_company  inner join epha_m_area_complex pa on c.seq = pa.id_company and pa.id = t.id_plant_area inner join EPHA_M_AREA a on a.id = pa.id_area and  a.id = t.id_area 
                   inner join EPHA_T_GENERAL g on g.id_apu = a.id and g.id_unit_no = t.id 
                   inner join EPHA_T_HEADER h on h.id = g.id_pha and h.approver_user_name is not null
                   inner join VW_EPHA_DATA_DOC_BY_USER du on h.id = du.id_pha  
                   where lower(h.pha_sub_software) = lower(@sub_software)";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });

                if (role_type != "admin")
                {
                    sqlstr += @" and lower(du.user_name)  = lower(@user_name)";
                    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
                }
                sqlstr += @" order by h.approver_user_name";

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

                dt.TableName = "approver";
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
                #endregion approver
            }

            if (sub_software.ToLower() == "jsea" || sub_software.ToLower() == "hra" || sub_software.ToLower() == "followup" || sub_software.ToLower() == "search")
            {
                #region company
                sqlstr = @"select seq as id, name from EPHA_M_COMPANY t order by id";

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
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
                #endregion company

                #region Complex
                sqlstr = @"select t.id_company, t.id, t.name, a.id as id_area, a.id as id_apu, lower(a.name) as area_check 
                   from epha_m_area_complex t
                   inner join EPHA_M_AREA a on t.id_area = a.id 
                   order by t.id_company, t.id";

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
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
                #endregion Complex

                #region Tag ID
                sqlstr = @"select id_company, id_apu, id_area, id, name from EPHA_M_TAGID t order by id_company, id_apu, id";

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

                dt.TableName = "tagid";
                _dsData.Tables.Add(dt.Copy());
                _dsData.AcceptChanges();
                #endregion Tag ID

                if (sub_software.ToLower() == "hra")
                {
                    #region master request_type
                    parameters = new List<SqlParameter>();

                    parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 100) { Value = sub_software.ToLower() });

                    sqlstr = @"select distinct a.id, a.name from EPHA_M_REQUEST_TYPE a where pha_sub_software = @sub_software order by a.id";

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

                    dt.TableName = "request_type";
                    _dsData.Tables.Add(dt.Copy());
                    _dsData.AcceptChanges();
                    #endregion master request_type
                }
            }
        }

        public void get_data_search(ref DataSet dsData, string user_name, string seq, string sub_software)
        {
            // ตรวจสอบค่า เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (dsData == null || string.IsNullOrEmpty(user_name) || string.IsNullOrEmpty(sub_software))
            {
                return;
            }
            var pha_sub_software = sub_software;
            // Define a whitelist of allowed sub_software values
            var allowedSubSoftware = new HashSet<string> { "jsea", "whatif", "hra", "hazop" };

            // Check if sub_software is valid
            if (!allowedSubSoftware.Contains(pha_sub_software.ToLower()))
            {
                return;
            }

            sub_software = sub_software ?? "";
            if (user_name.IndexOf("@") > -1)
            {
                user_name = user_name.Split("@")[0];
            }

            string role_type = "";
            check_role_user_active(user_name, ref role_type);

            DataTable dtma = new DataTable();
            int id_pha = 0;
            string year_now = System.DateTime.Now.Year.ToString();
            if (Convert.ToInt64(year_now) > 2500) { year_now = (Convert.ToInt64(year_now) - 543).ToString(); }

            DataTable dt = new DataTable();
            cls = new ClassFunctions();

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name.ToLower() ?? "x" });

            sqlstr = @"select * from VW_EPHA_PERSON_DETAILS a 
               where a.seq in (select max(seq) from vw_epha_max_seq_by_pha_no group by pha_no) 
               and lower(a.user_name) = lower(@user_name)";

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

            if (dt != null)
            {
                dt.TableName = "employee_list";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }

            #region conditions 
            sqlstr = @"select b.*, '' as functional_location_audition, '' as business_unit_name, '' as unit_no_name, 'update' as action_type, 0 as action_change
               , '' as emp_active_search
               from epha_t_header a inner join EPHA_T_GENERAL b on a.id = b.id_pha
               where 1=2 ";

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
                else
                {
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["seq"] = id_pha;
                    dt.Rows[0]["id"] = id_pha;
                    dt.Rows[0]["id_pha"] = id_pha;
                    dt.Rows[0]["functional_location_audition"] = "";
                    dt.Rows[0]["id_ram"] = 5;
                    dt.Rows[0]["expense_type"] = "OPEX";
                    dt.Rows[0]["sub_expense_type"] = "Normal";
                    dt.Rows[0]["create_by"] = user_name;
                    dt.Rows[0]["action_type"] = "insert";
                    dt.Rows[0]["action_change"] = 0;
                    dt.AcceptChanges();
                }

            }
            if (dt != null)
            {
                dt.TableName = "conditions";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }

            #endregion conditions

            #region results  

            sqlstr = @"select distinct seq, emp_name from (
                select h.id as seq, sm.pha_sub_software, pha_request_by as emp_name 
                from EPHA_T_HEADER h 
                inner join VW_EPHA_MAX_SEQ_BY_PHA_NO sm on lower(h.id) = lower(sm.id_pha)  
                left join VW_EPHA_PERSON_DETAILS vw on lower(h.pha_request_by) = lower(vw.user_name)
                union
                select a.id_pha as seq, sm.pha_sub_software, a.user_displayname as emp_name from EPHA_T_MEMBER_TEAM a  
                inner join VW_EPHA_MAX_SEQ_BY_PHA_NO sm on lower(a.id_pha) = lower(sm.id_pha)  
                union
                select a.id_pha as seq, sm.pha_sub_software, a.user_displayname as emp_name from EPHA_T_APPROVER a  
                inner join VW_EPHA_MAX_SEQ_BY_PHA_NO sm on lower(a.id_pha) = lower(sm.id_pha) 
                union
                select a.id_pha as seq, sm.pha_sub_software, a.user_displayname as emp_name 
                from (select ta3.user_name, ta3.user_displayname, ta3.id_pha from epha_t_approver_ta3 ta3 inner join epha_t_approver ta2 on ta3.id_approver = ta2.id) a  
                inner join VW_EPHA_MAX_SEQ_BY_PHA_NO sm on lower(a.id_pha) = lower(sm.id_pha) 
                union
                select a.id_pha as seq, sm.pha_sub_software, a.user_displayname as emp_name from EPHA_T_RELATEDPEOPLE a 
                inner join VW_EPHA_MAX_SEQ_BY_PHA_NO sm on lower(a.id_pha) = lower(sm.id_pha) 
                union
                select a.id_pha as seq, sm.pha_sub_software, a.user_displayname as emp_name from EPHA_T_RELATEDPEOPLE_OUTSIDER a  
                inner join VW_EPHA_MAX_SEQ_BY_PHA_NO sm on lower(a.id_pha) = lower(sm.id_pha) 
                ) t where emp_name is not null ";

            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " and t.seq = @seq";
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq?.ToLower() ?? "" });
            }
            if (!string.IsNullOrEmpty(sub_software))
            {
                sqlstr += " and lower(t.pha_sub_software) = lower(@sub_software) ";
                parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software?.ToLower() ?? "" });
            }
            DataTable dtEmpActive = new DataTable();
            //dtEmpActive = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                    dtEmpActive = new DataTable();
                    dtEmpActive = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dtEmpActive.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable



            parameters = new List<SqlParameter>();
            sqlstr = @"select distinct seq, search_desc from (
                select a.id_pha as seq, sm.pha_sub_software, a.potentailhazard + possiblecase as search_desc 
                from EPHA_T_TASKS_WORKSHEET a  
                inner join VW_EPHA_MAX_SEQ_BY_PHA_NO sm on lower(a.id_pha) = lower(sm.id_pha)  
                ) t where search_desc is not null";
            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " and t.seq = @seq";
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq?.ToLower() ?? "" });
            }
            if (!string.IsNullOrEmpty(sub_software))
            {
                sqlstr += " and lower(t.pha_sub_software) = lower(@sub_software) ";
                parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software?.ToLower() ?? "" });
            }
            DataTable dtWorksheetActive = new DataTable();
            //dtWorksheetActive = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                    dtWorksheetActive = new DataTable();
                    dtWorksheetActive = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "data";
                    dtWorksheetActive.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            parameters = new List<SqlParameter>();
            sqlstr = @"select ms.name as pha_status_name, ms.descriptions as pha_status_displayname
               ,case when a.year = year(getdate()) then vw.user_name else a.request_user_name end request_user_name
               ,case when a.year = year(getdate()) then vw.user_name else a.request_user_displayname end request_user_displayname
               ,null as approver_user_img
               , 'update' as action_type, 0 as action_change  
               , '' as emp_active_search, '' as worksheet_active_search
               , b.id_request_type, rt.name as request_type
               , a.*, b.*
               from EPHA_T_HEADER a
               inner join EPHA_T_GENERAL b on a.id = b.id_pha
               left join EPHA_M_STATUS ms on a.pha_status = ms.id
               left join VW_EPHA_PERSON_DETAILS vw on lower(a.pha_request_by) = lower(vw.user_name)
               left join EPHA_M_REQUEST_TYPE rt on b.id_request_type = rt.id 
               where 1=1 and a.seq in (select max(seq) from vw_epha_max_seq_by_pha_no group by pha_no) ";
            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " and a.seq = @seq ";
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq?.ToLower() ?? "" });
            }
            if (!string.IsNullOrEmpty(sub_software))
            {
                sqlstr += " and lower(a.pha_sub_software) = lower(@sub_software) ";
                parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software?.ToLower() ?? "" });
            }

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


            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                else
                {
                    dt.Rows.Add(dt.NewRow());
                    dt.AcceptChanges();
                }
            }
            if (dt != null && dtEmpActive != null)
            {
                if (dt?.Rows.Count > 0 && dtEmpActive?.Rows.Count > 0)
                {
                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        string emp_def = "";
                        string seq_def = dt.Rows[i]["seq"]?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(seq_def))
                        {
                            //DataRow[] drEmpActive = dtEmpActive.Select("seq = '" + seq_def + "'");
                            var filterParameters = new Dictionary<string, object>();
                            filterParameters.Add("seq", seq_def);
                            var (drEmpActive, iEmpActive) = FilterDataTable(dtEmpActive, filterParameters);
                            if (drEmpActive != null)
                            {
                                if (drEmpActive?.Length > 0)
                                {
                                    foreach (DataRow dr in drEmpActive)
                                    {
                                        emp_def += dr["emp_name"] + ";";
                                    }
                                    dt.Rows[i]["emp_active_search"] = emp_def.ToLower();
                                }
                            }

                        }
                    }
                }

                if (sub_software.ToLower() == "jsea" && dt?.Rows.Count > 0 && dtWorksheetActive?.Rows.Count > 0)
                {
                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        string val_def = "";
                        string seq_def = dt.Rows[i]["seq"]?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(seq_def))
                        {
                            //DataRow[] drValActive = dtWorksheetActive.Select("seq = '" + seq_def + "'");
                            var filterParameters = new Dictionary<string, object>();
                            filterParameters.Add("seq", seq_def);
                            var (drValActive, iValActive) = FilterDataTable(dtWorksheetActive, filterParameters);
                            if (drValActive != null)
                            {
                                if (drValActive?.Length > 0)
                                {
                                    foreach (DataRow dr in drValActive)
                                    {
                                        val_def += dr["search_desc"] + ";";
                                    }
                                    dt.Rows[i]["worksheet_active_search"] = val_def.ToLower();
                                }
                            }
                        }
                    }
                }
            }
            if (dt != null)
            {
                dt.TableName = "results";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }
            #endregion results
            if (dsData != null)
            {
                dsData.DataSetName = "dsData";
                dsData.AcceptChanges();
            }
        }


        #endregion function master / all

        #region Data Page follow

        public string get_followup(LoadDocModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            DataSet dsData = new DataSet();
            string user_name = param.user_name ?? "";
            string token_doc = param.token_doc ?? "";
            string sub_software = param.sub_software ?? "";

            if (string.IsNullOrEmpty(sub_software))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid sub_software value"));
            }
            // กำหนด whitelist ของ software ที่อนุญาต
            var allowedSoftwares = new List<string> { "hazop", "jsea", "whatif", "hra" };

            if (!allowedSoftwares.Contains(sub_software.ToLower()))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid sub_software allowedSoftwares"));
            }

            //if (sub_software == "hazop")
            //{
            //    DataHazopSearchFollowUp(ref dsData, user_name, sub_software);
            //}
            //else if (sub_software == "whatif" || sub_software == "hra")
            //{
            DataSearchFollowUp(ref dsData, user_name, sub_software);
            //}
            //else
            //{
            //    dsData.Tables.Add(new DataTable());
            //}  
            if (dsData == null)
            {
                dsData.Tables.Add(new DataTable());
            }

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);
            return json;
        }

        public string get_followup_detail(LoadDocFollowModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            try
            {
                DataSet dsData = new DataSet();
                string user_name = (param.user_name?.ToString() ?? "");
                string token_doc = (param.token_doc?.ToString() ?? "");
                string sub_software = (param.sub_software?.ToString() ?? "");
                string pha_no = (param.pha_no?.ToString() ?? "");
                string responder_user_name = (param.responder_user_name?.ToString() ?? "");
                string pha_seq = token_doc?.ToString() ?? "";

                DataSearchFollowUpDetail(ref dsData, user_name, pha_seq, pha_no, responder_user_name, sub_software);

                get_master_ram(ref dsData);
                if (dsData != null)
                {
                    string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);
                    return json;
                }
                else { return "No Data Page follow."; }
            }
            catch (Exception ex)
            {
                return ex.Message.ToString() + "--> last sql query :" + sqlstr;
            }
        }

        #endregion Data Page follow

        #region Data Page Worksheet

        public string get_details(LoadDocModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name) || string.IsNullOrEmpty(param.sub_software))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            string json = "";
            DataSet dsData = new DataSet();
            string user_name = param.user_name?.Trim() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string token_doc = param.token_doc?.Trim() ?? "";
            string sub_software = param.sub_software?.Trim() ?? "";
            string type_doc = param.type_doc?.Trim() ?? ""; // review_document
            string seq = token_doc;
            string document_module = param.document_module?.Trim() ?? "";
            string msg = "";

            // ตรวจสอบค่า sub_software ว่ามีอยู่ใน whitelist
            var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };
            if (!allowedSubSoftware.Contains(sub_software.ToLower()))
            {
                return "Invalid sub_software.";
            }

            // ตรวจสอบค่า sub_software ว่ามีอยู่ใน whitelist
            var allowedType_doc = new HashSet<string> { "create", "edit", "review_document", "preview", "search", "update" };
            if (!allowedType_doc.Contains(type_doc.ToLower()))
            {
                return "Invalid type_doc.";
            }
            else
            {
                if (type_doc == "create") { seq = "-1"; }
            }


            try
            {
                // เรียกข้อมูล master โดยส่งค่า sub_software เป็น parameter
                get_master(ref dsData, sub_software.ToLower(), "");

                // เรียกข้อมูล history doc โดยส่งค่า sub_software เป็น parameter
                get_history_doc(ref dsData, sub_software);

                // ดึงข้อมูลการไหลของเอกสาร (data flow) โดยใช้ค่า user_name, seq, sub_software, และ document_module
                DataFlow(ref dsData, user_name, seq, sub_software, document_module);

                // ตรวจสอบว่าเป็นการตรวจสอบเอกสาร (review_document)
                if (type_doc == "review_document")
                {
                    #region review_document
                    if (dsData?.Tables["session"]?.Rows.Count > 0)
                    {
                        string action_type = dsData?.Tables["session"]?.Rows[0]["action_type"]?.ToString() ?? "";
                        if (action_type != "insert")
                        {
                            int icount_session = dsData?.Tables["session"]?.Rows.Count - 1 ?? 0;
                            string id_session = dsData?.Tables["session"]?.Rows[icount_session]["id"]?.ToString() ?? "";

                            if (dsData?.Tables["memberteam"] != null && dsData?.Tables["memberteam"]?.Rows.Count > 0)
                            {
                                var drTeam = dsData.Tables["memberteam"]?.AsEnumerable()
                                              .Where(row => row.Field<int>("id_session").ToString() == id_session &&
                                                            string.Equals(row.Field<string>("user_name"), user_name, StringComparison.OrdinalIgnoreCase) &&
                                                            row.Field<int>("action_review") == 0)
                                              .ToArray();

                                if (drTeam?.Length > 0)
                                {
                                    ClassHazopSet cls_set = new ClassHazopSet();
                                    cls_set.set_member_review(user_name, role_type, token_doc, sub_software);
                                }
                            }
                        }
                    }
                    #endregion review_document
                }

                // แปลง DataSet เป็น JSON
                json = JsonConvert.SerializeObject(dsData, Formatting.Indented);
            }
            catch (Exception ex)
            {
                json = $"{ex.Message}-->**msg**{msg}-->**sqlstr**{sqlstr}";
            }

            return json;
        }

        public static void WriteLog(string message)
        {
            //try
            //{
            //    // สร้างเส้นทางไฟล์อัตโนมัติในโฟลเดอร์ Log ภายใต้โฟลเดอร์ Documents ของผู้ใช้
            //    string logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Logs");
            //    Directory.CreateDirectory(logDirectory); // สร้างไดเรกทอรีถ้าไม่มีอยู่

            //    // กำหนดชื่อไฟล์ log เป็น "log.txt"
            //    string logFilePath = Path.Combine(logDirectory, "log.txt");

            //    File.AppendAllText(logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            //}
            //catch (Exception ex)
            //{
            //    // Handle exceptions (optional)
            //    //Console.WriteLine($"Failed to write log: {ex.Message}");
            //}
        }
        public void DataFlow(ref DataSet dsData, string user_name, string seq, string sub_software, string document_module)
        {
            try
            {
                // ตรวจสอบค่า user_name และ sub_software เพื่อป้องกันปัญหา Dereference หลังจาก null check
                if (dsData == null || string.IsNullOrEmpty(user_name) || string.IsNullOrEmpty(sub_software))
                {
                    return;
                }


                DataTable dt = new DataTable();
                List<SqlParameter> parameters = new List<SqlParameter>();
                DataTable dtma = new DataTable();
                string pha_sub_software = sub_software?.ToString() ?? "";
                string pha_no = "";
                int id_pha = 0;

                // Define a whitelist of allowed sub_software values
                var allowedSubSoftware = new HashSet<string> { "jsea", "whatif", "hra", "hazop" };

                // Check if sub_software is valid
                if (!allowedSubSoftware.Contains(pha_sub_software.ToLower()))
                {
                    return;
                }

                if (document_module == "") { document_module = pha_sub_software; }

                string year_now = System.DateTime.Now.Year.ToString();
                if (Convert.ToInt64(year_now) > 2500) { year_now = (Convert.ToInt64(year_now) - 543).ToString(); }

                dt = new DataTable();
                cls = new ClassFunctions();

                parameters = new List<SqlParameter>();
                sqlstr = @" select *  from VW_EPHA_PERSON_DETAILS a where 1=1 and lower(a.user_name) = lower(coalesce(@user_name,'x'))";
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });

                cls_conn = new ClassConnectionDb();
                DataTable dtemp = new DataTable();
                //dtemp = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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
                        dtemp = new DataTable();
                        dtemp = _conn.ExecuteAdapter(command).Tables[0];
                        //dt.TableName = "data";
                        dtemp.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                #region header 
                id_pha = (get_max("epha_t_header", ""));

                parameters = new List<SqlParameter>();

                sqlstr = @" select case when a.year = year(getdate()) then vw.user_name else a.request_user_name end request_user_name
                ,case when a.year = year(getdate()) then vw.user_name else a.request_user_displayname end request_user_displayname
                ,null as approver_user_img
                ,a.*,b.name as pha_status_name, b.descriptions as pha_status_displayname
                ,a.pha_version_text, a.pha_version_desc
                , 'update' as action_type, 0 as action_change, 1 as active_notification
                from epha_t_header a
                left join EPHA_M_STATUS b on a.pha_status = b.id
                left join VW_EPHA_PERSON_DETAILS vw on lower(a.pha_request_by) = lower(vw.user_name)
                where 1=1";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }

                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        pha_no = get_pha_no(sub_software?.ToString() ?? "", year_now);

                        //กรณีที่เป็นใบงานใหม่
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["seq"] = id_pha;
                        dt.Rows[0]["id"] = id_pha;
                        dt.Rows[0]["year"] = year_now;
                        dt.Rows[0]["pha_no"] = pha_no;
                        dt.Rows[0]["pha_version"] = 0;
                        dt.Rows[0]["pha_version_text"] = "-";
                        dt.Rows[0]["pha_version_desc"] = "";

                        dt.Rows[0]["pha_status"] = 11;
                        dt.Rows[0]["pha_sub_software"] = sub_software;
                        dt.Rows[0]["request_approver"] = 0;

                        dt.Rows[0]["pha_status_name"] = "DF";
                        dt.Rows[0]["pha_status_displayname"] = "Draft";
                        if (dtemp?.Rows.Count > 0)
                        {
                            dt.Rows[0]["pha_request_by"] = (dtemp.Rows[0]["user_name"] + "");
                            dt.Rows[0]["request_user_name"] = (dtemp.Rows[0]["user_name"] + "");
                            dt.Rows[0]["request_user_displayname"] = (dtemp.Rows[0]["user_displayname"] + "");
                        }
                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;

                        dt.Rows[0]["active_notification"] = 1;
                        dt.AcceptChanges();
                    }
                }
                if (dt != null)
                {
                    dt.TableName = "header";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }

                if (dt?.Rows.Count > 0)
                {
                    pha_no = (dt.Rows[0]["pha_no"]?.ToString() ?? "");
                    id_pha = Convert.ToInt32(dt.Rows[0]["id"] ?? "");
                }

                set_max_id(ref dtma, "header", (id_pha + 1).ToString());

                #endregion header

                #region general 
                parameters = new List<SqlParameter>();
                sqlstr = @" select b.* 
                , isnull(fa.functional_location,'') as functional_location_audition
                , isnull(fa.functional_location,'') as tagid_audition_def
                , '' as business_unit_name, '' as unit_no_name
                , 'update' as action_type, 0 as action_change
                from epha_t_header a inner join EPHA_T_GENERAL b on a.id  = b.id_pha
                left join EPHA_T_FUNCTIONAL_AUDITION fa on b.id_pha = fa.id_pha 
                where  1=1  ";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                sqlstr += $" order by a.seq  ";


                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                WriteLog("general:" + sqlstr);

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {

                        //กรณีที่เป็นใบงานใหม่
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["seq"] = id_pha;
                        dt.Rows[0]["id"] = id_pha;// (get_max("EPHA_T_GENERAL")); ข้อมูล 1 ต่อ 1 ให้ใช้กับ header ได้เลย
                        dt.Rows[0]["id_pha"] = id_pha;

                        if (!(sub_software == "hra"))
                        {
                            dt.Rows[0]["functional_location_audition"] = "";

                            string selectRam = "5";
                            if (sub_software.ToLower() == "hazop" || sub_software.ToLower() == "whatif")
                            {
                                selectRam = "5";
                            }
                            //default values ram
                            DataTable dtram = dsData.Tables["ram"].Copy(); dtram.AcceptChanges();
                            DataRow[] drRam = dtram.Select("id=" + selectRam);
                            if (drRam.Length > 0)
                            {
                                dt.Rows[0]["id_ram"] = drRam[0]["id"];
                            }

                            dt.Rows[0]["expense_type"] = "OPEX";//OPEX or CAPEX
                            dt.Rows[0]["sub_expense_type"] = "Normal";

                        }
                        else if (sub_software == "hra")
                        {

                            dt.Rows[0]["expense_type"] = "MOC";//MOC or 5YEAR
                            dt.Rows[0]["sub_expense_type"] = "Normal";

                        }

                        dt.Rows[0]["types_of_hazard"] = 1;

                        if (sub_software.ToLower() == "jsea")
                        {
                            if (dsData.Tables["mandatory_note"]?.Rows.Count > 0)
                            {
                                DataRow[] rows = dsData.Tables["mandatory_note"].Select("active_def = 1");
                                if (rows.Length > 0)
                                {
                                    dt.Rows[0]["mandatory_note"] = rows[0]["name"];
                                }
                            }
                        }

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;

                        dt.AcceptChanges();
                    }
                }
                else
                {
                    if (dt?.Rows.Count > 0)
                    {
                        dt.Rows[0]["tagid_audition"] = dt.Rows[0]["tagid_audition_def"];

                        try
                        {
                            if (dsData.Tables["business_unit"]?.Rows.Count > 0)
                            {
                                string id_unit_no = dt.Rows[0]["id_unit_no"]?.ToString() ?? "";
                                string ref_unit_no_name = "";

                                if (id_unit_no != "")
                                {
                                    //DataRow[] dr_unit = dsData.Tables["business_unit"].Select("id=" + id_unit_no); 
                                    DataTable dtUnit = new DataTable();
                                    dtUnit = dsData.Tables["business_unit"].Copy(); dtUnit.AcceptChanges();

                                    var filterParameters = new Dictionary<string, object>();
                                    filterParameters.Add("id", id_unit_no);
                                    var (dr_unit, iMerge) = FilterDataTable(dtUnit, filterParameters);
                                    if (dr_unit != null)
                                    {
                                        if (dr_unit?.Length > 0)
                                        {
                                            ref_unit_no_name = dr_unit[0]["name"]?.ToString() ?? "";
                                        }
                                    }
                                }
                                dt.Rows[0]["unit_no_name"] = ref_unit_no_name;
                            }
                        }
                        catch (Exception ex_unit_no_name) { dt.Rows[0]["unit_no_name"] = ex_unit_no_name.Message.ToString(); }
                        dt.AcceptChanges();
                    }
                }

                if (dt != null)
                {
                    dt.TableName = "general";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }

                #endregion general

                #region general - Department, Sectioin - hra
                if (sub_software.ToLower() == "hra")
                {
                    sqlstr = @"  select emp.user_name, emp.functions, emp.departments, emp.sections 
                     from vw_epha_person_details emp
                     where lower(emp.user_name) = lower(@user_name)";

                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });

                    cls_conn = new ClassConnectionDb();
                    dt = new DataTable();
                    //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                    if (dt != null)
                    {
                        dt.TableName = "org_originator";
                        dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                    }

                    sqlstr = @" select emp.user_name, emp.user_displayname, 'assets/img/team/avatar.webp' as user_img 
                     from vw_epha_person_details emp
                     inner join ( select distinct emp.user_name, emp.functions, emp.departments, emp.sections from vw_epha_person_details emp) emp2 
                     on emp.functions = emp2.functions and emp.departments = emp2.departments and emp.sections = emp2.sections 
                     where emp.main_head_sect = 1 or emp.main_head_vp = 1 or emp.main_head_evp = 1
                     and lower(emp.user_name) = lower(@user_name) order by emp.functions, emp.departments, emp.sections";

                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });

                    cls_conn = new ClassConnectionDb();
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
                    if (dt?.Rows.Count == 0) { dt.Rows.Add(dt.NewRow()); }
                    if (dt != null)
                    {
                        dt.TableName = "org_section_head";
                        dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                    }

                }
                #endregion general - Department, Sectioin - hra 

                #region functional_audition / tagid_audition
                if (sub_software == "hazop" || sub_software == "whatif" || sub_software == "jsea")
                {
                    int id_functional_audition = (get_max("EPHA_T_FUNCTIONAL_AUDITION", seq ?? "-1"));
                    parameters = new List<SqlParameter>();
                    sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                                from epha_t_header a inner join EPHA_T_FUNCTIONAL_AUDITION b on a.id  = b.id_pha
                                where 1=1 ";


                    if (!string.IsNullOrEmpty(seq))
                    {
                        sqlstr += $" and a.seq = @seq ";
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                    }
                    if (!string.IsNullOrEmpty(pha_sub_software))
                    {
                        sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                        parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                    }

                    sqlstr += $"  order by a.seq,b.seq ";

                    cls_conn = new ClassConnectionDb();
                    dt = new DataTable();
                    //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                    if (dt == null || dt?.Rows.Count == 0)
                    {
                        if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                        {
                            dt = new DataTable();
                        }
                        else
                        {
                            //กรณีที่เป็นใบงานใหม่
                            dt.Rows.Add(dt.NewRow());
                            dt.Rows[0]["seq"] = id_functional_audition;
                            dt.Rows[0]["id"] = id_functional_audition;
                            dt.Rows[0]["id_pha"] = id_pha;
                            dt.Rows[0]["create_by"] = user_name;
                            dt.Rows[0]["action_type"] = "insert";
                            dt.Rows[0]["action_change"] = 0;
                            dt.AcceptChanges();
                        }

                    }
                    if (dt != null)
                    {
                        dt.TableName = "functional_audition";
                        dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

                        dt.TableName = "tagid_audition";
                        dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                    }

                    set_max_id(ref dtma, "functional_audition", (id_functional_audition + 1).ToString());
                    set_max_id(ref dtma, "tagid_audition", (id_functional_audition + 1).ToString());

                }
                #endregion functional_audition

                #region session
                int id_session = (get_max("EPHA_T_SESSION", seq ?? "-1"));

                parameters = new List<SqlParameter>();
                sqlstr = @" select b.* , 0 as no, 'update' as action_type, 0 as action_change
                , isnull(format(b.date_to_approve_moc,'dd-MMM-yyyy'),'') as date_to_approve_moc_text
                , isnull(format(b.date_approve_moc,'dd-MMM-yyyy'),'') as date_approve_moc_text
                , RIGHT('0' + CAST(ISNULL(DATEPART(HOUR, b.meeting_start_time), '00') AS varchar), 2) as meeting_start_time_hh
                , RIGHT('0' + CAST(ISNULL(DATEPART(MINUTE, b.meeting_start_time), '00') AS varchar), 2) as meeting_start_time_mm
                , RIGHT('0' + CAST(ISNULL(DATEPART(HOUR, b.meeting_end_time), '00') AS varchar), 2) as meeting_end_time_hh
                , RIGHT('0' + CAST(ISNULL(DATEPART(MINUTE, b.meeting_end_time), '00') AS varchar), 2) as meeting_end_time_mm
                , case when meeting_date is null then 0 else 1 end action_new_row
                from epha_t_header a inner join EPHA_T_SESSION b on a.id  = b.id_pha
                where 1=1  ";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                sqlstr += $" order by a.seq  ";

                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        //กรณีที่เป็นใบงานใหม่
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["seq"] = id_session;
                        dt.Rows[0]["id"] = id_session;
                        dt.Rows[0]["id_pha"] = id_pha;

                        dt.Rows[0]["no"] = 1;

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;

                        dt.Rows[0]["action_new_row"] = 0;

                        dt.AcceptChanges();
                    }

                }
                if (dt != null)
                {
                    dt.TableName = "session";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
                set_max_id(ref dtma, "session", (id_session + 1).ToString());

                //Get Last Session 
                parameters = new List<SqlParameter>();
                sqlstr = @" select distinct  a.seq as id_pha, s2.id_session
                from epha_t_header a
                 inner join EPHA_T_GENERAL g on a.id = g.id_pha
                 inner join EPHA_T_SESSION s on a.id = s.id_pha
                 inner join(select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s2 on a.id = s2.id_pha and s.id = s2.id_session
                 inner join EPHA_T_APPROVER ta2 on a.id = ta2.id_pha and s2.id_session = ta2.id_session
                  where 1=1  ";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                sqlstr += $" order by a.seq ";

                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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
                if (dt != null)
                {
                    dt.TableName = "session_last";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }

                //Get Last Session Reject 
                parameters = new List<SqlParameter>();
                sqlstr = @"  select max(b.id ) as id_pha, max(b.id_session) as id_session 
				             from  epha_t_header a 
				             inner join EPHA_T_APPROVER b on a.id = b.id_pha
				             inner join (select distinct seq,id,pha_no from  epha_t_header ) c on  a.pha_no = c.pha_no 
				             where b.action_review = 2 and b.action_status = 'reject'   ";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and c.seq = @seq  and a.seq <> @seq_ref ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                    parameters.Add(new SqlParameter("@seq_ref", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }

                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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
                if (dt != null)
                {
                    dt.TableName = "session_last_reject";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
                #endregion session

                #region memberteam 
                int id_memberteam = (get_max("EPHA_T_MEMBER_TEAM", seq));

                parameters = new List<SqlParameter>();
                sqlstr = @" select c.* , 'assets/img/team/avatar.webp' as user_img, 'update' as action_type, 0 as action_change
                from epha_t_header a 
                inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                inner join EPHA_T_MEMBER_TEAM c on a.id  = c.id_pha and b.id  = c.id_session
                where 1=1 ";
                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                sqlstr += $"  order by a.seq,b.seq,c.seq ";


                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        //กรณีที่เป็นใบงานใหม่
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["seq"] = id_memberteam;
                        dt.Rows[0]["id"] = id_memberteam;
                        dt.Rows[0]["id_pha"] = id_pha;
                        dt.Rows[0]["id_session"] = id_session;
                        dt.Rows[0]["no"] = 1;

                        dt.Rows[0]["no"] = 1;
                        dt.Rows[0]["user_img"] = "assets/img/team/avatar.webp";

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                    }

                }
                if (dt != null)
                {
                    dt.TableName = "memberteam";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
                set_max_id(ref dtma, "memberteam", (id_memberteam + 1).ToString());
                #endregion memberteam

                #region approver 
                int id_approver = (get_max("EPHA_T_APPROVER", seq));

                parameters = new List<SqlParameter>();
                sqlstr = @" select c.* , 'assets/img/team/avatar.webp' as user_img, 'update' as action_type, 0 as action_change
                 , format(c.date_review,'dd MMM yyyy') as date_review_show
                from epha_t_header a 
                inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                inner join EPHA_T_APPROVER c on a.id  = c.id_pha and b.id  = c.id_session
                 where 1=1 ";
                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                sqlstr += $"  order by a.seq,b.seq,c.seq ";


                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        //กรณีที่เป็นใบงานใหม่
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["seq"] = id_approver;
                        dt.Rows[0]["id"] = id_approver;
                        dt.Rows[0]["id_pha"] = id_pha;
                        dt.Rows[0]["id_session"] = id_session;
                        dt.Rows[0]["no"] = 1;

                        dt.Rows[0]["no"] = 1;
                        dt.Rows[0]["user_img"] = "assets/img/team/avatar.webp";

                        if (sub_software == "hazop" || sub_software == "whatif")
                        {
                            dt.Rows[0]["approver_type"] = "approver";
                        }
                        else if (sub_software == "jsea")
                        {
                            //approver or safety
                            dt.Rows[0]["approver_type"] = "safety";
                        }
                        else if (sub_software == "hra")
                        {
                            //approver or section_head
                            dt.Rows[0]["approver_type"] = "approver";
                        }


                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                    }
                }
                if (dt != null)
                {
                    dt.TableName = "approver";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
                set_max_id(ref dtma, "approver", (id_approver + 1).ToString());
                #endregion approver

                #region approver ta3
                int id_approver_ta3 = (get_max("EPHA_T_APPROVER_TA3", seq));

                parameters = new List<SqlParameter>();
                sqlstr = @" select d.* , 'assets/img/team/avatar.webp' as user_img, 'update' as action_type, 0 as action_change
                 , format(c.date_review,'dd MMM yyyy') as date_review_show
                from epha_t_header a 
                inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                inner join EPHA_T_APPROVER c on a.id  = c.id_pha and b.id  = c.id_session 
                inner join EPHA_T_APPROVER_TA3 d on a.id  = d.id_pha and c.id  = d.id_approver
                where 1=1 ";
                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                sqlstr += $"  order by a.seq,b.seq,c.seq ,d.seq ";


                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {

                        //กรณีที่เป็นใบงานใหม่
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["seq"] = id_approver_ta3;
                        dt.Rows[0]["id"] = id_approver_ta3;
                        dt.Rows[0]["id_pha"] = id_pha;
                        dt.Rows[0]["id_session"] = id_session;
                        dt.Rows[0]["id_approver"] = id_approver;
                        dt.Rows[0]["no"] = 1;

                        dt.Rows[0]["no"] = 1;
                        dt.Rows[0]["user_img"] = "assets/img/team/avatar.webp";

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                    }

                }
                if (dt != null)
                {
                    dt.TableName = "approver_ta3";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
                set_max_id(ref dtma, "approver_ta3", (id_approver_ta3 + 1).ToString());

                #endregion approver ta3

                #region relatedpeople   
                if (sub_software.ToLower() == "hazop" || sub_software.ToLower() == "whatif"
                     || sub_software.ToLower() == "jsea" || sub_software.ToLower() == "hra")
                {
                    int id_relatedpeople = (get_max("EPHA_T_RELATEDPEOPLE", seq));

                    parameters = new List<SqlParameter>();
                    sqlstr = @" select c.* , 'assets/img/team/avatar.webp' as user_img, 'update' as action_type, 0 as action_change
                 , format(c.date_review,'dd MMM yyyy') as date_review_show
                from epha_t_header a 
                inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                inner join EPHA_T_RELATEDPEOPLE c on a.id  = c.id_pha and b.id  = c.id_session
                where 1=1 ";
                    if (!string.IsNullOrEmpty(seq))
                    {
                        sqlstr += $" and a.seq = @seq ";
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                    }
                    if (!string.IsNullOrEmpty(pha_sub_software))
                    {
                        sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                        parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                    }

                    sqlstr += $"  order by a.seq,b.seq,c.seq ";


                    cls_conn = new ClassConnectionDb();
                    dt = new DataTable();
                    //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                    if (dt == null || dt?.Rows.Count == 0)
                    {
                        if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                        {
                            dt = new DataTable();
                        }
                        else { }


                        if (true)
                        {
                            //attendees,specialist,reviewer,approver
                            string[] xsplit = ("attendees,specialist").Split(',');//reviewer,approver แยกไปอยู่ในส่วนข้อมูลหลัก
                            if (sub_software.ToLower() == "jsea")
                            {
                                //reviewer safety (main) แยกไปอยู่ในส่วนข้อมูลหลัก
                                //reviewer safety,approver, safety, ae, agsi แยกมาไว้ที่ reviewer 
                                xsplit = ("attendees,specialist,reviewer").Split(',');
                                xsplit = ("specialist").Split(',');
                            }
                            for (int i = 0; i < xsplit.Length; i++)
                            {
                                string _user_type = xsplit[i].Trim();

                                DataRow[] drUt = dt.Select("user_type ='" + _user_type + "'");
                                if (drUt.Length == 0)
                                {
                                    int irow = dt?.Rows.Count ?? 0;
                                    //กรณีที่เป็นใบงานใหม่ 
                                    dt.Rows.Add(dt.NewRow());
                                    dt.Rows[irow]["seq"] = id_relatedpeople;
                                    dt.Rows[irow]["id"] = id_relatedpeople;
                                    dt.Rows[irow]["id_pha"] = id_pha;
                                    dt.Rows[irow]["id_session"] = id_session;
                                    dt.Rows[irow]["no"] = 1;

                                    dt.Rows[irow]["user_img"] = "assets/img/team/avatar.webp";
                                    dt.Rows[irow]["user_type"] = _user_type;//attendees,specialist 
                                    dt.Rows[irow]["approver_type"] = (_user_type == "reviewer" ? "safety" : "member");//member, free_text, safety, ae, agsi  

                                    dt.Rows[irow]["create_by"] = user_name;
                                    dt.Rows[irow]["action_type"] = "insert";
                                    dt.Rows[irow]["action_change"] = 0;
                                    dt.AcceptChanges();
                                    id_relatedpeople += 1;
                                }
                            }
                        }
                    }
                    if (dt != null)
                    {
                        dt.TableName = "relatedpeople";
                        dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                    }
                    set_max_id(ref dtma, "relatedpeople", (id_relatedpeople + 1).ToString());


                    #region relatedpeople outsider  
                    int id_relatedpeople_outsider = (get_max("EPHA_T_RELATEDPEOPLE_OUTSIDER", seq));

                    parameters = new List<SqlParameter>();
                    sqlstr = @" select c.* , 'assets/img/team/avatar.webp' as user_img, 'update' as action_type, 0 as action_change
                    , format(c.date_review,'dd MMM yyyy') as date_review_show
                    from epha_t_header a 
                    inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                    inner join EPHA_T_RELATEDPEOPLE_OUTSIDER c on a.id  = c.id_pha and b.id  = c.id_session
                   where 1=1 ";
                    if (!string.IsNullOrEmpty(seq))
                    {
                        sqlstr += $" and a.seq = @seq ";
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                    }
                    if (!string.IsNullOrEmpty(pha_sub_software))
                    {
                        sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                        parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                    }

                    sqlstr += $"  order by a.seq,b.seq,c.seq ";


                    cls_conn = new ClassConnectionDb();
                    dt = new DataTable();
                    //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                    if (true)
                    {
                        //attendees,specialist,reviewer,approver,member
                        string[] xsplit = ("attendees,specialist,member").Split(',');//reviewer,approver แยกไปอยู่ในส่วนข้อมูลหลัก
                        if (sub_software.ToLower() == "jsea")
                        {
                            //reviewer safety (main) แยกไปอยู่ในส่วนข้อมูลหลัก
                            //reviewer safety,approver, safety, ae, agsi แยกมาไว้ที่ reviewer 
                            xsplit = ("attendees,specialist,reviewer").Split(',');
                            xsplit = ("member,specialist,reviewer").Split(',');
                        }
                        int icount_session = dsData?.Tables["session"]?.Rows.Count - 1 ?? 0;
                        string _id_session = dsData?.Tables["session"].Rows[icount_session]["seq"].ToString() ?? "0";

                        for (int i = 0; i < xsplit.Length; i++)
                        {
                            string _user_type = xsplit[i].Trim();
                            DataRow[] drUt = dt.Select("user_type ='" + _user_type + "'");
                            if (dsData?.Tables["relatedpeople"] != null)
                            {
                                drUt = dsData.Tables["relatedpeople"].Select("user_type ='" + _user_type + "'");
                            }
                            if (drUt.Length == 0)
                            {
                                int irow = dt?.Rows.Count ?? 0;

                                //กรณีที่เป็นใบงานใหม่ 
                                dt.Rows.Add(dt.NewRow());
                                dt.Rows[irow]["seq"] = id_relatedpeople_outsider;
                                dt.Rows[irow]["id"] = id_relatedpeople_outsider;
                                dt.Rows[irow]["id_pha"] = id_pha;
                                dt.Rows[irow]["id_session"] = _id_session;
                                dt.Rows[irow]["no"] = 1;

                                dt.Rows[irow]["user_img"] = "assets/img/team/avatar.webp";
                                dt.Rows[irow]["user_type"] = _user_type;//attendees,specialist 
                                dt.Rows[irow]["approver_type"] = "free_text";//member, free_text, safety, ae, agsi  

                                dt.Rows[irow]["create_by"] = user_name;
                                dt.Rows[irow]["action_type"] = "insert";
                                dt.Rows[irow]["action_change"] = 0;
                                dt.AcceptChanges();
                                id_relatedpeople_outsider += 1;
                            }

                        }
                    }

                    if (dt != null)
                    {
                        dt.TableName = "relatedpeople_outsider";
                        dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                    }

                    set_max_id(ref dtma, "relatedpeople_outsider", (id_relatedpeople_outsider + 1).ToString());

                    #endregion relatedpeople outsider

                }
                #endregion relatedpeople

                #region drawing 
                int id_drawing = (get_max("EPHA_T_DRAWING", seq));

                parameters = new List<SqlParameter>();
                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                from epha_t_header a inner join EPHA_T_DRAWING b on a.id  = b.id_pha
                where 1=1 ";
                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                if (!string.IsNullOrEmpty(document_module))
                {
                    sqlstr += " and lower(b.document_module) = lower(@document_module)";
                    parameters.Add(new SqlParameter("@document_module", SqlDbType.VarChar, 50) { Value = document_module });
                }
                sqlstr += " order by a.seq,b.seq";

                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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


                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        //กรณีที่เป็นใบงานใหม่
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["seq"] = id_drawing;
                        dt.Rows[0]["id"] = id_drawing;
                        dt.Rows[0]["id_pha"] = id_pha;

                        dt.Rows[0]["no"] = 1;

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                    }

                }
                if (dt != null)
                {
                    dt.TableName = "drawing";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
                set_max_id(ref dtma, "drawing", (id_drawing + 1).ToString());

                #endregion drawing

                #region worksheet drawing responder & reviewer
                string pha_seq = id_pha.ToString();
                int id_drawing_worksheet = (get_max("EPHA_T_DRAWING_WORKSHEET", pha_seq));

                parameters = new List<SqlParameter>();
                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                from epha_t_header a inner join EPHA_T_DRAWING_WORKSHEET b on a.id  = b.id_pha
                where lower(document_module) = lower('followup') ";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                sqlstr += $"  order by a.seq,b.seq ";


                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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
                if (dt != null)
                {
                    dt.TableName = "drawingworksheet_responder";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }

                parameters = new List<SqlParameter>();
                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                from epha_t_header a inner join EPHA_T_DRAWING_WORKSHEET b on a.id  = b.id_pha
                where lower(document_module) = lower('review_followup') ";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                sqlstr += $"  order by a.seq,b.seq ";


                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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
                if (dt != null)
                {
                    dt.TableName = "drawingworksheet_reviewer";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }

                set_max_id(ref dtma, "drawingworksheet", (id_drawing_worksheet + 1).ToString());
                #endregion worksheet drawing responder & reviewer

                #region Approver Drawing
                int id_drawing_approver = (get_max("EPHA_T_DRAWING_APPROVER", pha_seq));

                parameters = new List<SqlParameter>();
                sqlstr = @" select da.* , 'update' as action_type, 0 as action_change 
                from epha_t_header a  
                inner join EPHA_T_GENERAL g on a.id = g.id_pha   
                inner join EPHA_T_SESSION s on a.id = s.id_pha 
                inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s2 on a.id = s2.id_pha and s.id = s2.id_session 
                inner join EPHA_T_APPROVER ta2 on a.id = ta2.id_pha and s2.id_session = ta2.id_session 
                inner join EPHA_T_DRAWING_APPROVER da on a.id = da.id_pha and ta2.id_pha = da.id_pha and ta2.id_session = da.id_session and ta2.seq = da.id_approver
                where 1=1 ";
                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += $" and a.seq = @seq ";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                if (!string.IsNullOrEmpty(pha_sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                    parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                }

                if (!string.IsNullOrEmpty(document_module))
                {
                    sqlstr += " and lower(da.document_module) = lower(@document_module)";
                    parameters.Add(new SqlParameter("@document_module", SqlDbType.VarChar, 50) { Value = document_module });
                }
                sqlstr += " order by a.seq,da.seq";

                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        if (dsData.Tables["session_last"]?.Rows.Count > 0)
                        {
                            if (dsData.Tables["approver"]?.Rows.Count > 0)
                            {
                                DataTable dtApprover = dsData.Tables["approver"]; dtApprover.AcceptChanges();
                                DataTable dtSessionLast = dsData.Tables["session_last"]; dtApprover.AcceptChanges();
                                string _id_session = (dtSessionLast.Rows[0]["id_session"] + "").ToString();
                                int irow = dt?.Rows.Count ?? 0;
                                for (int i = 0; i < dtApprover?.Rows.Count; i++)
                                {
                                    string _id_approver = (dtApprover.Rows[i]["seq"] + "").ToString();
                                    DataRow[] dr = dt.Select("id_approver = " + _id_approver);
                                    if (dr.Length == 0)
                                    {
                                        //กรณีที่เป็นใบงานใหม่
                                        dt.Rows.Add(dt.NewRow());
                                        dt.Rows[irow]["seq"] = id_drawing_approver;
                                        dt.Rows[irow]["id"] = id_drawing_approver;
                                        dt.Rows[irow]["id_pha"] = id_pha;
                                        dt.Rows[irow]["id_session"] = _id_session;
                                        dt.Rows[irow]["id_approver"] = _id_approver;

                                        dt.Rows[irow]["no"] = 1;

                                        dt.Rows[irow]["document_module"] = document_module;
                                        dt.Rows[irow]["create_by"] = user_name;
                                        dt.Rows[irow]["action_type"] = "insert";
                                        dt.Rows[irow]["action_change"] = 0;
                                        dt.AcceptChanges();
                                        irow++;

                                        id_drawing_approver++;
                                    }
                                }
                            }
                            else
                            {
                                if (dt?.Rows.Count == 0)
                                {
                                    //กรณีที่เป็นใบงานใหม่
                                    dt.Rows.Add(dt.NewRow());
                                    dt.Rows[0]["seq"] = id_drawing_approver;
                                    dt.Rows[0]["id"] = id_drawing_approver;
                                    dt.Rows[0]["id_pha"] = id_pha;
                                    dt.Rows[0]["id_session"] = id_session;
                                    dt.Rows[0]["id_approver"] = id_approver;

                                    dt.Rows[0]["no"] = 1;

                                    dt.Rows[0]["document_module"] = document_module;
                                    dt.Rows[0]["create_by"] = user_name;
                                    dt.Rows[0]["action_type"] = "insert";
                                    dt.Rows[0]["action_change"] = 0;
                                    dt.AcceptChanges();
                                }
                            }
                        }
                    }
                }
                else
                {
                    //กรณีที่มีข้อมูล drawing ไม่ครบทุกคน
                    if (dsData.Tables["session_last"]?.Rows.Count > 0)
                    {
                        if (dsData.Tables["approver"]?.Rows.Count > 0)
                        {
                            DataTable dtApprover = dsData.Tables["approver"]; dtApprover.AcceptChanges();
                            DataTable dtSessionLast = dsData.Tables["session_last"]; dtApprover.AcceptChanges();
                            string _id_session = (dtSessionLast.Rows[0]["id_session"] + "").ToString();
                            int irow = dt?.Rows.Count ?? 0;
                            for (int i = 0; i < dtApprover?.Rows.Count; i++)
                            {
                                string _id_approver = (dtApprover.Rows[i]["seq"] + "").ToString();
                                DataRow[] dr = dt.Select("id_approver = " + _id_approver);
                                if (dr.Length == 0)
                                {
                                    //กรณีที่เป็นใบงานใหม่
                                    dt.Rows.Add(dt.NewRow());
                                    dt.Rows[irow]["seq"] = id_drawing_approver;
                                    dt.Rows[irow]["id"] = id_drawing_approver;
                                    dt.Rows[irow]["id_pha"] = id_pha;
                                    dt.Rows[irow]["id_session"] = _id_session;
                                    dt.Rows[irow]["id_approver"] = _id_approver;

                                    dt.Rows[irow]["no"] = 1;

                                    dt.Rows[irow]["document_module"] = document_module;
                                    dt.Rows[irow]["create_by"] = user_name;
                                    dt.Rows[irow]["action_type"] = "insert";
                                    dt.Rows[irow]["action_change"] = 0;
                                    dt.AcceptChanges();
                                    irow++;

                                    id_drawing_approver++;
                                }
                            }
                        }
                    }

                }
                if (dt != null)
                {
                    dt.TableName = "drawing_approver";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }

                set_max_id(ref dtma, "drawing_approver", (id_drawing_approver + 1).ToString());

                #endregion Approver Drawing

                #region worksheet
                if (dsData != null)
                {
                    if (sub_software.ToLower() == "hazop")
                    {
                        _hazop_data(ref dsData, user_name, seq, id_pha, ref dtma);
                    }
                    else if (sub_software.ToLower() == "jsea")
                    {
                        _jsea_data(ref dsData, user_name, seq, id_pha, ref dtma);
                    }
                    else if (sub_software.ToLower() == "whatif")
                    {
                        _whatif_data(ref dsData, user_name, seq, id_pha, ref dtma);
                    }
                    else if (sub_software.ToLower() == "hra")
                    {
                        _hra_data(ref dsData, user_name, seq, id_pha, ref dtma);
                    }
                }
                #endregion worksheet
                if (dt != null)
                {
                    dtma.TableName = "max";
                    dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();
                }

                #region check user in pha_no
                string role_type = "";
                try
                {
                    check_role_user_active(user_name ?? "", ref role_type);
                    if (!string.IsNullOrEmpty(role_type))
                    {
                        parameters = new List<SqlParameter>();
                        sqlstr = @" select a.pha_no
                                from epha_t_header a
                                inner join EPHA_T_GENERAL b on a.id  = b.id_pha
                                left join EPHA_M_STATUS ms on a.pha_status = ms.id
                                left join VW_EPHA_PERSON_DETAILS vw on lower(a.pha_request_by) = lower(vw.user_name)                 
                                inner join VW_EPHA_DATA_DOC_BY_USER du on a.id = du.id_pha  
                                where 1=1";
                        if (!string.IsNullOrEmpty(seq))
                        {
                            sqlstr += " and a.seq = @seq";
                            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                        }
                        if (!string.IsNullOrEmpty(pha_sub_software))
                        {
                            sqlstr += " and lower(a.pha_sub_software) = lower(@pha_sub_software)";
                            parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software ?? "" });
                        }
                        if (role_type != "admin")
                        {
                            if (!string.IsNullOrEmpty(user_name))
                            {
                                sqlstr += @" and lower(du.user_name)  = lower(@user_name)";
                                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
                            }
                        }
                        sqlstr += " order by a.pha_no";

                        WriteLog($"check user in pha_no: {sqlstr}");
                        WriteLog($"parameters: {parameters[0].Value.ToString()}");

                        DataTable dtcheck = new DataTable();
                        //dtcheck = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters,"user_in_pha_no");
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
                                dtcheck = new DataTable();
                                dtcheck = _conn.ExecuteAdapter(command).Tables[0];
                                dtcheck.TableName = "user_in_pha_no";
                                dtcheck.AcceptChanges();
                            }
                            catch { }
                            finally { _conn.CloseConnection(); }
                        }
                        catch { }
                        #endregion Execute to Datable
                        if (dtcheck != null)
                        {
                            if (dtcheck?.Rows.Count == 0)
                            {
                                dtcheck.Rows.Add(dtcheck.NewRow()); dtcheck.AcceptChanges();
                                dtcheck.Rows[0]["pha_no"] = ""; dtcheck.AcceptChanges();
                            }
                            if (dsData != null)
                            {
                                if (dtcheck != null)
                                {
                                    dsData.Tables.Add(dtcheck.Copy()); dsData.AcceptChanges();
                                }
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    WriteLog($"DataFlow:{ex.ToString()}");
                }
                #endregion check user in pha_no
                if (dsData != null)
                {
                    dsData.DataSetName = "dsData"; dsData.AcceptChanges();
                }
            }
            catch (Exception ex_function) { WriteLog(ex_function.ToString()); }
        }

        private void _recomment_data(string user_name, string seq, int id_pha, ref DataTable dtma, int id_worksheet)
        {
            try
            {
                // ตรวจสอบค่า  เพื่อป้องกันปัญหา Dereference หลังจาก null check
                if (dtma == null || string.IsNullOrEmpty(user_name))
                {
                    return;
                }
                seq = seq ?? "-1";

                int id_recommendations = 0;
                int id_recom_setting = 0;
                int id_recom_follow = 0;

                string recommendations = "";
                DataSet dsData = new DataSet();
                List<SqlParameter> parameters = new List<SqlParameter>();

                #region recommendations 
                sqlstr = @" select r.*, 'update' as action_type, 0 as action_change, b.index_rows 
                    from epha_t_header a     
                    inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha  
                    inner join EPHA_T_RECOMMENDATIONS r on a.id  = r.id_pha and b.id = r.id_worksheet  
                    where  1=1  ";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " and a.seq = @seq";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                sqlstr += "  order by a.seq, b.no, r.no ";

                DataTable dt = new DataTable();
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

                id_recommendations = get_max("EPHA_T_RECOMMENDATIONS", seq);

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["id_pha"] = id_pha;
                        dt.Rows[0]["seq"] = id_recommendations;
                        dt.Rows[0]["id"] = id_recommendations;

                        dt.Rows[0]["id_worksheet"] = id_worksheet;

                        dt.Rows[0]["index_rows"] = 0;
                        dt.Rows[0]["no"] = 1;

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "new";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                    }
                }

                set_max_id(ref dtma, "recommendations", (id_recommendations + 1).ToString());

                if (dt?.Rows.Count > 0) { recommendations = dt.Rows[0]["recommendations"]?.ToString() ?? ""; }

                if (dt != null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt.TableName = "recommendations";
                    dsData.Tables.Add(dt.Copy());
                    dsData.AcceptChanges();
                }

                #endregion recommendations

                #region recom_setting
                sqlstr = @" select rs.*  
                    , 'update' as action_type, 0 as action_change, b.index_rows 
                    from epha_t_header a     
                    inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha  
                    inner join EPHA_T_RECOMMENDATIONS r on a.id  = r.id_pha and b.id = r.id_worksheet  
                    left join EPHA_T_RECOM_SETTING rs on a.id  = rs.id_pha and lower(r.recommendations) = lower(rs.recommendations) 
                    where 1=1 ";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " and a.seq = @seq";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                sqlstr += "  order by a.seq, b.no, r.no ";

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

                id_recom_setting = get_max("EPHA_T_RECOM_SETTING", seq);

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["id_pha"] = id_pha;
                        dt.Rows[0]["seq"] = id_recom_setting;
                        dt.Rows[0]["id"] = id_recom_setting;

                        dt.Rows[0]["recommendations"] = recommendations;
                        dt.Rows[0]["id_rangtype"] = 1;

                        dt.Rows[0]["index_rows"] = 0;
                        dt.Rows[0]["no"] = 1;

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "new";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                    }
                }

                set_max_id(ref dtma, "recom_setting", (id_recom_setting + 1).ToString());

                dt.TableName = "recom_setting";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();

                #endregion recom_setting

                #region recom_follow
                sqlstr = @" select f.*  
                    , 'update' as action_type, 0 as action_change, b.index_rows 
                    from epha_t_header a     
                    inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha  
                    inner join EPHA_T_RECOMMENDATIONS r on a.id  = r.id_pha and b.id = r.id_worksheet  
                    inner join EPHA_T_RECOM_SETTING rs on a.id  = rs.id_pha and lower(r.recommendations) = lower(rs.recommendations) 
                    inner join EPHA_T_RECOM_FOLLOW f on a.id  = f.id_pha and rs.id = f.id
                    where 1=1 ";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " and a.seq = @seq";
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                }
                sqlstr += "   order by a.seq, b.no, r.no, rs.no, f.no  ";


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

                id_recom_follow = get_max("EPHA_T_RECOM_FOLLOW", seq);


                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["id_pha"] = id_pha;
                        dt.Rows[0]["seq"] = id_recom_follow;
                        dt.Rows[0]["id"] = id_recom_follow;

                        dt.Rows[0]["id_recom"] = id_recom_setting;

                        dt.Rows[0]["index_rows"] = 0;
                        dt.Rows[0]["no"] = 1;

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "new";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                    }
                }

                set_max_id(ref dtma, "recom_follow", (id_recom_follow + 1).ToString());

                dt.TableName = "recom_follow";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();

                #endregion recom_follow
            }
            catch (Exception ex_function) { WriteLog(ex_function.ToString()); }
        }
        private void _hazop_data(ref DataSet? dsData, string user_name, string seq, int id_pha, ref DataTable dtma)
        {
            // ตรวจสอบค่า  เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (dsData == null || dtma == null || string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(user_name))
            {
                return;
            }

            seq = seq ?? "-1";

            int id_node = 0;
            int id_nodeworksheet = 0;
            DataTable dt = new DataTable();
            List<SqlParameter> parameters = new List<SqlParameter>();

            #region node 
            sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                from epha_t_header a inner join EPHA_T_NODE b on a.id  = b.id_pha
                where a.seq = @seq
                order by a.seq,b.no";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

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

            id_node = get_max("EPHA_T_NODE", seq);
            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                else
                {
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["seq"] = id_node;
                    dt.Rows[0]["id"] = id_node;
                    dt.Rows[0]["id_pha"] = id_pha;

                    dt.Rows[0]["no"] = 1;

                    dt.Rows[0]["create_by"] = user_name;
                    dt.Rows[0]["action_type"] = "insert";
                    dt.Rows[0]["action_change"] = 0;
                    dt.AcceptChanges();
                }
            }
            set_max_id(ref dtma, "node", (id_node + 1).ToString());

            dt.TableName = "node";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();
            #endregion node

            #region nodedrawing 
            sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                , b.id_node as seq_node
                from epha_t_header a inner join EPHA_T_NODE_DRAWING b on a.id  = b.id_pha
                where a.seq = @seq
                order by a.seq,b.no";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

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

            int id_nodedrawing = get_max("EPHA_T_NODE_DRAWING", seq);
            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                else
                {
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["seq"] = id_nodedrawing;
                    dt.Rows[0]["id"] = id_nodedrawing;
                    dt.Rows[0]["id_node"] = id_node;
                    dt.Rows[0]["id_pha"] = id_pha;

                    dt.Rows[0]["seq_node"] = id_node;

                    dt.Rows[0]["no"] = 1;

                    dt.Rows[0]["create_by"] = user_name;
                    dt.Rows[0]["action_type"] = "insert";
                    dt.Rows[0]["action_change"] = 0;
                    dt.AcceptChanges();
                }
            }
            set_max_id(ref dtma, "nodedrawing", (id_nodedrawing + 1).ToString());

            dt.TableName = "nodedrawing";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();
            #endregion nodedrawing

            #region nodeguidwords 
            sqlstr = @"  select b.* ,coalesce(def_selected,0) as selected_type , 'update' as action_type, 0 as action_change
                , b.id_node as seq_node, g.guide_words as guidewords, g.deviations, g.no_guide_words as guidewords_no, g.no_deviations as deviations_no
                , g.id 
                from epha_t_header a inner join EPHA_T_NODE_GUIDE_WORDS b on a.id  = b.id_pha
                inner join EPHA_M_GUIDE_WORDS g on b.id_guide_word = g.id
                where a.seq = @seq
                order by a.seq, b.no, g.no";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

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

            int id_nodeguidwords = get_max("EPHA_T_NODE_GUIDE_WORDS", seq);
            int no_nodeguidwords = 1;
            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                else
                {
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["seq"] = id_nodeguidwords;
                    dt.Rows[0]["id"] = id_nodeguidwords;
                    dt.Rows[0]["id_node"] = id_node;
                    dt.Rows[0]["id_pha"] = id_pha;

                    dt.Rows[0]["seq_node"] = id_node;
                    dt.Rows[0]["no"] = 1;

                    dt.Rows[0]["create_by"] = user_name;
                    dt.Rows[0]["action_type"] = "insert";
                    dt.Rows[0]["action_change"] = 0;
                    dt.AcceptChanges();
                }
            }
            set_max_id(ref dtma, "nodeguidwords", (id_nodeguidwords + 1).ToString());

            DataTable dtnodeguidwords = dt.Copy();
            dtnodeguidwords.AcceptChanges();

            dt.TableName = "nodeguidwords";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();
            #endregion nodeguidwords

            #region nodeworksheet 
            sqlstr = @" select b.* , 0 as no  
                , 'update' as action_type, 0 as action_change
                , b.id_node as seq_node, g.guide_words as guidewords, g.deviations, g.no_guide_words as guidewords_no, g.no_deviations as deviations_no
                , vw.user_id as responder_user_id, vw.user_email as responder_user_email
                , 'assets/img/team/avatar.webp' as responder_user_img
                , n.no as node_no
                , n.no as node_no, format(b.reviewer_action_date,'dd MMM yyyy') as reviewer_date  
                , b.index_rows
                from epha_t_header a   
                inner join EPHA_T_NODE n on a.id  = n.id_pha 
                inner join EPHA_T_NODE_WORKSHEET b on a.id  = b.id_pha and n.id = b.id_node 
                inner join EPHA_M_GUIDE_WORDS g on b.id_guide_word = g.id    
                left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name) 
                where a.seq = @seq
                order by b.index_rows, n.no, g.id, b.no, b.causes_no, b.consequences_no, b.category_no";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

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

            id_nodeworksheet = get_max("EPHA_T_NODE_WORKSHEET", seq);
            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                else
                {
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["seq"] = id_nodeworksheet;
                    dt.Rows[0]["id"] = id_nodeworksheet;
                    dt.Rows[0]["id_node"] = id_node;
                    dt.Rows[0]["id_pha"] = id_pha;

                    dt.Rows[0]["seq_node"] = id_node;

                    dt.Rows[0]["id_guide_word"] = id_nodeguidwords;
                    dt.Rows[0]["seq_guide_word"] = id_nodeguidwords;
                    dt.Rows[0]["guidewords_no"] = no_nodeguidwords;

                    dt.Rows[0]["index_rows"] = 0;
                    dt.Rows[0]["no"] = 1;

                    dt.Rows[0]["row_type"] = "causes";

                    dt.Rows[0]["seq_causes"] = 1;
                    dt.Rows[0]["seq_consequences"] = 1;
                    dt.Rows[0]["seq_category"] = 1;
                    dt.Rows[0]["seq_recommendations"] = id_nodeworksheet;
                    dt.Rows[0]["fk_recommendations"] = id_nodeworksheet;

                    dt.Rows[0]["causes_no"] = 1;
                    dt.Rows[0]["consequences_no"] = 1;
                    dt.Rows[0]["category_no"] = 1;

                    dt.Rows[0]["recommendations_no"] = 1;
                    dt.Rows[0]["recommendations_action_no"] = 1;

                    dt.Rows[0]["implement"] = 0;

                    dt.Rows[0]["create_by"] = user_name;
                    dt.Rows[0]["action_type"] = "new";
                    dt.Rows[0]["action_change"] = 0;
                    dt.Rows[0]["action_status"] = "Open";

                    dt.AcceptChanges();
                }
            }
            set_index_worksheet(ref dt);

            set_max_id(ref dtma, "nodeworksheet", (id_nodeworksheet + 1).ToString());

            dt.TableName = "nodeworksheet";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();

            _recomment_data(user_name, seq, id_pha, ref dtma, id_nodeworksheet);

            #endregion nodeworksheet
        }
        private void _jsea_data(ref DataSet? dsData, string user_name, string seq, int id_pha, ref DataTable dtma)
        {
            try
            {
                // ตรวจสอบค่า  เพื่อป้องกันปัญหา Dereference หลังจาก null check
                if (dsData == null || dtma == null || string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(user_name))
                {
                    return;
                }
                seq = seq ?? "-1";

                int id_tasks = 0;
                //int id_related = 0; 
                DataTable dt = new DataTable();
                List<SqlParameter> parameters = new List<SqlParameter>();

                #region tasks_worksheet 

                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                , b.index_rows
                from epha_t_header a inner join EPHA_T_TASKS_WORKSHEET b on a.id  = b.id_pha
                where a.seq = @seq
                order by b.index_rows, b.workstep_no, b.taskdesc_no, b.potentailhazard_no, b.possiblecase_no, b.category_no";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });


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

                id_tasks = get_max("EPHA_T_TASKS_WORKSHEET", seq);
                if (dt?.Rows.Count == 0)
                {
                    //กรณีที่เป็นใบงานใหม่
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["seq"] = id_tasks;
                    dt.Rows[0]["id"] = id_tasks;
                    dt.Rows[0]["id_pha"] = id_pha;

                    dt.Rows[0]["index_rows"] = 0;//ใช้ในการค้นหาลำดับ
                    dt.Rows[0]["no"] = 1;
                    dt.Rows[0]["row_type"] = "workstep";//workstep,taskdesc,potentailhazard,possiblecase,cat

                    dt.Rows[0]["seq_workstep"] = 1;
                    dt.Rows[0]["seq_taskdesc"] = 1;
                    dt.Rows[0]["seq_potentailhazard"] = 1;
                    dt.Rows[0]["seq_possiblecase"] = 1;
                    dt.Rows[0]["seq_category"] = 1;
                    dt.Rows[0]["seq_recommendations"] = id_tasks;
                    dt.Rows[0]["fk_recommendations"] = id_tasks;

                    dt.Rows[0]["workstep_no"] = 1;
                    dt.Rows[0]["taskdesc_no"] = 1;
                    dt.Rows[0]["potentailhazard_no"] = 1;
                    dt.Rows[0]["possiblecase_no"] = 1;
                    dt.Rows[0]["category_no"] = 1;

                    dt.Rows[0]["recommendations_no"] = 1;
                    dt.Rows[0]["recommendations_action_no"] = 1;

                    dt.Rows[0]["implement"] = 0;

                    dt.Rows[0]["action_status"] = "Open";

                    dt.Rows[0]["create_by"] = user_name;
                    dt.Rows[0]["action_type"] = "insert";
                    dt.Rows[0]["action_change"] = 0;
                    dt.AcceptChanges();
                }
                set_index_worksheet(ref dt);

                dt.TableName = "tasks_worksheet";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();

                set_max_id(ref dtma, "tasks_worksheet", (id_tasks + 1).ToString());

                #endregion tasks_worksheet

            }
            catch (Exception ex_function) { WriteLog(ex_function.ToString()); }
        }
        private void _hra_data(ref DataSet? dsData, string user_name, string seq, int id_pha, ref DataTable dtma)
        {
            try
            {
                // ตรวจสอบค่า  เพื่อป้องกันปัญหา Dereference หลังจาก null check
                if (dsData == null || dtma == null || string.IsNullOrEmpty(user_name))
                {
                    return;
                }
                seq = seq ?? "-1";

                int id_subareas_def = 0;
                int id_subareas = 0;
                int id_hazard_def = 0;
                int id_hazard = 0;

                int id_tasks_def = 0;
                int id_tasks = 0;
                int id_workers_def = 0;
                int id_workers = 0;
                int id_desc = 0;

                Boolean bDataHazard = false;
                Boolean bDataTasks = false;
                int id_worksheet = 0;

                DataTable dt = new DataTable();
                List<SqlParameter> parameters = new List<SqlParameter>();

                //data group table1
                if (true)
                {
                    #region subareas
                    id_subareas = get_max("EPHA_T_TABLE1_SUBAREAS", seq);

                    sqlstr = @" select b.* , 'update' as action_type, 0 as action_change, 0 as index_rows
                    from epha_t_header a inner join EPHA_T_TABLE1_SUBAREAS b on a.id  = b.id_pha
                    where 1=1 ";

                    if (!string.IsNullOrEmpty(seq))
                    {
                        sqlstr += $@" and a.seq = @seq";
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                    }
                    sqlstr += $@" order by a.seq, b.id_sub_area, b.no";

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


                    if (dt == null || dt?.Rows.Count == 0)
                    {
                        if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                        {
                            dt = new DataTable();
                        }
                        else
                        {
                            dt.Rows.Add(dt.NewRow());
                            dt.Rows[0]["seq"] = id_subareas;
                            dt.Rows[0]["id"] = id_subareas;
                            dt.Rows[0]["id_pha"] = id_pha;

                            dt.Rows[0]["no"] = 1;
                            dt.Rows[0]["index_rows"] = 0;

                            dt.Rows[0]["create_by"] = user_name;
                            dt.Rows[0]["action_type"] = "insert";
                            dt.Rows[0]["action_change"] = 0;
                            dt.AcceptChanges();
                        }
                    }

                    if (dt != null)
                    {
                        id_subareas_def = Convert.ToInt32(dt.Rows[0]["id"].ToString());

                        set_index_worksheet(ref dt);

                        dt.TableName = "subareas";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                    }
                    set_max_id(ref dtma, "subareas", (id_subareas + 1).ToString());

                    #endregion subareas

                    #region hazard 
                    id_hazard = get_max("EPHA_T_TABLE1_HAZARD", seq);

                    sqlstr = @" select b.* , 'update' as action_type, 0 as action_change, 0 as index_rows 
                    from epha_t_header a 
                    inner join EPHA_T_TABLE1_HAZARD b on a.id  = b.id_pha 
                    where 1=1  ";

                    if (!string.IsNullOrEmpty(seq))
                    {
                        sqlstr += $@" and a.seq = @seq";
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                    }
                    sqlstr += $@"  order by a.seq, b.id_type_hazard, b.no";

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

                    if (dt == null || dt?.Rows.Count == 0)
                    {
                        if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                        {
                            dt = new DataTable();
                        }
                        else
                        {
                            dt.Rows.Add(dt.NewRow());
                            dt.Rows[0]["seq"] = id_hazard;
                            dt.Rows[0]["id"] = id_hazard;
                            dt.Rows[0]["id_pha"] = id_pha;
                            dt.Rows[0]["id_subareas"] = id_subareas;

                            dt.Rows[0]["no"] = 1;
                            dt.Rows[0]["index_rows"] = 0;

                            dt.Rows[0]["create_by"] = user_name;
                            dt.Rows[0]["action_type"] = "insert";
                            dt.Rows[0]["action_change"] = 0;
                            dt.AcceptChanges();
                        }
                    }
                    else
                    {
                        bDataHazard = true;
                    }
                    id_hazard_def = Convert.ToInt32(dt.Rows[0]["id"].ToString());

                    if (dt != null)
                    {
                        set_index_worksheet(ref dt);

                        dt.TableName = "hazard";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                    }

                    set_max_id(ref dtma, "hazard", (id_hazard + 1).ToString());

                    #endregion hazard
                }

                //data group table2
                if (true)
                {
                    #region tasks 
                    sqlstr = @" select b.* , 'update' as action_type, 0 as action_change, 0 as index_rows
                    from epha_t_header a inner join EPHA_T_TABLE2_TASKS b on a.id  = b.id_pha
                    where 1=1 ";

                    if (!string.IsNullOrEmpty(seq))
                    {
                        sqlstr += $@" and a.seq = @seq";
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                    }
                    sqlstr += $@"  order by a.seq,b.no ";

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

                    id_tasks = get_max("EPHA_T_TABLE2_TASKS", seq);
                    if (dt == null || dt?.Rows.Count == 0)
                    {
                        if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                        {
                            dt = new DataTable();
                        }
                        else
                        {
                            dt.Rows.Add(dt.NewRow());
                            dt.Rows[0]["seq"] = id_tasks;
                            dt.Rows[0]["id"] = id_tasks;
                            dt.Rows[0]["id_pha"] = id_pha;

                            dt.Rows[0]["no"] = 1;
                            dt.Rows[0]["index_rows"] = 0;
                            dt.Rows[0]["tasks_type_other"] = 0;

                            dt.Rows[0]["create_by"] = user_name;
                            dt.Rows[0]["action_type"] = "insert";
                            dt.Rows[0]["action_change"] = 0;
                            dt.AcceptChanges();
                        }
                    }
                    else
                    {
                        bDataTasks = true;
                    }
                    id_tasks_def = Convert.ToInt32(dt.Rows[0]["id"].ToString());

                    set_max_id(ref dtma, "tasks", (id_tasks + 1).ToString());
                    set_index_worksheet(ref dt);

                    dt.TableName = "tasks";
                    dsData.Tables.Add(dt.Copy());
                    dsData.AcceptChanges();
                    #endregion tasks

                    #region workers 
                    sqlstr = @" select b.* , 'update' as action_type, 0 as action_change, 0 as index_rows
                    from epha_t_header a inner join EPHA_T_TABLE2_WORKERS b on a.id  = b.id_pha
                    where 1=1 ";

                    if (!string.IsNullOrEmpty(seq))
                    {
                        sqlstr += $@" and a.seq = @seq";
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                    }
                    sqlstr += $@"  order by a.seq,b.no ";

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

                    id_workers = get_max("EPHA_T_TABLE2_WORKERS", seq);
                    if (dt == null || dt?.Rows.Count == 0)
                    {
                        if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                        {
                            dt = new DataTable();
                        }
                        else
                        {
                            dt.Rows.Add(dt.NewRow());
                            dt.Rows[0]["seq"] = id_workers;
                            dt.Rows[0]["id"] = id_workers;
                            dt.Rows[0]["id_pha"] = id_pha;
                            dt.Rows[0]["id_tasks"] = id_tasks;

                            dt.Rows[0]["no"] = 1;
                            dt.Rows[0]["index_rows"] = 0;

                            dt.Rows[0]["create_by"] = user_name;
                            dt.Rows[0]["action_type"] = "new";
                            dt.Rows[0]["action_change"] = 0;
                            dt.AcceptChanges();
                        }
                    }
                    id_workers_def = Convert.ToInt32(dt.Rows[0]["id"].ToString());

                    set_max_id(ref dtma, "workers", (id_workers + 1).ToString());
                    set_index_worksheet(ref dt);

                    dt.TableName = "workers";
                    dsData.Tables.Add(dt.Copy());
                    dsData.AcceptChanges();
                    #endregion workers

                    #region description of task
                    sqlstr = @" select b.* , 'update' as action_type, 0 as action_change, 0 as index_rows
                    from epha_t_header a inner join EPHA_T_TABLE2_DESCRIPTIONS b on a.id  = b.id_pha
                    where 1=1 ";

                    if (!string.IsNullOrEmpty(seq))
                    {
                        sqlstr += $@" and a.seq = @seq";
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                    }
                    sqlstr += $@" order by b.no_tasks, b.no";

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

                    id_desc = get_max("EPHA_T_TABLE2_DESCRIPTIONS", seq);
                    if (dt == null || dt?.Rows.Count == 0)
                    {
                        if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                        {
                            dt = new DataTable();
                        }
                        else
                        {
                            dt.Rows.Add(dt.NewRow());
                            dt.Rows[0]["seq"] = id_desc;
                            dt.Rows[0]["id"] = id_desc;
                            dt.Rows[0]["id_pha"] = id_pha;
                            dt.Rows[0]["id_tasks"] = id_tasks;

                            dt.Rows[0]["no"] = 1;
                            dt.Rows[0]["index_rows"] = 0;

                            dt.Rows[0]["create_by"] = user_name;
                            dt.Rows[0]["action_type"] = "new";
                            dt.Rows[0]["action_change"] = 0;
                            dt.AcceptChanges();
                        }

                    }
                    set_max_id(ref dtma, "descriptions", (id_desc + 1).ToString());
                    set_index_worksheet(ref dt);

                    dt.TableName = "descriptions";
                    dsData.Tables.Add(dt.Copy());
                    dsData.AcceptChanges();
                    #endregion description of task
                }

                //data group table3
                if (true)
                {
                    #region worksheet 
                    sqlstr = @" select b.* , 0 as no  
                    , 'update' as action_type, 0 as action_change
                    , b.id_hazard as seq_hazard, b.id_tasks as seq_tasks
                    , vw.user_id as responder_user_id, vw.user_email as responder_user_email
                    , 'assets/img/team/avatar.webp' as responder_user_img
                    , ts.no as tasks_no, hz.no_subareas as subarea_no, hz.no as hazard_no
                    , b.index_rows
                    , b.standard_value + ' '+ b.standard_unit + ' ' + b.standard_desc as tlv_standard
                    from epha_t_header a   
                    inner join EPHA_T_TABLE1_SUBAREAS sa on a.id  = sa.id_pha 
                    inner join EPHA_T_TABLE1_HAZARD hz  on a.id  = hz.id_pha and sa.id = hz.id_subareas
                    inner join EPHA_T_TABLE2_TASKS ts on a.id  = ts.id_pha 
                    inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha and hz.id = b.id_hazard and ts.id = b.id_tasks  
                    left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name) 
                    where 1=1 ";


                    if (!string.IsNullOrEmpty(seq))
                    {
                        sqlstr += $@" and a.seq = @seq";
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                    }
                    sqlstr += $@"  order by a.seq, ts.no, hz.no_subareas, hz.no, b.no ";

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

                    id_worksheet = get_max("EPHA_T_TABLE3_WORKSHEET", seq);
                    if (dt == null || dt?.Rows.Count == 0)
                    {
                        if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                        {
                            dt = new DataTable();
                        }
                        else
                        {
                            dt.Rows.Add(dt.NewRow());
                            dt.Rows[0]["id_pha"] = id_pha;
                            dt.Rows[0]["seq"] = id_worksheet;
                            dt.Rows[0]["id"] = id_worksheet;

                            dt.Rows[0]["id_hazard"] = id_hazard_def;
                            dt.Rows[0]["id_tasks"] = id_tasks_def;

                            dt.Rows[0]["seq_hazard"] = id_hazard_def;
                            dt.Rows[0]["seq_tasks"] = id_tasks_def;
                            dt.Rows[0]["seq_recommendations"] = id_worksheet;
                            dt.Rows[0]["fk_recommendations"] = id_worksheet;

                            dt.Rows[0]["index_rows"] = 0;
                            dt.Rows[0]["no"] = 1;

                            dt.Rows[0]["create_by"] = user_name;
                            dt.Rows[0]["action_type"] = "new";
                            dt.Rows[0]["action_change"] = 0;
                            dt.Rows[0]["action_status"] = "Open";
                            dt.AcceptChanges();
                        }
                    }

                    set_index_worksheet(ref dt);
                    set_max_id(ref dtma, "worksheet", (id_worksheet + 1).ToString());

                    dt.TableName = "worksheet";
                    dsData.Tables.Add(dt.Copy());
                    dsData.AcceptChanges();

                    _recomment_data(user_name, seq, id_pha, ref dtma, id_worksheet);
                    #endregion worksheet
                }

            }
            catch (Exception ex_function) { WriteLog(ex_function.ToString()); }
        }

        private void _whatif_data(ref DataSet? dsData, string user_name, string seq, int id_pha, ref DataTable dtma)
        {
            try
            {

                // ตรวจสอบค่า  เพื่อป้องกันปัญหา Dereference หลังจาก null check
                if (dsData == null || dtma == null || string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(user_name))
                {
                    return;
                }

                seq = seq ?? "-1";

                int id_list_def = 0;
                int id_list = 0;
                int id_listworksheet = 0;

                DataTable dt = new DataTable();
                List<SqlParameter> parameters = new List<SqlParameter>();

                #region list 
                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                from epha_t_header a inner join EPHA_T_LIST b on a.id  = b.id_pha
                where a.seq = @seq
                order by b.no";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

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

                id_list = get_max("EPHA_T_LIST", seq);
                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["seq"] = id_list;
                        dt.Rows[0]["id"] = id_list;
                        dt.Rows[0]["id_pha"] = id_pha;

                        dt.Rows[0]["no"] = 1;

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                    }
                }
                id_list_def = Convert.ToInt32(dt.Rows[0]["id"].ToString());

                set_max_id(ref dtma, "tasklist", (id_list + 1).ToString());

                dt.TableName = "tasklist";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                #endregion list

                #region listdrawing 
                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change 
                from epha_t_header a inner join EPHA_T_LIST_DRAWING b on a.id  = b.id_pha
                where a.seq = @seq
                order by b.no";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

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

                int id_listdrawing = get_max("EPHA_T_LIST_DRAWING", seq);
                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["id_pha"] = id_pha;
                        dt.Rows[0]["seq"] = id_listdrawing;
                        dt.Rows[0]["id"] = id_listdrawing;
                        dt.Rows[0]["id_list"] = id_list_def;

                        dt.Rows[0]["no"] = 1;

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                    }
                }
                set_max_id(ref dtma, "tasklistdrawing", (id_listdrawing + 1).ToString());

                dt.TableName = "tasklistdrawing";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                #endregion listdrawing

                #region listworksheet 
                id_listworksheet = get_max("EPHA_T_LIST_WORKSHEET", seq);

                sqlstr = @" select b.* , 0 as no  
                , 'update' as action_type, 0 as action_change  
                , vw.user_id as responder_user_id, vw.user_email as responder_user_email
                , 'assets/img/team/avatar.webp' as responder_user_img
                , n.no as list_no
                , b.index_rows
                from epha_t_header a   
                inner join EPHA_T_LIST n on a.id  = n.id_pha 
                inner join EPHA_T_LIST_WORKSHEET b on a.id  = b.id_pha and n.id = b.id_list  
                left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name) 
                where a.seq = @seq
                order by b.index_rows, n.no, b.no, b.list_system_no, b.list_sub_system_no, b.causes_no, b.consequences_no, b.category_no, b.recommendations_no, b.recommendations_action_no";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

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

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["id_pha"] = id_pha;
                        dt.Rows[0]["seq"] = id_listworksheet;
                        dt.Rows[0]["id"] = id_listworksheet;
                        dt.Rows[0]["id_list"] = id_list_def;

                        dt.Rows[0]["index_rows"] = 0;
                        dt.Rows[0]["no"] = 1;
                        dt.Rows[0]["row_type"] = "list_system";

                        dt.Rows[0]["seq_list_system"] = 1;
                        dt.Rows[0]["seq_list_sub_system"] = 1;
                        dt.Rows[0]["seq_causes"] = 1;
                        dt.Rows[0]["seq_consequences"] = 1;
                        dt.Rows[0]["seq_category"] = 1;

                        dt.Rows[0]["list_no"] = 1;
                        dt.Rows[0]["list_system_no"] = 1;
                        dt.Rows[0]["list_sub_system_no"] = 1;
                        dt.Rows[0]["causes_no"] = 1;
                        dt.Rows[0]["consequences_no"] = 1;

                        dt.Rows[0]["recommendations_no"] = 1;
                        dt.Rows[0]["recommendations_action_no"] = 1;

                        dt.Rows[0]["implement"] = 0;

                        dt.Rows[0]["create_by"] = user_name;
                        dt.Rows[0]["action_type"] = "new";
                        dt.Rows[0]["action_change"] = 0;
                        dt.Rows[0]["action_status"] = "Open";
                        dt.AcceptChanges();
                    }
                }
                set_index_worksheet(ref dt);

                set_max_id(ref dtma, "listworksheet", (id_listworksheet + 1).ToString());

                dt.TableName = "listworksheet";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

                _recomment_data(user_name, seq, id_pha, ref dtma, id_listworksheet);
                #endregion listworksheet
            }
            catch (Exception ex_function) { WriteLog(ex_function.ToString()); }
        }

        private void set_index_worksheet(ref DataTable _dt)
        {
            if (_dt == null)
            {
                return;
            }
            else
            {
                for (int i = 0; i < _dt?.Rows.Count; i++)
                {
                    _dt.Rows[i]["index_rows"] = i;

                    try
                    {
                        if (string.IsNullOrEmpty(_dt.Rows[i]["implement"]?.ToString()))
                        {
                            _dt.Rows[i]["implement"] = 0;
                        }
                    }
                    catch
                    {
                        // Handle exception if necessary
                    }

                    _dt.AcceptChanges();
                }
            }
        }

        public void check_role_user_active(string user_name, ref string role_type)
        {
            ClassLogin classLogin = new ClassLogin();
            DataTable dtrole = new DataTable();
            dtrole = classLogin.dataEmployeeRole(user_name);

            if (dtrole?.Rows.Count > 0)
            {
                for (int i = 0; i < dtrole?.Rows.Count; i++)
                {
                    role_type = dtrole.Rows[i]["role_type"]?.ToString() ?? "";
                    if (role_type == "admin")
                    {
                        break;
                    }
                }
            }
            else
            {
                dtrole = new DataTable();
                dtrole = classLogin._dtAuthorization_Page(user_name, "");

                if (dtrole?.Rows.Count > 0)
                {
                    for (int i = 0; i < dtrole?.Rows.Count; i++)
                    {
                        role_type = dtrole.Rows[i]["role_type"]?.ToString() ?? "";
                        if (role_type == "admin")
                        {
                            break;
                        }
                    }
                }
            }
        }

        #endregion Data Page Worksheet

        #region Data Page Search
        public string get_search_details(LoadDocModel param)
        {
            // ตรวจสอบค่า param และ user_name เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (param == null || string.IsNullOrEmpty(param.user_name))
            {
                cls_json = new ClassJSON(); return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid parameters."));
            }

            DataSet dsData = new DataSet();
            DataTable dt = new DataTable();
            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string token_doc = param.token_doc?.Trim() ?? "";
            string sub_software = param.sub_software?.Trim() ?? "";
            string type_doc = param.type_doc?.Trim() ?? "";
            string seq = token_doc;

            // ตรวจสอบ sub_software ว่ามีอยู่ใน whitelist
            var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };
            if (!allowedSubSoftware.Contains(sub_software.ToLower()))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid sub_software."));
            }

            get_master_search(ref dsData, sub_software, user_name);

            get_data_search(ref dsData, user_name, seq, sub_software);

            //authorization_page
            get_authorization_page(ref dsData, user_name, role_type);

            //authorization_page_by_doc
            authorization_page_by_doc(ref dsData, user_name, role_type);

            //subsoftware
            dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("name");
            dt.Columns.Add("field_check");
            dt.AcceptChanges();
            string[] xsub_software = ("hazop,jsea,whatif,hra,bowtie").Split(",");
            string[] xsub_software_name = (("hazop,jsea,whatif,hra,bowtie").ToUpper()).Split(",");

            for (int i = 0; i < xsub_software.Length; i++)
            {
                string _id = xsub_software[i];
                string _name = xsub_software_name[i].ToUpper();
                if (dsData.Tables["authorization_page"]?.Rows.Count > 0)
                {
                    if (dsData.Tables["authorization_page"]?.Select("page_controller='" + _id + "'").Length > 0)
                    {
                        dt.Rows.Add(_name, _name, _id);
                        continue;
                    }
                }

                if (dsData.Tables["authorization_page_by_doc"]?.Rows.Count > 0)
                {
                    if (dsData.Tables["authorization_page_by_doc"]?.Select("page_controller='" + _id + "'").Length > 0)
                    {
                        dt.Rows.Add(_name, _name, _id);
                        continue;
                    }
                }
            }

            dt.TableName = "subsoftware";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        //public void DataHazopSearchFollowUp(ref DataSet dsData, string user_name, string sub_software)
        //{
        //    // ตรวจสอบค่า  เพื่อป้องกันปัญหา Dereference หลังจาก null check
        //    if (dsData == null || string.IsNullOrEmpty(user_name) || string.IsNullOrEmpty(sub_software))
        //    {
        //        return;
        //    }

        //    DataTable dt = new DataTable();
        //    DataTable dtma = new DataTable();
        //    int id_pha = 0;
        //    string seq = "";
        //    user_name = (string.IsNullOrEmpty(user_name) ? "" : user_name);

        //    string role_type = "";
        //    check_role_user_active(user_name, role_type,   ref role_type);

        //    string year_now = DateTime.Now.Year.ToString();
        //    if (Convert.ToInt64(year_now) > 2500) { year_now = (Convert.ToInt64(year_now) - 543).ToString(); }

        //    dt = new DataTable();
        //    cls = new ClassFunctions();

        //    var parameters = new List<SqlParameter>();
        //    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });

        //    // --user_name, user_displayname, user_email
        //    string sqlstr = @"SELECT * FROM VW_EPHA_PERSON_DETAILS a WHERE 1=1 ";
        //    sqlstr += " AND LOWER(a.user_name) = LOWER(COALESCE(@user_name, 'x'))  ";
        //    DataTable dtemp = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

        //    #region header
        //    string sqlstr_w = "";
        //    string sqlstr_r = "";
        //    string sqlstr_o = "";

        //    // followup
        //    sqlstr_w = @"SELECT 0 as no, a.pha_sub_software, a.seq as pha_seq,a.pha_no, g.pha_request_name, '' as responder_user_displayname , ''  as responder_user_name_check
        //    ,count(1) as status_total
        //    , count(CASE WHEN LOWER(nw.action_status) IN ( 'closed','close with condition') THEN NULL ELSE 1 END) status_open
        //    , count(CASE WHEN LOWER(nw.action_status) IN ( 'closed','close with condition') THEN 1 ELSE NULL END) status_closed
        //    , 'worksheet' as data_by, '' as responder_user_name
        //    , a.pha_status, CASE WHEN a.pha_status  = 13 THEN 'Waiting Follow Up' ELSE 'Waiting Review Follow Up' END as pha_status_name
        //    , 'update' as action_type, 0 as action_change
        //    FROM epha_t_header a 
        //    INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha 
        //    INNER JOIN EPHA_T_NODE_WORKSHEET nw ON a.id = nw.id_pha  
        //    WHERE a.pha_status IN (13,14) AND nw.responder_user_name IS NOT NULL AND a.seq IN (SELECT MAX(seq) FROM vw_epha_max_seq_by_pha_no GROUP BY pha_no)";

        //    parameters = new List<SqlParameter>();
        //    if (!string.IsNullOrEmpty(seq))
        //    {
        //        sqlstr_w += @" AND a.seq = @seq  ";
        //        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
        //    }
        //    if (role_type != "admin")
        //    {
        //        sqlstr_w += @" AND ( a.pha_status IN (13,14) AND ISNULL(nw.responder_action_type,0) <> 2 )";
        //    }
        //    if (!string.IsNullOrEmpty(user_name) && role_type != "admin")
        //    {
        //        sqlstr_w += @" AND LOWER(nw.responder_user_name) = LOWER(@user_name)  ";
        //        parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
        //    }
        //    if (!string.IsNullOrEmpty(sub_software) && role_type != "admin")
        //    {
        //        sqlstr_w += @" AND LOWER(a.pha_sub_software) = LOWER(@sub_software)  ";
        //        parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });
        //    }

        //    sqlstr_w += @" GROUP BY a.pha_status, a.pha_sub_software, a.seq, a.pha_no, g.pha_request_name";

        //    // followup
        //    sqlstr_r = @"SELECT 0 as no, a.pha_sub_software, '' as pha_seq, '' as pha_no, '' as pha_request_name, vw.user_displayname as responder_user_displayname, LOWER(nw.responder_user_name) as responder_user_name_check
        //    ,count(1) as status_total
        //    , count(CASE WHEN LOWER(nw.action_status) IN ( 'closed','close with condition') THEN NULL ELSE 1 END) status_open
        //    , count(CASE WHEN LOWER(nw.action_status) IN ( 'closed','close with condition') THEN 1 ELSE NULL END) status_closed
        //    , 'responder' as data_by, nw.responder_user_name
        //    , a.pha_status, CASE WHEN a.pha_status  = 13 THEN 'Waiting Follow Up' ELSE 'Waiting Review Follow Up' END as pha_status_name
        //    , 'update' as action_type, 0 as action_change
        //    FROM epha_t_header a 
        //    INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha 
        //    INNER JOIN EPHA_T_NODE_WORKSHEET nw ON a.id = nw.id_pha  
        //    INNER JOIN VW_EPHA_PERSON_DETAILS vw ON LOWER(nw.responder_user_name) = LOWER(vw.user_name) 
        //    WHERE a.pha_status IN (13,14) AND nw.responder_user_name IS NOT NULL AND a.seq IN (SELECT MAX(seq) FROM vw_epha_max_seq_by_pha_no GROUP BY pha_no)";

        //    if (!string.IsNullOrEmpty(seq)) { sqlstr_r += @" AND a.seq = @seq  "; }
        //    if (role_type != "admin") { sqlstr_r += @" AND ( a.pha_status IN (13,14) AND ISNULL(nw.responder_action_type,0) <> 2 )"; }
        //    if (!string.IsNullOrEmpty(user_name) && role_type != "admin") { sqlstr_r += @" AND LOWER(nw.responder_user_name) = LOWER(@user_name)  "; }
        //    if (!string.IsNullOrEmpty(sub_software) && role_type != "admin") { sqlstr_r += @" AND LOWER(a.pha_sub_software) = LOWER(@sub_software)  "; }

        //    sqlstr_r += @" GROUP BY a.pha_status, a.pha_sub_software, vw.user_displayname, nw.responder_user_name";

        //    sqlstr_o = @"SELECT 0 as no, a.pha_sub_software, a.seq as pha_seq, a.pha_no, g.pha_request_name, vw.user_displayname as responder_user_displayname, LOWER(nw.responder_user_name) as responder_user_name_check
        //    ,count(1) as status_total
        //    , count(CASE WHEN LOWER(nw.action_status) IN ( 'closed','close with condition') THEN NULL ELSE 1 END) status_open
        //    , count(CASE WHEN LOWER(nw.action_status) IN ( 'closed','close with condition') THEN 1 ELSE NULL END) status_closed
        //    , 'owner' as data_by, '' as responder_user_name
        //    , a.pha_status, CASE WHEN a.pha_status  = 13 THEN 'Waiting Follow Up' ELSE 'Waiting Review Follow Up' END as pha_status_name
        //    , 'update' as action_type, 0 as action_change 
        //    FROM epha_t_header a 
        //    INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha 
        //    INNER JOIN EPHA_T_NODE_WORKSHEET nw ON a.id = nw.id_pha  
        //    INNER JOIN VW_EPHA_PERSON_DETAILS vw ON LOWER(nw.responder_user_name) = LOWER(vw.user_name) 
        //    WHERE a.pha_status IN (13) AND nw.responder_user_name IS NOT NULL AND a.seq IN (SELECT MAX(seq) FROM vw_epha_max_seq_by_pha_no GROUP BY pha_no)";

        //    if (!string.IsNullOrEmpty(seq)) { sqlstr_o += @" AND a.seq = @seq  "; }
        //    if (role_type != "admin") { sqlstr_o += @" AND ( a.pha_status IN (13) AND ISNULL(nw.responder_action_type,0) <> 2 )"; }
        //    if (!string.IsNullOrEmpty(user_name) && role_type != "admin") { sqlstr_o += @" AND LOWER(nw.responder_user_name) = LOWER(@user_name)  "; }
        //    if (!string.IsNullOrEmpty(sub_software) && role_type != "admin") { sqlstr_o += @" AND LOWER(a.pha_sub_software) = LOWER(@sub_software)  "; }

        //    sqlstr_o += @" GROUP BY a.pha_status, a.pha_sub_software, a.seq, a.pha_no, g.pha_request_name, vw.user_displayname, nw.responder_user_name";

        //    // รวม
        //    sqlstr = "SELECT t.* FROM (" + sqlstr_w + " UNION " + sqlstr_r + " UNION " + sqlstr_o + ")t ORDER BY data_by, pha_sub_software, pha_no, pha_request_name, responder_user_displayname";

        //    dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

        //    if (dt == null || dt?.Rows.Count == 0)
        //    {
        //        if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
        //        {
        //            dt = new DataTable();
        //        }
        //        else
        //        {
        //            // กรณีที่เป็นใบงานใหม่
        //            dt.Rows.Add(dt.NewRow());
        //            dt.Rows[0]["pha_sub_software"] = sub_software;

        //            dt.Rows[0]["action_type"] = "insert";
        //            dt.Rows[0]["action_change"] = 0;

        //            dt.AcceptChanges();
        //        }
        //    }
        //    if (dt != null)
        //    {
        //        for (int i = 0; i < dt?.Rows.Count; i++)
        //        {
        //            dt.Rows[i]["no"] = (i + 1);
        //            dt.AcceptChanges();
        //        }
        //        dt.TableName = "header";
        //        dsData.Tables.Add(dt.Copy());
        //        dsData.AcceptChanges();
        //    }

        //    #region ดึงข้อมูลทั้งหมด ตาม pha_no เอาไปใช้ในการหา seq, pha ในหน้าถัดๆ ไป
        //    sqlstr_r = @"SELECT DISTINCT a.pha_sub_software, a.seq as pha_seq, a.pha_no, g.pha_request_name, a.pha_status, vw.user_displayname as responder_user_displayname
        //     , LOWER(nw.responder_user_name) as responder_user_name
        //    FROM epha_t_header a 
        //    INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha 
        //    INNER JOIN EPHA_T_NODE_WORKSHEET nw ON a.id = nw.id_pha  
        //    INNER JOIN VW_EPHA_PERSON_DETAILS vw ON LOWER(nw.responder_user_name) = LOWER(vw.user_name) 
        //    WHERE a.pha_status IN (13,14) AND nw.responder_user_name IS NOT NULL AND a.seq IN (SELECT MAX(seq) FROM vw_epha_max_seq_by_pha_no GROUP BY pha_no)";

        //    parameters = new List<SqlParameter>();
        //    if (!string.IsNullOrEmpty(seq))
        //    {
        //        sqlstr_r += @" AND a.seq = @seq  ";
        //        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
        //    }
        //    if (role_type != "admin") { sqlstr_r += @" AND ( a.pha_status IN (13,14) AND ISNULL(nw.responder_action_type,0) <> 2 )"; }
        //    if (!string.IsNullOrEmpty(user_name) && role_type != "admin")
        //    {
        //        sqlstr_r += @" AND LOWER(nw.responder_user_name) = LOWER(@user_name)  ";
        //        parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
        //    }
        //    if (!string.IsNullOrEmpty(sub_software) && role_type != "admin")
        //    {
        //        sqlstr_r += @" AND LOWER(a.pha_sub_software) = LOWER(@sub_software)  ";
        //        parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });
        //    }

        //    dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr_r, parameters);

        //    if (dt != null)
        //    {
        //        dt.TableName = "header_all";
        //        dsData.Tables.Add(dt.Copy());
        //        dsData.AcceptChanges();
        //    }

        //    #endregion ดึงข้อมูลทั้งหมด ตาม pha_no เอาไปใช้ในการหา seq, pha ในหน้าถัดๆ ไป

        //    #endregion header

        //    #region general
        //    sqlstr = @"SELECT b.* , '' as functional_location_audition, '' as business_unit_name, '' as unit_no_name, 'update' as action_type, 0 as action_change
        //    , 'PHA No.:' + a.pha_no as txt_project_no
        //    , 'Revision ' + b.pha_request_name as txt_revision 
        //    , (CASE WHEN LOWER(b.expense_type) = 'opex' THEN 'MOC Title ' ELSE 'Project Name ' END) as txt_project_name_header 
        //    , (CASE WHEN LOWER(b.expense_type) = 'opex' THEN 'MOC Title : ' ELSE 'Project Name : ' END) + b.pha_request_name as txt_project_name 
        //    , b.pha_request_name, a.pha_no, a.pha_version, a.pha_no, a.pha_version_text, a.pha_sub_software, a.pha_version_desc
        //    FROM epha_t_header a 
        //    INNER JOIN EPHA_T_GENERAL b ON a.id  = b.id_pha
        //    WHERE 1=2 ";
        //    sqlstr += " ORDER BY a.pha_no";

        //    parameters = new List<SqlParameter>();
        //    dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

        //    if (dt == null || dt?.Rows.Count == 0)
        //    {
        //        if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
        //        {
        //            dt = new DataTable();
        //        }
        //        else
        //        {
        //            // กรณีที่เป็นใบงานใหม่
        //            dt.Rows.Add(dt.NewRow());
        //            dt.Rows[0]["seq"] = id_pha;
        //            dt.Rows[0]["id"] = id_pha; // ข้อมูล 1 ต่อ 1 ให้ใช้กับ header ได้เลย
        //            dt.Rows[0]["id_pha"] = id_pha;

        //            dt.Rows[0]["functional_location_audition"] = "";

        //            // default values 
        //            if (dsData?.Relations?.Count > 0)
        //            {
        //                DataTable dtram = dsData.Tables["ram"].Copy();
        //                dtram.AcceptChanges();
        //                dt.Rows[0]["id_ram"] = dtram.Rows[0]["id"];
        //            }

        //            dt.Rows[0]["expense_type"] = "OPEX";
        //            dt.Rows[0]["sub_expense_type"] = "Normal";

        //            dt.Rows[0]["create_by"] = user_name;
        //            dt.Rows[0]["action_type"] = "insert";
        //            dt.Rows[0]["action_change"] = 0;
        //            dt.AcceptChanges();
        //        }

        //    }
        //    if (dt != null)
        //    {
        //        dt.TableName = "general";
        //        dsData.Tables.Add(dt.Copy());
        //        dsData.AcceptChanges();
        //    }
        //    #endregion general

        //    sqlstr = @" SELECT a.seq, a.pha_no, a.pha_version, a.pha_status, b.pha_request_name, ms.descriptions as pha_status_desc
        //                FROM epha_t_header a
        //                INNER JOIN EPHA_T_GENERAL b ON a.id  = b.id_pha  
        //                LEFT JOIN EPHA_M_STATUS ms ON a.pha_status = ms.id
        //                WHERE 1=1  AND a.seq IN (SELECT MAX(seq) FROM vw_epha_max_seq_by_pha_no GROUP BY pha_no)";

        //    parameters = new List<SqlParameter>();
        //    if (!string.IsNullOrEmpty(seq))
        //    {
        //        sqlstr += @" AND a.seq = @seq  ";
        //        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
        //    }
        //    sqlstr += " ORDER BY a.seq, b.seq";

        //    dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
        //    if (dt != null)
        //    {
        //        dt.TableName = "pha_doc";
        //        dsData.Tables.Add(dt.Copy());
        //        dsData.AcceptChanges();
        //    }
        //    if (dsData != null)
        //    {
        //        dsData.DataSetName = "dsData";
        //        dsData.AcceptChanges();
        //    }
        //}

        public void DataSearchFollowUp(ref DataSet dsData, string user_name, string sub_software)
        {
            if (string.IsNullOrEmpty(user_name)) { return; }//{ throw new ArgumentException("Invalid user_name"); } 
            if (string.IsNullOrEmpty(sub_software)) { return; }//{ throw new ArgumentException("Invalid sub_software."); }

            var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };

            if (!allowedSubSoftware.Contains(sub_software.ToLower()))
            {
                { return; }// throw new ArgumentException("Invalid sub_software value");
            }

            DataTable dt = new DataTable();
            int id_pha = 0;
            string seq = "";
            bool hra_active = sub_software == "hra";

            //string table_worksheet = sub_software switch
            //{
            //    "hazop" => "EPHA_T_NODE_WORKSHEET",
            //    "whatif" => "EPHA_T_LIST_WORKSHEET",
            //    "hra" => "EPHA_T_TABLE3_WORKSHEET",
            //    _ => "EPHA_T_TASKS_WORKSHEET",
            //};

            //// ตรวจสอบชื่อ table ให้มีเฉพาะอักขระที่ปลอดภัย
            //if (!Regex.IsMatch(table_worksheet, @"^[a-zA-Z0-9_]+$"))
            //{
            //    { return; }// throw new ArgumentException("Invalid table name format.");
            //}

            string role_type = "";
            check_role_user_active(user_name, ref role_type);

            if (string.IsNullOrEmpty(role_type)) { }
            string year_now = DateTime.Now.Year.ToString();
            if (Convert.ToInt64(year_now) > 2500)
            {
                year_now = (Convert.ToInt64(year_now) - 543).ToString();
            }

            dt = new DataTable();
            cls = new ClassFunctions();

            var parameters = new List<SqlParameter>();

            string sqlstr = @"SELECT * FROM VW_EPHA_PERSON_DETAILS a WHERE 1=1 AND LOWER(a.user_name) = LOWER(COALESCE(@user_name, 'x'))";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
            //DataTable dtemp = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
            DataTable dtemp = new DataTable();
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
                    dtemp = new DataTable();
                    dtemp = _conn.ExecuteAdapter(command).Tables[0];
                    //dtemp.TableName = "data";
                    dtemp.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            #region header 
            sqlstr = @"select a.* from VW_EPHA_DATA_FOLLOWUP_LIST a where a.seq is not null ";

            if (role_type != "admin")
            {
                sqlstr += " AND ( ( a.pha_status = 13 or a.pha_status = 14) AND ISNULL(a.responder_action_type, 0) <> 2)";
            }

            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " AND a.seq = @seq";
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            }
            //if (!string.IsNullOrEmpty(sub_software) && role_type != "admin")
            if (!string.IsNullOrEmpty(sub_software))
            {
                sqlstr += " AND LOWER(a.pha_sub_software) = LOWER(@sub_software)";
                parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });
            }

            if (!hra_active || (hra_active && user_name != "" && role_type != "admin"))
            {
                sqlstr += " AND a.responder_user_name IS NOT NULL";
                if (!string.IsNullOrEmpty(user_name) && role_type != "admin")
                {
                    sqlstr += " AND LOWER(a.responder_user_name) = LOWER(@user_name)";
                    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
                }
            }
            if (!hra_active)
            {
                sqlstr += " AND a.hra_active = @hra_active";
                parameters.Add(new SqlParameter("@hra_active", SqlDbType.Int) { Value = 0 });
            }
            sqlstr += " ORDER BY a.data_by, a.pha_sub_software, a.pha_no, a.pha_request_name, a.responder_user_displayname ";

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

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                else
                {
                    // กรณีที่เป็นใบงานใหม่
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["pha_sub_software"] = sub_software;

                    dt.Rows[0]["action_type"] = "insert";
                    dt.Rows[0]["action_change"] = 0;

                    dt.AcceptChanges();
                }

            }
            if (dt != null)
            {
                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    dt.Rows[i]["no"] = (i + 1);
                    dt.AcceptChanges();
                }
                dt.TableName = "header";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }

            #endregion header

            #region ดึงข้อมูลทั้งหมด ตาม pha_no เอาไปใช้ในการหา seq, pha ในหน้าถัดๆ ไป

            sqlstr = @"select DISTINCT a.pha_sub_software, a.seq, a.pha_seq, a.pha_no, a.pha_request_name, a.pha_status,
                      a.responder_user_displayname,
                      a.responder_user_name
                      from VW_EPHA_DATA_FOLLOWUP_LIST a where a.seq is not null ";
            if (role_type != "admin")
            {
                sqlstr += " AND ( ( a.pha_status = 13 or a.pha_status = 14) AND ISNULL(a.responder_action_type, 0) <> 2)";
            }

            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " AND a.seq = @seq";
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            }
            if (!string.IsNullOrEmpty(user_name) && role_type != "admin")
            {
                sqlstr += " AND LOWER(a.responder_user_name) = LOWER(@user_name)";
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
            }
            if (!string.IsNullOrEmpty(sub_software) && role_type != "admin")
            {
                sqlstr += " AND LOWER(a.pha_sub_software) = LOWER(@sub_software)";
                parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });
            }
            if (!hra_active)
            {
                sqlstr += " AND a.hra_active = @hra_active";
                parameters.Add(new SqlParameter("@hra_active", SqlDbType.Int) { Value = 0 });
            }

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

            if (dt != null)
            {
                dt.TableName = "header_all";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }
            #endregion ดึงข้อมูลทั้งหมด ตาม pha_no เอาไปใช้ในการหา seq, pha ในหน้าถัดๆ ไป


            #region general
            sqlstr = @"SELECT b.*
                       , '' as functional_location_audition, '' as business_unit_name, '' as unit_no_name, '' as action_type, 0 as action_change
                       , '' as txt_project_no, '' as txt_revision,'' as txt_project_name
                       FROM EPHA_T_GENERAL b WHERE 1=2 ";

            parameters = new List<SqlParameter>();
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

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                else
                {
                    // กรณีที่เป็นใบงานใหม่
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["seq"] = id_pha;
                    dt.Rows[0]["id"] = id_pha; // ข้อมูล 1 ต่อ 1 ให้ใช้กับ header ได้เลย
                    dt.Rows[0]["id_pha"] = id_pha;

                    dt.Rows[0]["functional_location_audition"] = "";

                    // default values
                    if (dsData.Tables["ram"]?.Rows.Count > 0)
                    {
                        DataTable dtram = dsData.Tables["ram"]?.Copy() ?? new DataTable();
                        dtram.AcceptChanges();
                        dt.Rows[0]["id_ram"] = dtram.Rows[0]["id"];
                    }

                    dt.Rows[0]["expense_type"] = "OPEX";
                    dt.Rows[0]["sub_expense_type"] = "Normal";

                    dt.Rows[0]["create_by"] = user_name;
                    dt.Rows[0]["action_type"] = "insert";
                    dt.Rows[0]["action_change"] = 0;
                    dt.AcceptChanges();
                }
            }
            if (dt != null)
            {
                dt.TableName = "general";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }
            #endregion general

            parameters = new List<SqlParameter>();
            sqlstr = @$"SELECT a.seq, a.pha_no, a.pha_version, a.pha_status, b.pha_request_name, ms.descriptions as pha_status_desc
                FROM epha_t_header a
                INNER JOIN EPHA_T_GENERAL b ON a.id = b.id_pha
                LEFT JOIN EPHA_M_STATUS ms ON a.pha_status = ms.id
                WHERE a.seq IN (SELECT MAX(seq) FROM vw_epha_max_seq_by_pha_no GROUP BY pha_no)";

            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " AND a.seq = @seq";
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            }
            sqlstr += " ORDER BY a.seq, b.seq";

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
            if (dt != null)
            {
                dt.TableName = "pha_doc";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }
            if (dsData != null)
            {
                dsData.DataSetName = "dsData";
                dsData.AcceptChanges();
            }
        }

        //public string QueryFollowUpDetail(string seq, string pha_no, string responder_user_name, string sub_software, Boolean bOrderBy, ref List<SqlParameter> parameters)
        //{
        //    sub_software = sub_software?.ToLower() ?? "";

        //    // กำหนด whitelist ของ software ที่อนุญาต
        //    var allowedSoftwares = new List<string> { "hazop", "jsea", "whatif", "hra" };

        //    if (!allowedSoftwares.Contains(sub_software.ToLower()))
        //    {
        //        throw new ArgumentException("Invalid sub_software value.");
        //    }

        //    if (sub_software == "hazop")
        //    {
        //        sqlstr = @"SELECT 'update' AS action_type, 0 AS action_change, 0 AS responder_active_row,
        //            0 AS no, a.id AS id_pha, UPPER(a.pha_sub_software) AS pha_sub_software, a.pha_no, g.pha_request_name,
        //            CASE WHEN ISNULL(nw.action_project_team, 0) > 0 THEN ISNULL(nw.project_team_text, '') ELSE vw.user_displayname END AS responder_user_displayname, nw.responder_user_name,
        //            nw.action_status, COUNT(1) AS status_total,
        //            COUNT(CASE WHEN LOWER(nw.action_status) IN ('closed','responed') THEN NULL ELSE 1 END) AS status_open,
        //            COUNT(CASE WHEN LOWER(nw.action_status) IN ('closed','responed','close with condition') THEN 1 ELSE NULL END) AS status_closed,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_name ELSE nw.document_file_admin_name END AS document_file_name,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_path ELSE nw.document_file_admin_path END AS document_file_path,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_size ELSE nw.document_file_admin_size END AS document_file_size,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_name ELSE '' END AS document_file_name_owner,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_path ELSE '' END AS document_file_path_owner,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_size ELSE '' END AS document_file_size_owner,
        //            FORMAT(nw.estimated_start_date,'dd MMM yyyy') AS estimated_start_date_text,
        //            FORMAT(nw.estimated_end_date,'dd MMM yyyy') AS estimated_end_date_text,
        //            ISNULL(DATEDIFF(day, CASE WHEN nw.estimated_end_date > GETDATE() THEN GETDATE() ELSE nw.estimated_end_date END, GETDATE()), 0) AS over_due,
        //            nw.seq, nw.id, ISNULL(nw.responder_action_type, 0) AS responder_action_type,
        //            nw.consequences_no, nw.recommendations, nw.causes_no, nw.causes_no AS causes, nw.recommendations_no, n.no AS node_no, n.node,
        //            g.id_ram, nw.ram_after_security, nw.ram_after_likelihood, nw.ram_after_risk, nw.ram_action_security, nw.ram_action_likelihood, nw.ram_action_risk,
        //            nw.responder_comment, nw.reviewer_comment, ISNULL(nw.reviewer_action_type, 0) AS reviewer_action_type, ISNULL(nw.implement, 0) AS implement,
        //            ISNULL(nw.action_project_team, 0) AS action_project_team, ISNULL(nw.project_team_text, '') AS project_team_text,
        //            0 AS responder_active_row
        //            FROM epha_t_header a 
        //            INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha 
        //            INNER JOIN EPHA_T_NODE_WORKSHEET nw ON a.id = nw.id_pha  
        //            INNER JOIN EPHA_T_NODE n ON a.id = n.id_pha AND nw.id_node = n.id 
        //            LEFT JOIN VW_EPHA_PERSON_DETAILS vw ON LOWER(nw.responder_user_name) = LOWER(vw.user_name)
        //            WHERE LOWER(nw.ram_after_risk) IN ('h','m') AND a.pha_status IN (13,14) AND a.seq IN (SELECT MAX(seq) FROM vw_epha_max_seq_by_pha_no GROUP BY pha_no)";
        //    }
        //    else if (sub_software == "jsea")
        //    {
        //        sqlstr = @"SELECT 'update' AS action_type, 0 AS action_change, 0 AS responder_active_row,
        //            0 AS no, a.id AS id_pha, UPPER(a.pha_sub_software) AS pha_sub_software, a.pha_no, g.pha_request_name,
        //            vw.user_displayname AS responder_user_displayname, nw.responder_user_name,
        //            nw.action_status, COUNT(1) AS status_total,
        //            COUNT(CASE WHEN LOWER(nw.action_status) IN ('closed','responed') THEN NULL ELSE 1 END) AS status_open,
        //            COUNT(CASE WHEN LOWER(nw.action_status) IN ('closed','responed','close with condition') THEN 1 ELSE NULL END) AS status_closed,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_name ELSE nw.document_file_admin_name END AS document_file_name,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_path ELSE nw.document_file_admin_path END AS document_file_path,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_size ELSE nw.document_file_admin_size END AS document_file_size,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_name ELSE '' END AS document_file_name_owner,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_path ELSE '' END AS document_file_path_owner,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_size ELSE '' END AS document_file_size_owner,
        //            FORMAT(nw.estimated_start_date,'dd MMM yyyy') AS estimated_start_date_text,
        //            FORMAT(nw.estimated_end_date,'dd MMM yyyy') AS estimated_end_date_text,
        //            ISNULL(DATEDIFF(day, CASE WHEN nw.estimated_end_date > GETDATE() THEN GETDATE() ELSE nw.estimated_end_date END, GETDATE()), 0) AS over_due,
        //            nw.seq, nw.id, ISNULL(nw.responder_action_type, 0) AS responder_action_type,
        //            nw.workstep, nw.taskdesc, nw.potentailhazard, nw.possiblecase, nw.recommendations,
        //            nw.workstep_no, nw.taskdesc_no, nw.potentailhazard_no, nw.possiblecase_no, nw.recommendations_no,
        //            g.id_ram, nw.ram_after_security, nw.ram_after_likelihood, nw.ram_after_risk, nw.ram_action_security, nw.ram_action_likelihood, nw.ram_action_risk,
        //            nw.responder_comment, nw.reviewer_comment, ISNULL(nw.reviewer_action_type, 0) AS reviewer_action_type,
        //            0 AS implement, 0 AS action_project_team, '' AS project_team_text,
        //            0 AS responder_active_row
        //            FROM epha_t_header a 
        //            INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha 
        //            INNER JOIN EPHA_T_TASKS_WORKSHEET nw ON a.id = nw.id_pha 
        //            LEFT JOIN VW_EPHA_PERSON_DETAILS vw ON LOWER(nw.responder_user_name) = LOWER(vw.user_name)
        //            WHERE LOWER(nw.ram_after_risk) IN ('h','m') AND a.pha_status IN (13,14) AND a.seq IN (SELECT MAX(seq) FROM vw_epha_max_seq_by_pha_no GROUP BY pha_no)";


        //    }
        //    else if (sub_software == "whatif")
        //    {
        //        sqlstr = @"SELECT 'update' AS action_type, 0 AS action_change, 0 AS responder_active_row,
        //            0 AS no, a.id AS id_pha, UPPER(a.pha_sub_software) AS pha_sub_software, a.pha_no, g.pha_request_name,
        //            CASE WHEN ISNULL(nw.action_project_team, 0) > 0 THEN ISNULL(nw.project_team_text, '') ELSE vw.user_displayname END AS responder_user_displayname, nw.responder_user_name,
        //            nw.action_status, COUNT(1) AS status_total,
        //            COUNT(CASE WHEN LOWER(nw.action_status) IN ('closed','responed') THEN NULL ELSE 1 END) AS status_open,
        //            COUNT(CASE WHEN LOWER(nw.action_status) IN ('closed','responed','close with condition') THEN 1 ELSE NULL END) AS status_closed,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_name ELSE nw.document_file_admin_name END AS document_file_name,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_path ELSE nw.document_file_admin_path END AS document_file_path,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_size ELSE nw.document_file_admin_size END AS document_file_size,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_name ELSE '' END AS document_file_name_owner,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_path ELSE '' END AS document_file_path_owner,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_size ELSE '' END AS document_file_size_owner,
        //            FORMAT(nw.estimated_start_date,'dd MMM yyyy') AS estimated_start_date_text,
        //            FORMAT(nw.estimated_end_date,'dd MMM yyyy') AS estimated_end_date_text,
        //            ISNULL(DATEDIFF(day, CASE WHEN nw.estimated_end_date > GETDATE() THEN GETDATE() ELSE nw.estimated_end_date END, GETDATE()), 0) AS over_due,
        //            nw.seq, nw.id, ISNULL(nw.responder_action_type, 0) AS responder_action_type, nw.recommendations,
        //            nw.list_system_no, nw.list_sub_system_no, nw.consequences_no, nw.recommendations, nw.causes_no, nw.causes_no AS causes, nw.recommendations_no,
        //            n.no AS node_no, g.id_ram, nw.ram_after_security, nw.ram_after_likelihood, nw.ram_after_risk, nw.ram_action_security, nw.ram_action_likelihood, nw.ram_action_risk,
        //            nw.responder_comment, nw.reviewer_comment, ISNULL(nw.reviewer_action_type, 0) AS reviewer_action_type, ISNULL(nw.implement, 0) AS implement,
        //            ISNULL(nw.action_project_team, 0) AS action_project_team, ISNULL(nw.project_team_text, '') AS project_team_text,
        //            0 AS responder_active_row
        //            FROM epha_t_header a 
        //            INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha 
        //            INNER JOIN EPHA_T_LIST_WORKSHEET nw ON a.id = nw.id_pha 
        //            INNER JOIN EPHA_T_LIST n ON a.id = n.id_pha AND nw.id_list = n.id 
        //            LEFT JOIN VW_EPHA_PERSON_DETAILS vw ON LOWER(nw.responder_user_name) = LOWER(vw.user_name)
        //            WHERE LOWER(nw.ram_after_risk) IN ('h','m') AND a.pha_status IN (13,14) AND a.seq IN (SELECT MAX(seq) FROM vw_epha_max_seq_by_pha_no GROUP BY pha_no)";
        //    }
        //    else if (sub_software == "hra")
        //    {
        //        sqlstr = @"SELECT 'update' AS action_type, 0 AS action_change, 0 AS responder_active_row,
        //            0 AS no, a.id AS id_pha, UPPER(a.pha_sub_software) AS pha_sub_software, a.pha_no, g.pha_request_name,
        //            CASE WHEN ISNULL(nw.action_project_team, 0) > 0 THEN ISNULL(nw.project_team_text, '') ELSE vw.user_displayname END AS responder_user_displayname, nw.responder_user_name,
        //            nw.action_status, COUNT(1) AS status_total,
        //            COUNT(CASE WHEN LOWER(nw.action_status) IN ('closed','responed') THEN NULL ELSE 1 END) AS status_open,
        //            COUNT(CASE WHEN LOWER(nw.action_status) IN ('closed','responed','close with condition') THEN 1 ELSE NULL END) AS status_closed,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_name ELSE nw.document_file_admin_name END AS document_file_name,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_path ELSE nw.document_file_admin_path END AS document_file_path,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_size ELSE nw.document_file_admin_size END AS document_file_size,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_name ELSE '' END AS document_file_name_owner,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_path ELSE '' END AS document_file_path_owner,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_size ELSE '' END AS document_file_size_owner,
        //            FORMAT(nw.estimated_start_date,'dd MMM yyyy') AS estimated_start_date_text,
        //            FORMAT(nw.estimated_end_date,'dd MMM yyyy') AS estimated_end_date_text,
        //            ISNULL(DATEDIFF(day, CASE WHEN nw.estimated_end_date > GETDATE() THEN GETDATE() ELSE nw.estimated_end_date END, GETDATE()), 0) AS over_due,
        //            nw.seq, nw.id, ISNULL(nw.responder_action_type, 0) AS responder_action_type,
        //            nw.recommendations, sa.no AS subarea_no, n.no AS tasks_no, hz.no AS hazard_no,
        //            ISNULL(nw.project_team_text, '') AS project_team_text, nw.initial_risk_rating AS initail_risk, nw.residual_risk_rating AS residual_risk,
        //            nw.responder_comment, nw.reviewer_comment, ISNULL(nw.reviewer_action_type, 0) AS reviewer_action_type,
        //            ISNULL(nw.implement, 0) AS implement, ISNULL(nw.action_project_team, 0) AS action_project_team, ISNULL(nw.project_team_text, '') AS project_team_text,
        //            0 AS responder_active_row, nw.id_tasks, n.worker_group, td.descriptions as descriptions, sa.sub_area,
        //            hz.type_hazard, hz.health_hazard AS riskfactors, hz.health_effect_rating, nw.initial_risk_rating, nw.residual_risk_rating
        //            FROM epha_t_header a 
        //            INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha 
        //            INNER JOIN EPHA_T_TABLE1_SUBAREAS sa ON a.id = sa.id_pha
        //            INNER JOIN EPHA_T_TABLE1_HAZARD hz ON a.id = hz.id_pha AND sa.id = hz.id_subareas
        //            INNER JOIN EPHA_T_TABLE2_TASKS n ON a.id = n.id_pha 
        //            INNER JOIN EPHA_T_TABLE2_DESCRIPTIONS td ON a.id = td.id_pha and n.id = td.id_tasks 
        //            INNER JOIN EPHA_T_TABLE3_WORKSHEET nw ON a.id = nw.id_pha AND nw.id_tasks = n.id AND nw.id_hazard = hz.id and nw.id_activity = td.id
        //            LEFT JOIN VW_EPHA_PERSON_DETAILS vw ON LOWER(nw.responder_user_name) = LOWER(vw.user_name)
        //            WHERE nw.id_tasks IS NOT NULL AND nw.recommendations IS NOT NULL AND a.pha_status IN (13,14) AND a.seq IN (SELECT MAX(seq) FROM vw_epha_max_seq_by_pha_no GROUP BY pha_no)";
        //    }


        //    if (!string.IsNullOrEmpty(seq))
        //    {
        //        sqlstr += " AND a.seq = @seq";
        //        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
        //    }

        //    if (!string.IsNullOrEmpty(pha_no))
        //    {
        //        sqlstr += " AND LOWER(a.pha_no) = LOWER(@pha_no)";
        //        parameters.Add(new SqlParameter("@pha_no", SqlDbType.VarChar, 50) { Value = pha_no });
        //    }

        //    if (!string.IsNullOrEmpty(responder_user_name))
        //    {
        //        sqlstr += " AND LOWER(nw.responder_user_name) = LOWER(@responder_user_name)";
        //        parameters.Add(new SqlParameter("@responder_user_name", SqlDbType.VarChar, 50) { Value = responder_user_name });
        //    }

        //    if (sub_software == "hazop")
        //    {
        //        sqlstr += @" GROUP BY a.id, UPPER(a.pha_sub_software), a.pha_no, g.pha_request_name,
        //            CASE WHEN ISNULL(nw.action_project_team, 0) > 0 THEN ISNULL(nw.project_team_text, '') ELSE vw.user_displayname END , nw.responder_user_name,
        //            nw.action_status, 
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_name ELSE nw.document_file_admin_name END,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_path ELSE nw.document_file_admin_path END,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_size ELSE nw.document_file_admin_size END,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_name ELSE '' END,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_path ELSE '' END,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_size ELSE '' END,
        //            FORMAT(nw.estimated_start_date,'dd MMM yyyy'),
        //            FORMAT(nw.estimated_end_date,'dd MMM yyyy'),
        //            ISNULL(DATEDIFF(day, CASE WHEN nw.estimated_end_date > GETDATE() THEN GETDATE() ELSE nw.estimated_end_date END, GETDATE()), 0),
        //            nw.seq, nw.id, ISNULL(nw.responder_action_type, 0),
        //            nw.consequences_no, nw.recommendations, nw.causes_no, nw.causes_no, nw.recommendations_no, n.no, n.node,
        //            g.id_ram, nw.ram_after_security, nw.ram_after_likelihood, nw.ram_after_risk, nw.ram_action_security, nw.ram_action_likelihood, nw.ram_action_risk,
        //            nw.responder_comment, nw.reviewer_comment, ISNULL(nw.reviewer_action_type, 0), ISNULL(nw.implement, 0),
        //            ISNULL(nw.action_project_team, 0), ISNULL(nw.project_team_text, '')";

        //    }
        //    else if (sub_software == "jsea")
        //    {
        //        sqlstr += @" GROUP BY a.id, UPPER(a.pha_sub_software), a.pha_no, g.pha_request_name,
        //            vw.user_displayname, nw.responder_user_name, nw.action_status, 
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_name ELSE nw.document_file_admin_name END,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_path ELSE nw.document_file_admin_path END,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_size ELSE nw.document_file_admin_size END,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_name ELSE '' END,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_path ELSE '' END,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_size ELSE '' END,
        //            FORMAT(nw.estimated_start_date,'dd MMM yyyy'),
        //            FORMAT(nw.estimated_end_date,'dd MMM yyyy'),
        //            ISNULL(DATEDIFF(day, CASE WHEN nw.estimated_end_date > GETDATE() THEN GETDATE() ELSE nw.estimated_end_date END, GETDATE()), 0),
        //            nw.seq, nw.id, ISNULL(nw.responder_action_type, 0),
        //            nw.workstep, nw.taskdesc, nw.potentailhazard, nw.possiblecase, nw.recommendations,
        //            nw.workstep_no, nw.taskdesc_no, nw.potentailhazard_no, nw.possiblecase_no, nw.recommendations_no,
        //            g.id_ram, nw.ram_after_security, nw.ram_after_likelihood, nw.ram_after_risk, nw.ram_action_security, nw.ram_action_likelihood, nw.ram_action_risk,
        //            nw.responder_comment, nw.reviewer_comment, ISNULL(nw.reviewer_action_type, 0)";

        //    }
        //    else if (sub_software == "whatif")
        //    {
        //        sqlstr += @" GROUP BY a.id, nw.seq, nw.id, a.pha_sub_software, a.pha_no, g.pha_request_name,
        //         vw.user_displayname, nw.responder_user_name,
        //         nw.document_file_name, nw.document_file_path, nw.document_file_size, nw.estimated_start_date, nw.estimated_end_date, nw.action_status,
        //         ISNULL(nw.responder_action_type, 0), nw.recommendations,
        //         nw.list_system_no, nw.list_sub_system_no, nw.consequences_no, nw.recommendations, nw.causes_no, nw.causes_no, nw.recommendations_no,
        //         n.no, g.id_ram,
        //         nw.ram_after_security, nw.ram_after_likelihood, nw.ram_after_risk, nw.ram_action_security, nw.ram_action_likelihood, nw.ram_action_risk,
        //         nw.responder_comment, nw.reviewer_comment, nw.document_file_admin_name, nw.document_file_admin_path, nw.document_file_admin_size,
        //         a.pha_status, nw.reviewer_action_type, nw.implement, ISNULL(nw.action_project_team, 0), ISNULL(nw.project_team_text, '')";

        //    }
        //    else if (sub_software == "hra")
        //    {
        //        sqlstr += @" GROUP BY  a.id, UPPER(a.pha_sub_software), a.pha_no, g.pha_request_name,
        //            CASE WHEN ISNULL(nw.action_project_team, 0) > 0 THEN ISNULL(nw.project_team_text, '') ELSE vw.user_displayname END, nw.responder_user_name,
        //            nw.action_status,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_name ELSE nw.document_file_admin_name END,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_path ELSE nw.document_file_admin_path END,
        //            CASE WHEN a.pha_status = 13 THEN nw.document_file_size ELSE nw.document_file_admin_size END,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_name ELSE '' END,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_path ELSE '' END,
        //            CASE WHEN a.pha_status > 13 THEN nw.document_file_size ELSE '' END,
        //            FORMAT(nw.estimated_start_date,'dd MMM yyyy'),
        //            FORMAT(nw.estimated_end_date,'dd MMM yyyy'),
        //            ISNULL(DATEDIFF(day, CASE WHEN nw.estimated_end_date > GETDATE() THEN GETDATE() ELSE nw.estimated_end_date END, GETDATE()), 0),
        //            nw.seq, nw.id, ISNULL(nw.responder_action_type, 0),
        //            nw.recommendations, sa.no, n.no, hz.no,
        //            ISNULL(nw.project_team_text, ''), nw.initial_risk_rating, nw.residual_risk_rating,
        //            nw.responder_comment, nw.reviewer_comment, ISNULL(nw.reviewer_action_type, 0),
        //            ISNULL(nw.implement, 0), ISNULL(nw.action_project_team, 0), ISNULL(nw.project_team_text, ''),
        //            nw.id_tasks, n.worker_group, sa.sub_area,
        //            hz.type_hazard, hz.health_hazard, hz.health_effect_rating, nw.initial_risk_rating, nw.residual_risk_rating, td.descriptions";

        //    }
        //    if (bOrderBy)
        //    {
        //        sqlstr += " ORDER BY CONVERT(int, a.id), UPPER(a.pha_sub_software), a.pha_no, g.pha_request_name, CASE WHEN ISNULL(nw.action_project_team, 0) > 0 THEN ISNULL(nw.project_team_text, '') ELSE vw.user_displayname END";
        //    }

        //    return sqlstr;
        //}

        public void DataSearchFollowUpDetail(ref DataSet dsData, string user_name, string pha_seq, string pha_no, string responder_user_name, string sub_software)
        {
            // ตรวจสอบค่า  เพื่อป้องกันปัญหา Dereference หลังจาก null check
            if (dsData == null || string.IsNullOrEmpty(user_name) || string.IsNullOrEmpty(pha_seq) || string.IsNullOrEmpty(sub_software))
            {
                return;
            }
            string seq = pha_seq ?? "-1";

            if (string.IsNullOrEmpty(sub_software))
            {
                throw new ArgumentException("Invalid sub_software value.");
            }
            // กำหนด whitelist ของ software ที่อนุญาต
            var allowedSoftwares = new List<string> { "hazop", "jsea", "whatif", "hra" };

            if (!allowedSoftwares.Contains(sub_software.ToLower()))
            {
                throw new ArgumentException("Invalid sub_software value.");
            }

            DataTable dt = new DataTable();
            Boolean hra_active = (sub_software == "hra" ? true : false);
            int pha_status = 14;
            string document_module = "followup";
            ClassLogin clslogin = new ClassLogin();
            string role_type = clslogin._dtAuthorization_RoleType(user_name);

            List<SqlParameter> parameters = new List<SqlParameter>();
            dt = new DataTable();
            cls = new ClassFunctions();

            sqlstr = @" select a.id_pha, b.id_tasks, b.user_name, lower(b.user_name) as user_name_check
                from epha_t_table2_workers a 
                inner join epha_t_table2_workers b on a.id = b.id_tasks
                where a.seq in (select max(seq) from vw_epha_max_seq_by_pha_no group by pha_no) ";
            sqlstr += " and a.seq = @seq ";
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });

            if (user_name != "" && role_type != "admin")
            {
                sqlstr += @" and lower(b.user_name) = lower(@user_name) ";
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }

            sqlstr += " order by b.user_name ";

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

            dt.TableName = "workers";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();


            #region header 

            sqlstr = @" select a.pha_no, a.request_user_name from epha_t_header a where a.seq is not null ";
            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });

            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += $" and a.seq = @seq ";
                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            }

            dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
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

            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                else
                {
                    //กรณีที่เป็นใบงานใหม่
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["pha_no"] = "";
                    dt.Rows[0]["request_user_name"] = "";
                    dt.AcceptChanges();
                }
            }
            if (dt != null)
            {
                dt.TableName = "header";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            #endregion header

            #region details

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });

            sqlstr = @" select distinct d.id_tasks, d.no, d.descriptions 
                from epha_t_header a 
                inner join EPHA_T_GENERAL g on a.id = g.id_pha 
                inner join EPHA_T_TABLE1_SUBAREAS sa on a.id = sa.id_pha
                inner join EPHA_T_TABLE1_HAZARD hz on a.id = hz.id_pha and sa.id = hz.id_subareas
                inner join EPHA_T_TABLE2_TASKS n on a.id = n.id_pha  
                inner join EPHA_T_TABLE2_DESCRIPTIONS d on a.id = d.id_pha and n.id = d.id_tasks
                inner join EPHA_T_TABLE3_WORKSHEET nw on a.id = nw.id_pha and nw.id_tasks = n.id and nw.id_hazard = hz.id 
                where nw.id_tasks is not null and a.pha_status in (13,14) and a.seq in (select max(seq) from vw_epha_max_seq_by_pha_no group by pha_no)";
            sqlstr += " and a.seq = @seq  ";

            if (user_name != "" && role_type != "admin")
            {
                sqlstr += @" and lower(nw.responder_user_name) = lower(@user_name) ";
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }
            sqlstr += " order by d.id_tasks, d.no ";

            //DataTable dtDesc = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
            DataTable dtDesc = new DataTable();
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
                    dtDesc = new DataTable();
                    dtDesc = _conn.ExecuteAdapter(command).Tables[0];
                    //dtDesc.TableName = "data";
                    dtDesc.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            // Reset the parameters for the next query
            dt = new DataTable();
            responder_user_name = "";
            //sqlstr = QueryFollowUpDetail(seq, pha_no, responder_user_name, sub_software, true, ref parameters);

            if (sub_software == "hazop")
            {
                sqlstr = "select a.* from VW_EPHA_DATA_FOLLOWUPDETAIL_HAZOP a where a.seq is not null ";
            }
            else if (sub_software == "jsea")
            {
                sqlstr = "select a.* from VW_EPHA_DATA_FOLLOWUPDETAIL_JSEA a  where a.seq is not null ";
            }
            else if (sub_software == "whatif")
            {
                sqlstr = "select a.* from VW_EPHA_DATA_FOLLOWUPDETAIL_WHATIF a  where a.seq is not null ";
            }
            else if (sub_software == "hra")
            {
                sqlstr = "select a.* from VW_EPHA_DATA_FOLLOWUPDETAIL_HRA a  where a.seq is not null ";
            }
            else { throw new ArgumentException("Invalid sub_software value."); }

            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(sqlstr))
            {
                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " AND a.id_pha = @id_pha";
                    parameters.Add(new SqlParameter("@id_pha", SqlDbType.VarChar, 50) { Value = seq });
                }

                if (!string.IsNullOrEmpty(pha_no))
                {
                    sqlstr += " AND LOWER(a.pha_no) = LOWER(@pha_no)";
                    parameters.Add(new SqlParameter("@pha_no", SqlDbType.VarChar, 50) { Value = pha_no });
                }

                if (!string.IsNullOrEmpty(responder_user_name))
                {
                    sqlstr += " AND LOWER(a.responder_user_name) = LOWER(@responder_user_name)";
                    parameters.Add(new SqlParameter("@responder_user_name", SqlDbType.VarChar, 50) { Value = responder_user_name });
                }
            }

            if (!string.IsNullOrEmpty(sqlstr))
            {
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

                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        // For new records
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[0]["pha_sub_software"] = pha_no;
                        dt.Rows[0]["action_type"] = "insert";
                        dt.Rows[0]["action_change"] = 0;
                        dt.AcceptChanges();
                    }
                }
                if (dt != null)
                {
                    if (dt?.Rows.Count > 0)
                    {
                        // Responder active row
                        foreach (DataRow row in dt.Rows)
                        {
                            int responder_active_row = 0;
                            string responderUserName = (row["responder_user_name"] + "").ToString().ToLower();

                            if (role_type == "admin")
                            {
                                responder_active_row = 1;
                            }
                            else if (responderUserName == user_name || role_type == "admin")
                            {
                                responder_active_row = 1;
                            }

                            row["responder_active_row"] = responder_active_row;

                            //if (sub_software == "hra")
                            //{
                            //    // Add descriptions
                            //    string id_tasks = row["id_tasks"]?.ToString() ?? "";
                            //    if (id_tasks != "")
                            //    {
                            //        string descriptions = "";
                            //        DataRow[] drDesc = dtDesc.Select("id_tasks=" + id_tasks);
                            //        if (drDesc != null && drDesc?.Length > 0)
                            //        {
                            //            for (int k = 0; k < drDesc?.Length; k++)
                            //            {
                            //                if (k > 0)
                            //                {
                            //                    descriptions += "\n";
                            //                }
                            //                descriptions += drDesc[k]["descriptions"].ToString();
                            //            }
                            //        }
                            //        row["descriptions"] = descriptions;
                            //    }
                            //}
                        }
                    }
                }
            }
            if (dt != null)
            {
                dt.TableName = "details";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }
            #endregion details

            #region general

            sqlstr = @" select a.pha_status, b.review_folow_comment, 'update' as action_type, 0 as action_change, ms.descriptions as pha_status_desc
                , 'PHA No.: ' + a.pha_no as txt_project_no
                , 'Revision ' + str(a.pha_version) as txt_revision
                , (case when lower(b.expense_type) = 'opex' then 'MOC Title ' else 'Project Name ' end) as txt_project_name_header 
                , (case when lower(b.expense_type) = 'opex' then 'MOC Title : ' else 'Project Name : ' end) + b.pha_request_name as txt_project_name 
                , b.pha_request_name, a.pha_no, a.pha_version, a.pha_no, a.pha_version_text, a.pha_sub_software, a.pha_version_desc
                , ms.descriptions as pha_status_displayname
                from epha_t_header a
                inner join EPHA_T_GENERAL b on a.id = b.id_pha 
                left join EPHA_M_STATUS ms on a.pha_status = ms.id
                where 1=1 and a.seq in (select max(seq) from vw_epha_max_seq_by_pha_no group by pha_no) ";
            sqlstr += "  and a.seq = @seq ";
            sqlstr += " order by a.seq,b.seq";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });

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


            if (dt == null || dt?.Rows.Count == 0)
            {
                if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                {
                    dt = new DataTable();
                }
                else
                {

                    pha_status = Convert.ToInt32(dt.Rows[0]["pha_status"] + "");
                    // Waiting Follow Up, Waiting Review Follow Up
                    if (pha_status == 14)
                    {
                        document_module = "review_followup";
                    }

                    if ((dt.Rows[0]["review_folow_comment"] + "") == "")
                    {
                        // Retrieve all responder comments for the document
                        string responder_comment = "";
                        DataTable dtDetail = dsData.Tables["details"].Copy();
                        dtDetail.AcceptChanges();
                        for (int i = 0; i < dtDetail?.Rows.Count; i++)
                        {
                            responder_comment += (dtDetail.Rows[i]["responder_comment"] + "") + System.Environment.NewLine;
                        }
                        dt.Rows[0]["review_folow_comment"] = responder_comment;
                        dt.AcceptChanges();
                    }
                }
            }
            if (dt != null)
            {
                dt.TableName = "general";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }
            #endregion general

            #region worksheet drawing

            sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                from epha_t_header a inner join EPHA_T_DRAWING_WORKSHEET b on a.id = b.id_pha
                where 1=1 ";
            sqlstr += " and lower(a.seq) = lower(@pha_seq) ";
            sqlstr += " and lower(document_module) = lower(@document_module) ";
            sqlstr += " order by a.seq,b.seq";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@pha_seq", SqlDbType.VarChar, 50) { Value = pha_seq });
            parameters.Add(new SqlParameter("@document_module", SqlDbType.VarChar, 100) { Value = document_module });

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

            int id_drawing = get_max("EPHA_T_DRAWING_WORKSHEET", pha_seq);
            int irow = 0;
            DataTable dtWorksheet = dsData.Tables["details"].Copy();
            dtWorksheet.AcceptChanges();

            if (dtWorksheet != null)
            {
                if (dt == null || dt?.Rows.Count == 0)
                {
                    if (dt == null)  // ตรวจสอบว่า dt ถูกสร้างหรือไม่
                    {
                        dt = new DataTable();
                    }
                    else
                    {
                        for (int i = 0; i < dtWorksheet?.Rows.Count; i++)
                        {
                            string id_worksheet = (dtWorksheet.Rows[i]["seq"] + "").ToString();
                            if (id_worksheet == "")
                            {
                                continue;
                            }
                            DataRow[] dr = dt.Select("id_worksheet=" + id_worksheet);
                            if (dr.Length == 0)
                            {
                                irow = dt?.Rows.Count ?? 0;
                                // For new records
                                dt.Rows.Add(dt.NewRow());
                                dt.Rows[irow]["seq"] = id_drawing;
                                dt.Rows[irow]["id"] = id_drawing;
                                dt.Rows[irow]["id_pha"] = pha_seq;
                                dt.Rows[irow]["id_worksheet"] = id_worksheet;
                                dt.Rows[irow]["no"] = (irow + 1);
                                dt.Rows[irow]["document_module"] = document_module;
                                dt.Rows[irow]["create_by"] = user_name;
                                dt.Rows[irow]["action_type"] = "insert";
                                dt.Rows[irow]["action_change"] = 0;
                                dt.AcceptChanges();
                                id_drawing += 1;
                            }
                        }
                    }
                }
                if (dt != null)
                {
                    dt.TableName = "drawingworksheet";
                    dsData.Tables.Add(dt.Copy());
                    dsData.AcceptChanges();
                }
            }

            DataTable dtma = new DataTable();
            set_max_id(ref dtma, "drawingworksheet", (id_drawing + 1).ToString());
            if (dtma != null)
            {
                dtma.TableName = "max";
                dsData.Tables.Add(dtma.Copy());
                dsData.AcceptChanges();
            }

            #endregion worksheet drawing

            #region worksheet drawing responder & reviewer

            sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                from epha_t_header a inner join EPHA_T_DRAWING_WORKSHEET b on a.id = b.id_pha
                where 1=1 ";
            sqlstr += "  and a.seq = @seq ";
            sqlstr += " and lower(document_module) = lower('followup') ";
            sqlstr += " order by a.seq,b.seq";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });

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
            if (dt != null)
            {
                dt.TableName = "drawingworksheet_responder";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }

            sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                from epha_t_header a inner join EPHA_T_DRAWING_WORKSHEET b on a.id = b.id_pha
                where 1=1 ";
            sqlstr += " and a.seq = @seq ";
            sqlstr += " and lower(document_module) = lower('review_followup') ";
            sqlstr += " order by a.seq,b.seq";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
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

            if (dt != null)
            {
                dt.TableName = "drawingworksheet_reviewer";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }
            #endregion worksheet drawing responder & reviewer

            #region get doc

            sqlstr = @" select a.seq, a.pha_no, a.pha_version, a.pha_status, b.pha_request_name, ms.descriptions as pha_status_desc
                from epha_t_header a
                inner join EPHA_T_GENERAL b on a.id = b.id_pha  
                left join EPHA_M_STATUS ms on a.pha_status = ms.id
                where 1=1 and a.seq in (select max(seq) from vw_epha_max_seq_by_pha_no group by pha_no) ";
            sqlstr += "  and a.seq = @seq ";
            sqlstr += " order by a.seq,b.seq";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });

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

            if (dt != null)
            {
                dt.TableName = "pha_doc";
                dsData.Tables.Add(dt.Copy());
                dsData.AcceptChanges();
            }
            #endregion get doc

            if (dsData != null)
            {
                dsData.DataSetName = "dsData";
                dsData.AcceptChanges();
            }
        }


        #endregion Data Page Search

        #region notification
        public string get_notification(LoadDocModel param)
        {
            DataSet dsData = new DataSet();

            // Define the SQL query with parameter placeholders
            sqlstr = @" select a.pha_sub_software, a.seq as pha_seq, a.pha_no, g.pha_request_name 
                , a.pha_status, ms.descriptions as pha_status_name 
                from epha_t_header a 
                inner join EPHA_T_GENERAL g on a.id = g.id_pha 
                inner join EPHA_M_STATUS ms on a.pha_status = ms.id  
                where a.pha_status in (13,14,21) and a.seq in (select max(seq) from vw_epha_max_seq_by_pha_no group by pha_no) 
                order by a.pha_sub_software, a.pha_no desc";

            // Execute the query and fetch the result
            //DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            DataTable dt = new DataTable();
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


            // Add the resulting DataTable to the DataSet
            dt.TableName = "resulte";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();

            dsData.DataSetName = "dsData";
            dsData.AcceptChanges();

            // Convert the DataSet to a JSON string
            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        #endregion notification 

        #region home task 
        public string get_hometasks(LoadDocModel param)
        {
            DataSet dsData = new DataSet();
            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            string token_doc = param.token_doc?.ToString() ?? "";
            string sub_software = (param.sub_software ?? "");
            string seq = token_doc;

            DataTable dt = DataHomeTask(user_name, role_type, sub_software, false, true);

            dt.TableName = "resultes";
            dsData.Tables.Add(dt.Copy());
            dsData.AcceptChanges();


            DataTable dtMasterList = MasterListHomeTask(sub_software);
            dtMasterList.TableName = "status";
            dsData.Tables.Add(dtMasterList.Copy());
            dsData.AcceptChanges();


            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        public DataTable DataHomeTask(string user_name, string role_type, string sub_software, Boolean group_by_user, Boolean all_sub_software, string status_code = "")
        {
            if (string.IsNullOrEmpty(sub_software))
            {
                throw new ArgumentException("Invalid sub_software value.");
            }
            // กำหนด whitelist ของ software ที่อนุญาต
            var allowedSoftwares = new List<string> { "hazop", "jsea", "whatif", "hra" };

            if (!allowedSoftwares.Contains(sub_software.ToLower()))
            {
                throw new ArgumentException("Invalid sub_software value.");
            }

            string seq = "";

            DataTable dt = new DataTable();
            cls = new ClassFunctions();

            sqlstr = "";
            var parameters = new List<SqlParameter>();

            if (group_by_user == true)
            {
                sqlstr = "select distinct a.user_name, a.user_displayname, a.user_email from VW_EPHA_DATA_HOMETASK a where a.id_pha is not null ";
            }
            else
            {
                sqlstr = @"select distinct id_pha, pha_status, user_name, user_displayname, user_email, user_name_ori, id_action, user_action_date, action_sort, task, pha_type, action_required, document_number, document_title, rev, originator, receivesd, due_date, action_date, remaining, consolidator 
                            from VW_EPHA_DATA_HOMETASK a where a.id_pha is not null";

            }
            if (!string.IsNullOrEmpty(user_name))
            {
                sqlstr += " and lower(a.user_name) = lower(@user_name)";
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }
            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " and a.id_pha = @seq";
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            }
            if (!string.IsNullOrEmpty(status_code))
            {
                sqlstr += " and lower(a.pha_status) = lower(@status_code)";
                parameters.Add(new SqlParameter("@status_code", SqlDbType.VarChar, 50) { Value = status_code });
            }

            if (!all_sub_software)
            {
                if (!string.IsNullOrEmpty(sub_software))
                {
                    sqlstr += " and a.pha_sub_software = @sub_software";
                    parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });
                }
            }

            if (group_by_user == true)
            {
                sqlstr += " order by a.user_name";
            }
            else
            {
                sqlstr += " order by a.user_name, a.action_sort, a.document_number, a.rev";
            }

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

            return dt;
        }
        public DataTable MasterListHomeTask(string sub_software)
        {
            // ตรวจสอบพารามิเตอร์ที่รับเข้ามา
            if (string.IsNullOrEmpty(sub_software))
            {
                //throw new ArgumentException("The sub_software parameter cannot be null or empty.", nameof(sub_software));
                sub_software = "";
            }

            // กำหนด DataTable และเพิ่มคอลัมน์
            DataTable dt = new DataTable();
            dt.Columns.Add("code");
            dt.Columns.Add("name");
            dt.Columns.Add("descriptions");
            dt.Columns.Add("sort_by");

            // กำหนดรายการสถานะ
            string[] statusList = { "Approver", "Approver Approve", "Recommendation Closing", "Approve", "Assigned" };
            string[] descriptionsList = { "Approver", "Approver Approve", "Recommendation Closing", "Approve", "Assigned" };

            //// ใช้ HashSet เพื่อตรวจสอบความซ้ำซ้อน
            //HashSet<string> existingCodes = new HashSet<string>();

            // วนลูปผ่านสถานะต่างๆ
            for (int i = 0; i < statusList.Length; i++)
            {
                string status = statusList[i] ?? "";
                string descriptions = descriptionsList[i] ?? "";

                // ตรวจสอบเงื่อนไขสำหรับ 'sub_software' และ 'status'
                if (sub_software == "jsea" && (status == "Recommendation Closing" || status == "Approve"))
                {
                    // ข้ามสถานะเหล่านี้เมื่อ sub_software เป็น "jsea"
                    continue;
                }

                //// ตรวจสอบไม่ให้มีการเพิ่มข้อมูลซ้ำใน DataTable
                //if (!existingCodes.Contains(status))
                //{
                dt.Rows.Add(status, status, descriptions, i.ToString());
                //    existingCodes.Add(status);
                //}
            }

            return dt;
        }

        #endregion home task 
    }
}
