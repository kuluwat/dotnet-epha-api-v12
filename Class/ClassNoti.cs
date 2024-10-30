using dotnet6_epha_api.Class;
using System.Data;
using System.Data.SqlClient;
namespace Class
{
    public class ClassNoti
    {
        ClassConnectionDb _conn = new ClassConnectionDb();
        string sqlstr = "";

        public DataTable DataDailyByActionRequired(string user_name_active, string seq, string sub_software, Boolean group_by_user, Boolean home_task)
        {
            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            // Define a whitelist of allowed sub_software values
            var allowedSubSoftware = new HashSet<string> { "jsea", "whatif", "hra", "hazop", "hra" };

            // Check if sub_software is valid
            if (!allowedSubSoftware.Contains(sub_software.ToLower()))
            {
                throw new ArgumentException("Invalid sub_software value");
            }
            string sqlstr = "";

            // group_by_user logic
            if (group_by_user)
            {
                sqlstr = "select distinct a.user_name, a.user_displayname, a.user_email from VW_EPHA_ACTION_FOLLOW a ";
                sqlstr += " where lower(a.pha_sub_software) = lower(@pha_sub_software) ";
                parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = sub_software });

                if (!string.IsNullOrEmpty(user_name_active))
                {
                    sqlstr += " and lower(a.user_name) = lower(@user_name)";
                    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name_active });
                }

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " and a.seq = @seq";
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                }

                sqlstr += " order by a.user_name";
            }
            else
            {
                if (home_task)
                {
                    //sqlstr = "select distinct id_pha, pha_status, user_name, user_displayname, user_email, user_name_ori, id_action, user_action_date, action_sort, task, pha_type, action_required, document_number, document_title, rev, originator, receivesd, due_date, action_date, remaining, consolidator from (" + sqlstr + ") t where 1=1";

                    sqlstr = "select distinct select distinct id_pha, pha_status, user_name, user_displayname, user_email, user_name_ori, id_action, user_action_date, action_sort, task, pha_type, action_required, document_number, document_title, rev, originator, receivesd, due_date, action_date, remaining, consolidator  from VW_EPHA_ACTION_FOLLOW a ";
                }
                else
                {
                    sqlstr = "select distinct a.* from VW_EPHA_ACTION_FOLLOW a ";
                }
                sqlstr += " where lower(a.pha_sub_software) = lower(@pha_sub_software) ";
                parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = sub_software });

                if (!string.IsNullOrEmpty(user_name_active))
                {
                    sqlstr += " and lower(a.user_name) = lower(@user_name)";
                    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name_active });
                }

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " and a.seq = @seq";
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                }

                sqlstr += " order by a.user_name, a.action_sort, a.document_number, a.rev";
            }

            // Execute the query and return the result
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
            #region Execute to Datable
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
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
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable



            return dt;
        }

        public DataTable DataDailyByActionRequired_Responder(string seq, string sub_software, Boolean group_by_user)
        {
            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            // Define a whitelist of allowed sub_software values
            var allowedSubSoftware = new HashSet<string> { "jsea", "whatif", "hra", "hazop" };

            // Check if sub_software is valid
            if (!allowedSubSoftware.Contains(sub_software.ToLower()))
            {
                throw new ArgumentException("Invalid sub_software value");
            }

            #region get data 
            string sqlstr = "";

            if (group_by_user)
            {
                sqlstr = "select distinct a.user_name, auser_displayname, a.user_email from VW_EPHA_ACTION_RESPONDER a ";
            }
            else
            {
                sqlstr = "select distinct a.* from VW_EPHA_ACTION_RESPONDER a ";
            }

            if (!string.IsNullOrEmpty(sub_software))
            {
                sqlstr += " and a.pha_sub_software = @sub_software";
                parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software });
            }

            //if (!string.IsNullOrEmpty(seq_worksheet_list))
            //{
            //    // แยกค่าที่คั่นด้วย comma เป็น array
            //    var seqList = seq_worksheet_list.Split(',');

            //    // สร้างเงื่อนไข OR สำหรับแต่ละค่า เช่น (t.seq_worksheet_list = @seq0 OR t.seq_worksheet_list = @seq1)
            //    var orConditions = seqList.Select((s, index) => $"a.seq_worksheet_list = @seq_worksheet_list{index}").ToArray();

            //    // รวมเงื่อนไข OR เข้าด้วยกัน
            //    sqlstr += $" and ({string.Join(" OR ", orConditions)})";

            //    // เพิ่ม parameters สำหรับแต่ละค่าใน seq_worksheet_list
            //    for (int i = 0; i < seqList.Length; i++)
            //    {
            //        parameters.Add(new SqlParameter($"@seq_worksheet_list{i}", SqlDbType.VarChar, 50) { Value = seqList[i] });
            //    }
            //}

            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " and a.seq = @seq";
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
            }

            if (group_by_user)
            {
                sqlstr += " order by a.user_name";
            }
            else
            {
                sqlstr += " order by a.user_name, a.action_sort, a.document_number, a.rev";
            }

            if (sqlstr.Length > 0)
            {
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                 
                #region Execute to Datable
                try
                {
                    _conn = new ClassConnectionDb(); _conn.OpenConnection();
                    try
                    {
                        var command = _conn.conn.CreateCommand();
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
                        dt.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


            }
            else
            {
                dt = new DataTable();
            }

            #endregion get data 

            return dt;
        }

        public DataTable DataDailyByActionRequired_TeammMember(string user_name_active, string seq, string sub_software, Boolean group_by_user)
        {
            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            // Define a whitelist of allowed sub_software values
            var allowedSubSoftware = new HashSet<string> { "jsea", "whatif", "hra", "hazop" };

            // Check if sub_software is valid
            if (!allowedSubSoftware.Contains(sub_software.ToLower()))
            {
                throw new ArgumentException("Invalid sub_software value");
            }

            #region get data  
            string sqlstr = "";
            if (group_by_user)
            {
                sqlstr = "SELECT DISTINCT a.user_name, a.user_displayname, a.user_email from VW_EPHA_ACTION_TEAMMEMBER a WHERE a.user_name IS NOT NULL  AND a.pha_status >= 12 ";
            }
            else
            {
                sqlstr = "SELECT DISTINCT a.* from VW_EPHA_ACTION_TEAMMEMBER a WHERE a.user_name IS NOT NULL  AND a.pha_status >= 12 ";
            }

            // Use SqlParameter for 'seq'
            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " AND a.seq = @seq";  // Changed t.id_pha to a.id for proper reference
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
            }

            // Grouping by user or returning full data
            if (group_by_user)
            {
                sqlstr += " ORDER BY a.user_name";
            }
            else
            {
                sqlstr += " ORDER BY a.user_name, a.action_sort, a.document_number, a.rev";
            }

            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters); 
            #region Execute to Datable
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
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
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            #endregion get data

            return dt;
        }

        public DataTable DataDailyByActionRequired_ReviewApprove(string id_pha, string responder_user_name, string sub_software, Boolean group_by_user, Boolean responder_close_all)
        {
            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            // Define a whitelist of allowed sub_software values
            var allowedSubSoftware = new HashSet<string> { "jsea", "whatif", "hra", "hazop", "hra" };

            // Check if sub_software is valid
            if (!allowedSubSoftware.Contains(sub_software.ToLower()))
            {
                throw new ArgumentException("Invalid sub_software value");
            }

            #region get data
            string sqlstr = "";

            if (group_by_user)
            {
                sqlstr = "SELECT DISTINCT a.user_name, a.user_displayname, a.user_email FROM VW_EPHA_ACTION_REVIEWAPPROVE a ";
            }
            else
            {
                sqlstr = "SELECT DISTINCT a.* FROM VW_EPHA_ACTION_REVIEWAPPROVE a ";
            }

            if (responder_close_all)
            {
                sqlstr += " where a.pha_status IN (14)";
            }
            else
            {
                sqlstr += " where a.pha_status IN (13)";
            }


            // Add parameters for responder_user_name and id_pha
            if (!string.IsNullOrEmpty(responder_user_name))
            {
                sqlstr += " AND LOWER(a.user_name) = LOWER(@responder_user_name)";
                parameters.Add(new SqlParameter("@responder_user_name", SqlDbType.VarChar, 50) { Value = responder_user_name });
            }
            if (!string.IsNullOrEmpty(id_pha))
            {
                sqlstr += " AND a.id_pha = @id_pha";
                parameters.Add(new SqlParameter("@id_pha", SqlDbType.VarChar, 50) { Value = id_pha });
            }

            // Grouping or full data retrieval
            if (group_by_user)
            {
                sqlstr = " ORDER BY a.user_name, a.user_displayname, a.user_email ";
            }
            else
            {
                sqlstr += " ORDER BY a.user_name, a.action_sort, a.document_number, a.rev";
            }

            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters); 
            #region Execute to Datable
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
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
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            #endregion get data

            return dt;
        }

    }
}
