using dotnet_epha_api.Class;
using dotnet6_epha_api.Class;
using Model;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
namespace Class
{
    public class ClassManage
    {
        string sqlstr = "";
        string jsper = "";
        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb _conn = new ClassConnectionDb();

        #region home task 
        public string DocumentCopy(ManageDocModel param)
        {
            try
            {
                DataSet dsData = new DataSet();
                string userName = param.user_name?.ToString() ?? "";
                string subSoftware = param.sub_software?.ToString() ?? "";
                string phaNo = param.pha_no?.ToString() ?? "";
                string phaSeq = param.pha_seq?.ToString() ?? "";

                string user_name = userName?.ToString() ?? "";
                string role_type = ClassLogin.GetUserRoleFromDb(user_name);
                // ตรวจสอบสิทธิ์ก่อนดำเนินการ
                if (!ClassLogin.IsAuthorized(user_name))
                {
                    return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
                }

                // Define a whitelist of allowed sub_software values
                var allowedSubSoftware = new HashSet<string> { "jsea", "whatif", "hra", "hazop" };

                // Check if sub_software is valid
                if (!allowedSubSoftware.Contains(subSoftware))
                {
                    return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid sub_software value."));
                }

                // ตรวจสอบว่า sub_software มีเฉพาะตัวอักษร a-z, A-Z, ตัวเลข และ underscore
                if (!Regex.IsMatch(subSoftware, @"^[a-zA-Z0-9_]+$"))
                {
                    return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid sub_software format."));
                }

                // ตรวจสอบว่า phaSeq เป็นตัวเลขเท่านั้น
                if (!Regex.IsMatch(phaSeq, @"^\d+$"))
                {
                    return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid pha_seq format."));
                }


                string phaNoNow = "";
                string versionNow = "";
                string seqHeaderNow = phaSeq;
                string phaStatusNow = "11";
                string phaSubSoftware = subSoftware;
                string ret = "";
                string msg = "";

                // pha_sub_software from epha_t_header
                ClassFunctions cls = new ClassFunctions();
                using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();
                    string sqlstr = "select distinct pha_sub_software from epha_t_header where seq = @seq";
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = phaSeq });

                    DataTable dt = new DataTable();
                    using (var cmd = new SqlCommand(sqlstr, conn))
                    {
                        cmd.Parameters.AddRange(parameters.ToArray());
                        using (var da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt);
                        }
                    }

                    if (dt?.Rows.Count > 0)
                    {
                        phaSubSoftware = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                    }

                    // Copy seqHeaderNow => New Seq
                    ClassHazopSet clsHazopSet = new ClassHazopSet();
                    ret = clsHazopSet.keep_version(user_name, role_type, ref seqHeaderNow, ref versionNow, phaStatusNow, phaSubSoftware, false, false, false, false);
                    if (!string.IsNullOrEmpty(ret))
                    {
                        versionNow = "1";
                        if (!string.IsNullOrEmpty(seqHeaderNow))
                        {
                            string yearNow = DateTime.Now.Year.ToString();
                            if (Convert.ToInt64(yearNow) > 2500)
                            {
                                yearNow = (Convert.ToInt64(yearNow) - 543).ToString();
                            }
                            ClassHazop getPhaNo = new ClassHazop();
                            phaNoNow = getPhaNo.get_pha_no(phaSubSoftware, yearNow);

                            string userNameChkSqlStr = cls.ChkSqlStr(userName, 100);
                            string phaNoChkSqlStr = cls.ChkSqlStr(phaNoNow, 100);
                            string requestUserDisplayName = "";

                            ClassLogin clsLogin = new ClassLogin();
                            DataTable dtUser = clsLogin.dataUserRole(userName);
                            if (dtUser?.Rows.Count > 0)
                            {
                                requestUserDisplayName = cls.ChkSqlStr(dtUser.Rows[0]["user_displayname"]?.ToString() ?? "", 4000);
                            }

                            using (ClassConnectionDb transaction = new ClassConnectionDb())
                            {
                                transaction.OpenConnection();
                                transaction.BeginTransaction();

                                try
                                {
                                    // pha_request_by, request_user_name, request_user_displayname
                                    sqlstr = @"update epha_t_header set flow_mail_to_member = null, pha_version = 1, pha_version_text = 'A', pha_version_desc = 'Issued for Review',
                               update_date = null, update_by = null, create_date = getdate(),
                               create_by = @userName, pha_request_by = @userName, request_user_name = @userName, request_user_displayname = @requestUserDisplayName,
                               pha_status = @phaStatusNow, pha_no = @phaNo where id = @seqHeaderNow";

                                    parameters = new List<SqlParameter>();
                                    parameters.Add(new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName });
                                    parameters.Add(new SqlParameter("@seqHeaderNow", SqlDbType.VarChar, 50) { Value = seqHeaderNow });
                                    parameters.Add(new SqlParameter("@requestUserDisplayName", SqlDbType.VarChar, 4000) { Value = requestUserDisplayName });
                                    parameters.Add(new SqlParameter("@phaStatusNow", SqlDbType.VarChar, 50) { Value = phaStatusNow });
                                    parameters.Add(new SqlParameter("@phaNo", SqlDbType.VarChar, 100) { Value = phaNoNow });

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
                                        sqlstr = @"update epha_t_member_team set action_review = null, date_review = null, comment = null,
                               update_date = null, update_by = null, create_date = getdate(),
                               create_by = @userName where id_pha = @seqHeaderNow";

                                        parameters = new List<SqlParameter>();
                                        parameters.Add(new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName });
                                        parameters.Add(new SqlParameter("@seqHeaderNow", SqlDbType.VarChar, 50) { Value = seqHeaderNow });

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
                                    }
                                    if (ret == "true")
                                    {
                                        sqlstr = @"update epha_t_approver set action_review = null, date_review = null, comment = null, approver_action_type = null, action_status = null,
                                                update_date = null, update_by = null, create_date = getdate(),
                                                create_by = @userName where id_pha = @seqHeaderNow";

                                        parameters = new List<SqlParameter>();
                                        parameters.Add(new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName });
                                        parameters.Add(new SqlParameter("@seqHeaderNow", SqlDbType.VarChar, 50) { Value = seqHeaderNow });

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
                                    }
                                    if (ret == "true")
                                    {
                                        try
                                        { 
                                            parameters = new List<SqlParameter>();
                                            parameters.Add(new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName });
                                            parameters.Add(new SqlParameter("@seqHeaderNow", SqlDbType.VarChar, 50) { Value = seqHeaderNow });

                                            if (!string.IsNullOrEmpty(sqlstr))
                                            {
                                                if (!ClassLogin.IsAuthorizedRole(user_name, role_type))
                                                {
                                                    ret = "User is not authorized to perform this action.";
                                                }
                                                else
                                                {
                                                    #region ExecuteNonQuerySQL Data
                                                    var command = transaction.conn.CreateCommand();
                                                    command.CommandType = CommandType.StoredProcedure;
                                                    command.CommandText = "usp_UpdateTablesDocumentCopy";
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
                                            }
                                        }
                                        catch (Exception ex_UpdateTablesDocumentCopy) { ret = ex_UpdateTablesDocumentCopy.Message.ToString(); }

                                    }

                                    if (ret == "true")
                                    {

                                        sqlstr = phaSubSoftware switch
                                        {
                                            "hazop" => @"update epha_t_node_worksheet set responder_action_type = null, responder_action_date = null, responder_receivesd_date = null,
                                     responder_comment = null, reviewer_action_type = null, reviewer_action_date = null, reviewer_comment = null,
                                     update_date = null, update_by = null, create_date = getdate(), create_by = @userName where id_pha = @seqHeaderNow",
                                            "whatif" => @"update epha_t_list_worksheet set responder_action_type = null, responder_action_date = null, responder_receivesd_date = null,
                                     responder_comment = null, reviewer_action_type = null, reviewer_action_date = null, reviewer_comment = null,
                                     update_date = null, update_by = null, create_date = getdate(), create_by = @userName where id_pha = @seqHeaderNow",
                                            "jsea" => @"update epha_t_tasks_worksheet set responder_action_type = null, responder_action_date = null, responder_receivesd_date = null,
                                    responder_comment = null, reviewer_action_type = null, reviewer_action_date = null, reviewer_comment = null,
                                    update_date = null, update_by = null, create_date = getdate(), create_by = @userName where id_pha = @seqHeaderNow",
                                            "hra" => @"update epha_t_table3_worksheet set responder_action_type = null, responder_action_date = null, responder_receivesd_date = null,
                                   responder_comment = null, reviewer_action_type = null, reviewer_action_date = null, reviewer_comment = null,
                                   update_date = null, update_by = null, create_date = getdate(), create_by = @userName where id_pha = @seqHeaderNow",
                                            _ => throw new ArgumentException("Invalid sub_software value")
                                        };


                                        parameters = new List<SqlParameter>();
                                        parameters.Add(new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName });
                                        parameters.Add(new SqlParameter("@seqHeaderNow", SqlDbType.VarChar, 50) { Value = seqHeaderNow });

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

                        }

                    }
                }

                return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, (seqHeaderNow == phaSeq ? "" : seqHeaderNow), seqHeaderNow, phaNoNow, phaStatusNow));

            }
            catch (Exception ex_function) { return ex_function.ToString(); }

        }
         
        public string DocumentCancel(ManageDocModel param)
        {
            DataSet dsData = new DataSet();
            string subSoftware = param.sub_software?.ToString() ?? "";
            string phaNo = param.pha_no?.ToString() ?? "";
            string phaSeq = param.pha_seq?.ToString() ?? "";
            string phaStatusComment = param.pha_status_comment?.ToString() ?? "";

            string user_name = param.user_name?.ToString() ?? "";
            string role_type = ClassLogin.GetUserRoleFromDb(user_name);
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            string phaSubSoftware = subSoftware;
            string ret = "";
            string msg = "";

            cls = new ClassFunctions();

            using (ClassConnectionDb transaction = new ClassConnectionDb())
            {
                transaction.OpenConnection();
                transaction.BeginTransaction();

                try
                {
                    sqlstr = @"update epha_t_header set pha_status = 81, pha_status_comment = @phaStatusComment, update_date = getdate(), update_by = @userName where id = @phaSeq";
                    var parameters = new List<SqlParameter>
                        {
                            new SqlParameter("@phaStatusComment", SqlDbType.VarChar, 4000) { Value = phaStatusComment },
                            new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = user_name },
                            new SqlParameter("@phaSeq", SqlDbType.VarChar, 50) { Value = phaSeq }
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

            return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, "", "", "", ""));
        }
        #endregion home task 
    }
}
