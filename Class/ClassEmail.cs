
using dotnet_epha_api.Class;
using dotnet6_epha_api.Class;
using Microsoft.Exchange.WebServices.Data;
using Model;
using System.Data;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace Class
{
    public class ClassEmail
    {
        private  ClassConnectionDb _conn;

        private byte[] GetKey()
        {
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            string keyBase64 = config["AesKey"];
            if (string.IsNullOrEmpty(keyBase64))
            {
                throw new InvalidOperationException("The AES key is missing in the configuration.");
            }
            return Convert.FromBase64String(keyBase64);
        }
        private byte[] GetIV()
        {
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            string ivBase64 = config["AesIV"];
            if (string.IsNullOrEmpty(ivBase64))
            {
                throw new InvalidOperationException("The AES IV is missing in the configuration.");
            }
            return Convert.FromBase64String(ivBase64);
        }

        public string EncryptString(string plainText)
        {
            if (string.IsNullOrEmpty(plainText))
            {
                throw new ArgumentException("plainText cannot be null or empty", nameof(plainText));
            }

            byte[] key = GetKey();
            byte[] iv = GetIV();

            return EncryptDataWithAes(plainText, key, iv);
        }

        private string EncryptDataWithAes(string plainText, byte[] key, byte[] iv)
        {
            if (plainText == null) throw new ArgumentNullException(nameof(plainText));
            if (key == null) throw new ArgumentNullException(nameof(key));
            if (iv == null) throw new ArgumentNullException(nameof(iv));

            using (Aes aesAlgorithm = Aes.Create())
            {
                if (aesAlgorithm == null)
                {
                    throw new InvalidOperationException("Failed to create AES algorithm instance.");
                }

                aesAlgorithm.Key = key;
                aesAlgorithm.IV = iv;

                ICryptoTransform encryptor = aesAlgorithm.CreateEncryptor(aesAlgorithm.Key, aesAlgorithm.IV);

                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {
                            swEncrypt.Write(plainText);
                        }
                        byte[] encrypted = msEncrypt.ToArray();
                        return Convert.ToBase64String(encrypted);
                    }
                }
            }
        }

        string server_url = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["WebServer_ePHA_Index"] ?? "";
        string server_url_home_task = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["WebServer_ePHA_HomeTask"] ?? "";

        string sqlstr = "";
        string jsper = "";
        string s_subject = "";
        string s_body = "";

        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb cls_conn = new ClassConnectionDb();

        private bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            Uri redirectionUri = new Uri(redirectionUrl);
            return redirectionUri.Scheme == "https";
        }
        public class sendEmailModel
        {
            public string mail_from { get; set; }
            public string mail_to { get; set; }
            public string mail_cc { get; set; }
            public string mail_subject { get; set; }
            public string mail_body { get; set; }
            public string mail_attachments { get; set; }

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

        private string server_url_by_action(string sub_software)
        {
            if (string.IsNullOrEmpty(sub_software))
            {
                throw new ArgumentException("Sub software cannot be null or empty", nameof(sub_software));
            }

            string xreplace_text = sub_software.ToLower();
            return server_url.Replace("dummy", xreplace_text);
        }

        public string sendMail(sendEmailModel value)
        {
            string s_mail_to = (value.mail_to ?? "");
            string s_mail_cc = (value.mail_cc ?? "");
            string s_subject = (value.mail_subject ?? "");
            string s_mail_body = (value.mail_body ?? "");
            string s_mail_attachments = (value.mail_attachments ?? "");

            string msg_mail = "";
            bool SendAndSaveCopy = false;

            // Load mail configuration from appsettings.json
            //IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            string mail_from = ""; //config.GetSection("MailConfig")["MailFrom"] ?? "";
            string mail_test = ""; //config.GetSection("MailConfig")["MailTest"] ?? "";
            string mail_user = "";//config.GetSection("MailConfig")["MailUser"] ?? "";
            string mail_password = ""; //config.GetSection("MailConfig")["MailPassword"] ?? "";

            // Get test config from the database
            string sqlstr = @"SELECT DISTINCT LOWER(key_name) AS key_name, key_value  FROM epha_m_config  WHERE active_type = 1";

            DataTable dtConfig = new DataTable();
            //dtConfig = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable 
            try
            {
                _conn = new ClassConnectionDb();
                _conn.OpenConnection();
                try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter);
                    dtConfig = new DataTable();
                    dtConfig = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "Table1";
                    dtConfig.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable



            if (dtConfig?.Rows.Count > 0)
            {
                // เก็บคีย์ที่ต้องการตรวจสอบใน array
                string[] keys = { "mailuser", "mailpassword", "mailfrom", "mailtest" };

                // วนลูปผ่านคีย์ต่างๆ และดึงค่าจาก DataTable
                foreach (var key in keys)
                {
                    DataRow[] dr = dtConfig.AsEnumerable().Where(row => row.Field<string>("key_name") == key).ToArray();
                    if (dr?.Length > 0)
                    {
                        string x = dr[0]["key_value"]?.ToString() ?? "";

                        // ตรวจสอบและกำหนดค่าตามคีย์
                        switch (key)
                        {
                            case "mailuser":
                                mail_user = x;
                                break;
                            case "mailpassword":
                                mail_password = x; // ปลดล็อกถ้าจำเป็น: mail_password = Decrypt(value);
                                break;
                            case "mailfrom":
                                mail_from = x;
                                break;
                            case "mailtest":
                                mail_test = x;
                                break;
                        }
                    }
                    else
                    {
                        return $"{key} not found in config.";
                    }
                }
            }
            else
            {
                return "No config mail.";
            }


            // Get test emails from the database
            sqlstr = @"SELECT DISTINCT EMAIL, EMAIL AS USER_EMAIL FROM EPHA_M_CONFIGMAIL WHERE ACTIVE_TYPE = 1";
            DataTable dt = new DataTable();
            //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            #region Execute to Datable 
            try
            {
                _conn = new ClassConnectionDb(); _conn.OpenConnection(); try
                {
                    var command = _conn.conn.CreateCommand();
                    command.CommandText = sqlstr;
                    //command.Parameters.Add(":costcenter", costcenter);
                    dt = new DataTable();
                    dt = _conn.ExecuteAdapter(command).Tables[0];
                    //dt.TableName = "Table1";
                    dt.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable 

            if (dt?.Rows.Count > 0)
            {
                mail_test = string.Join(";", dt.AsEnumerable().Select(row => row["user_email"].ToString()));
            }
            if (!string.IsNullOrEmpty(mail_test))
            {
                s_mail_body += $"</br></br>ข้อมูล mail to: {s_mail_to}</br></br>ข้อมูล mail cc: {s_mail_cc}";
                s_mail_to = mail_test;
                s_mail_cc = mail_test;
            }

            string mail_font = "Cordia New";
            string mail_fontsize = "18";

            s_mail_body = $"<html><div style='font-size:{mail_fontsize}px; font-family:{mail_font};'>{s_mail_body}</div></html>";

            try
            {
                ExchangeService service = new ExchangeService();
                service.Credentials = new WebCredentials(mail_user, mail_password);
                service.TraceEnabled = true;
                service.AutodiscoverUrl(mail_user, RedirectionUrlValidationCallback);

                EmailMessage email = new EmailMessage(service)
                {
                    From = new EmailAddress("Mail Display ใส่ไม่มีผล", mail_from),
                    Subject = !string.IsNullOrEmpty(mail_test) ? "(DEV)" + s_subject : s_subject,
                    Body = new MessageBody(BodyType.HTML, s_mail_body)
                };

                foreach (string to in s_mail_to.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    email.ToRecipients.Add(to);
                }

                foreach (string cc in s_mail_cc.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    email.CcRecipients.Add(cc);
                }

                if (!string.IsNullOrEmpty(s_mail_attachments))
                {
                    foreach (string attachment in s_mail_attachments.Split(new[] { '|', '|' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        email.Attachments.AddFileAttachment(attachment);
                    }
                }

                if (SendAndSaveCopy)
                {
                    email.SendAndSaveCopy();
                }
                else
                {
                    email.Send();
                }

                //msg_mail = "Email sent successfully.";
                msg_mail = "";
            }
            catch (Exception ex)
            {
                msg_mail = ex.Message;
            }

            return msg_mail;
        }
        //public string SendMailSync(sendEmailModel value)
        //{
        //    string msg_mail = string.Empty;

        //    // ข้อมูลอีเมล
        //    string s_mail_to = value.mail_to ?? "";
        //    string s_mail_cc = value.mail_cc ?? "";
        //    string s_subject = value.mail_subject ?? "";
        //    string s_mail_body = value.mail_body ?? "";
        //    string s_mail_attachments = value.mail_attachments ?? "";

        //    try
        //    {
        //        // สร้าง MSAL client สำหรับขอ access token
        //        IConfidentialClientApplication confidentialClient = ConfidentialClientApplicationBuilder.Create("your-client-id")
        //            .WithTenantId("your-tenant-id")
        //            .WithClientSecret("your-client-secret")
        //            .Build();

        //        // ขอ access token แบบ synchronous
        //        AuthenticationResult authResult = confidentialClient.AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
        //            .ExecuteAsync().GetAwaiter().GetResult();

        //        // สร้าง GraphServiceClient สำหรับส่งอีเมล
        //        GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
        //        {
        //            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
        //            return Task.CompletedTask;
        //        }));

        //        // เตรียมข้อมูลอีเมล
        //        var message = new Message
        //        {
        //            Subject = s_subject,
        //            Body = new ItemBody
        //            {
        //                ContentType = BodyType.Html,
        //                Content = $"<html><div style='font-size:18px; font-family:Cordia New;'>{s_mail_body}</div></html>"
        //            },
        //            ToRecipients = s_mail_to.Split(';', StringSplitOptions.RemoveEmptyEntries).Select(email => new Recipient
        //            {
        //                EmailAddress = new EmailAddress { Address = email }
        //            }).ToList(),
        //            CcRecipients = s_mail_cc.Split(';', StringSplitOptions.RemoveEmptyEntries).Select(email => new Recipient
        //            {
        //                EmailAddress = new EmailAddress { Address = email }
        //            }).ToList()
        //        };

        //        // เพิ่มไฟล์แนบถ้ามี
        //        if (!string.IsNullOrEmpty(s_mail_attachments))
        //        {
        //            foreach (string attachmentPath in s_mail_attachments.Split('|', StringSplitOptions.RemoveEmptyEntries))
        //            {
        //                byte[] fileBytes = System.IO.File.ReadAllBytes(attachmentPath);
        //                string fileName = System.IO.Path.GetFileName(attachmentPath);

        //                message.Attachments.Add(new FileAttachment
        //                {
        //                    ODataType = "#microsoft.graph.fileAttachment",
        //                    ContentBytes = fileBytes,
        //                    Name = fileName
        //                });
        //            }
        //        }

        //        // ส่งอีเมลแบบ synchronous
        //        graphClient.Me.SendMail(message, false).Request().PostAsync().GetAwaiter().GetResult();

        //        msg_mail = "Email sent successfully.";
        //    }
        //    catch (Exception ex)
        //    {
        //        msg_mail = $"Error: {ex.Message}";
        //    }

        //    return msg_mail;
        //}

        #region  mail create user
        public string MailToAdminRegisterAccount(string _user_displayname, string _user_email)
        {
            try
            {

                if (string.IsNullOrEmpty(_user_displayname)) { return "Invalid User Displayname."; }
                if (string.IsNullOrEmpty(_user_email)) { return "Invalid User email."; }

                string user_displayname = _user_displayname;

                string urlAccept = "";
                string urlNotAccept = "";

                string to_displayname = "All";
                string s_mail_to = get_mail_admin_group();  // ฟังก์ชันที่ดึงกลุ่มอีเมลของผู้ดูแลระบบ
                string s_mail_cc = "";
                string s_mail_from = "";

                string msg = "";

                // Generate URLs for Accept and Not Accept actions
                if (true)
                {
                    // URL for Accept
                    string plainText = $"user_email={_user_email}&accept_status=1";
                    string cipherText = EncryptString(plainText);  // เข้ารหัสข้อมูล
                    urlAccept = $"{server_url.Replace("dummy", "login").Replace("index", "RegisterAccount")}?data={cipherText}";

                    // URL for Not Accept
                    plainText = $"user_email={_user_email}&accept_status=0";
                    cipherText = EncryptString(plainText);
                    urlNotAccept = $"{server_url.Replace("dummy", "login").Replace("index", "RegisterAccount")}?data={cipherText}";
                }

                // ตั้งหัวข้ออีเมล
                string s_subject = "ePHA Staff or Contractor Register Account.";

                // เนื้อหาอีเมล (ลบการแสดงรหัสผ่าน)
                string s_body = "<html><body><font face='tahoma' size='2'>";
                s_body += $"Dear {to_displayname},";
                s_body += $"<br/><br/>{user_displayname} has registered for an account.";
                s_body += $"<br/>Email address: {_user_email}";
                s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review the registration.";
                s_body += $"<br/><font color='blue'><a href='{urlAccept}'>Accept</a></font>";
                s_body += $",<font color='red'><a href='{urlNotAccept}'>Not Accept</a></font>";
                s_body += "<br/><br/>Best Regards,";
                s_body += "<br/>ePHA Online System";
                s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
                s_body += "</font></body></html>";

                // ส่งอีเมล
                sendEmailModel data = new sendEmailModel
                {
                    mail_subject = s_subject,
                    mail_body = s_body,
                    mail_to = s_mail_to,
                    mail_cc = s_mail_cc,
                    mail_from = s_mail_from
                };

                msg = sendMail(data);

                return msg;
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }
        public string MailToUserRegisterAccount(string _user_displayname, string _user_email, string _accept_status)
        {
            if (string.IsNullOrEmpty(_user_displayname)) { return "Invalid User Displayname."; }
            if (string.IsNullOrEmpty(_user_email)) { return "Invalid User email."; }
            if (string.IsNullOrEmpty(_accept_status)) { return "Invalid Accept Statusl."; }

            try
            {

                // ตรวจสอบค่าที่เป็น null และจัดการค่าพื้นฐาน
                string user_displayname = _user_displayname ?? "User";
                string user_email = _user_email ?? "";
                string url = "";
                string to_displayname = user_displayname;
                string s_mail_to = user_email;
                string s_mail_cc = "";
                string s_mail_from = "";

                string msg = "";

                // URL for user login (ควรใช้การตรวจสอบความปลอดภัยเพิ่มเติม)
                if (!string.IsNullOrEmpty(user_email))
                {
                    string plainText = $"user_email={user_email}";
                    string cipherText = EncryptString(plainText);
                    url = $"{server_url_by_action("login")}{cipherText}";
                }

                string s_subject = "ePHA Staff or Contractor Register Account.";

                // เนื้อหาอีเมล (ลบการแสดงรหัสผ่าน)
                string s_body = "<html><body><font face='tahoma' size='2'>";
                s_body += $"Dear {to_displayname},";
                s_body += "<br/><br/>Register account.";
                s_body += $"<br/><br/>Name: {user_displayname}";
                s_body += $"<br/>Email address: {user_email}";
                s_body += $"<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Admin has {(_accept_status == "1" ? "accepted" : "not accepted")} your registration account.";
                s_body += $"<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;System Administrator has {(_accept_status == "1" ? "confirmed" : "not confirmed")} your system registration.";

                if (_accept_status == "1")
                {
                    s_body += $"<br/><font color='red'>You can now click <a href='{url}'>the link</a> to access the system.</font>";
                }

                s_body += "<br/><br/>Best Regards,";
                s_body += "<br/>ePHA Online System";
                s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
                s_body += "</font></body></html>";

                // ส่งอีเมล
                sendEmailModel data = new sendEmailModel
                {
                    mail_subject = s_subject,
                    mail_body = s_body,
                    mail_to = s_mail_to,
                    mail_cc = s_mail_cc,
                    mail_from = s_mail_from
                };

                msg = sendMail(data);

                return msg;
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }
        #endregion  mail create user

        #region mail workflow
        public string get_mail_admin_group()
        {
            StringBuilder emailList = new StringBuilder();

            ClassLogin cls_login = new ClassLogin();
            DataTable dt = cls_login.dataEmployeeRole("");

            if (dt?.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    string user_email = row["user_email"]?.ToString() ?? "";
                    string role_type = row["role_type"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(user_email)! && string.IsNullOrEmpty(role_type))
                    {
                        if (role_type == "admin")
                        {
                            if (emailList.Length > 0)
                            {
                                emailList.Append(";");
                            }
                            emailList.Append(user_email);
                        }
                    }
                }
            }

            return emailList.ToString();
        }

        public string QueryActionOwnerUpperTA2(string seq, string approver_user_name, string sub_software, ref List<SqlParameter> parameters)
        {
            cls = new ClassFunctions();
            StringBuilder sqlstr = new StringBuilder();

            sqlstr.Append(@"
                            select h.pha_status, h.pha_sub_software, h.pha_no, g.pha_request_name as pha_name, empre.user_email as request_email,
                            nw.responder_user_name, emp.user_displayname, emp.user_email,
                            count(1) as total,
                            count(case when lower(nw.action_status) = 'open' then 1 else null end) as 'open',
                            count(case when lower(nw.action_status) = 'closed' then 1 else null end) as 'closed',
                            g.reference_moc 
                            from epha_t_header h
                            inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                            inner join EPHA_T_SESSION s on lower(h.id) = lower(s.id_pha)  
                            inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s1 on h.id = s1.id_pha and s.id = s1.id_session and s.id_pha = s1.id_pha  
                            inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and s1.id_session = ta2.id_session and ta2.action_review not in (2) 
                            left join EPHA_T_NODE_WORKSHEET nw on lower(h.id) = lower(nw.id_pha) 
                            left join VW_EPHA_PERSON_DETAILS emp on lower(nw.responder_user_name) = lower(emp.user_name)  
                            left join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)  
                            where nw.responder_user_name is not null and h.seq = @seq ");

            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar) { Value = seq });

            if (!string.IsNullOrEmpty(approver_user_name))
            {
                sqlstr.Append(" and lower(ta2.user_name) = lower(@approver_user_name) ");
                parameters.Add(new SqlParameter("@approver_user_name", SqlDbType.VarChar, 50) { Value = approver_user_name });
            }

            sqlstr.Append(@"
                 group by h.pha_status, h.pha_sub_software, h.pha_no, g.pha_request_name, empre.user_email, nw.responder_user_name, emp.user_displayname, emp.user_email,
                 nw.action_status, g.reference_moc");

            string query = sqlstr.ToString();

            // เนื่องจากโครงสร้าง field ที่นำมาใช้งานเหมือนกัน
            if (sub_software == "jsea")
            {
                query = query.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_TASKS_WORKSHEET");
            }
            else if (sub_software == "whatif")
            {
                query = query.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_LIST_WORKSHEET");
            }

            return query;
        }

        //public string MailToActionOwner(string seq, string sub_software)
        //{
        //    if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
        //    if (string.IsNullOrEmpty(sub_software)) { return "Invalid Sub Software."; }
        //    try
        //    {

        //        // Define a whitelist of allowed sub_software values
        //        var allowedSubSoftware = new HashSet<string> { "hazop", "jsea", "whatif", "hra" };

        //        // Check if sub_software is valid
        //        if (!allowedSubSoftware.Contains(pha_sub_software.ToLower()))
        //        {
        //            return ("Invalid sub_software value");
        //        }

        //        #region call function export full report pdf  
        //        var sameRamTypeSoftwares = new HashSet<string> { "hazop", "jsea", "whatif" };
        //        string _ram_type = sameRamTypeSoftwares.Contains(sub_software) ? "5" : "";

        //        ReportByWorksheetModel param = new ReportByWorksheetModel
        //        {
        //            seq = seq,
        //            export_type = "pdf",
        //            sub_software = sub_software,
        //            user_name = "",
        //            ram_type = _ram_type
        //        };

        //        ClassExcel clsExcel = new ClassExcel();
        //        string file_fullpath_name = "";
        //        file_fullpath_name = clsExcel.export_report_recommendation(param, false, true, true);

        //        if (!string.IsNullOrEmpty(file_fullpath_name))
        //        {
        //            string file_fullpath_def = file_fullpath_name;
        //            string folder = sub_software ?? "";
        //            string msg_error = "";

        //            //string folder = sub_software;
        //            string fullPath = "";

        //            #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

        //            bool isValid = true;

        //            // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
        //            if (!Path.IsPathRooted(file_fullpath_def))
        //            {
        //                isValid = false;
        //                msg_error = "File path is not an absolute path.";
        //            }

        //            // ตรวจสอบชื่อไฟล์
        //            string safeFileName = Path.GetFileName(file_fullpath_def);
        //            if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
        //            {
        //                isValid = false;
        //                msg_error = "Invalid or potentially dangerous file name.";
        //            }

        //            // ตรวจสอบอักขระที่อนุญาต
        //            char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
        //                .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

        //            // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
        //            string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
        //            string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
        //            if (isValid && !allowedExtensions.Contains(fileExtension))
        //            {
        //                isValid = false;
        //                msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
        //            }
        //            // ตรวจสอบค่า folder
        //            if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
        //            {
        //                isValid = false;
        //                msg_error = "Invalid folder.";
        //            }
        //            if (isValid)
        //            {
        //                // สร้างพาธของโฟลเดอร์ที่อนุญาต
        //                string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
        //                if (isValid && !Directory.Exists(allowedDirectory))
        //                {
        //                    isValid = false;
        //                    msg_error = "Folder directory not found.";
        //                }

        //                // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
        //                string sourceFile = $"{safeFileName}";
        //                if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
        //                                string.IsNullOrWhiteSpace(sourceFile) ||
        //                                sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
        //                                sourceFile.Contains("..") || sourceFile.Contains("\\")))
        //                {
        //                    isValid = false;
        //                    msg_error = "Invalid fileName.";
        //                }

        //                // สร้างพาธเต็มของไฟล์
        //                if (isValid)
        //                {
        //                    fullPath = Path.Combine(allowedDirectory, sourceFile);
        //                    fullPath = Path.GetFullPath(fullPath);

        //                    // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
        //                    if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
        //                    {
        //                        isValid = false;
        //                        msg_error = "File is outside the allowed directory.";
        //                    }
        //                }

        //                // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
        //                if (isValid)
        //                {
        //                    FileInfo template = new FileInfo(fullPath);
        //                    if (!template.Exists)
        //                    {
        //                        isValid = false;
        //                        msg_error = "File not found.";
        //                    }

        //                    // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
        //                    if (isValid && template.IsReadOnly)
        //                    {
        //                        try
        //                        {
        //                            template.IsReadOnly = false;
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            isValid = false;
        //                            msg_error = $"Failed to modify file attributes: {ex.Message}";
        //                        }
        //                    }
        //                }

        //                // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
        //                if (isValid)
        //                {
        //                    try
        //                    {
        //                        using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
        //                        {
        //                            // สามารถเปิดไฟล์ได้
        //                        }
        //                    }
        //                    catch (UnauthorizedAccessException)
        //                    {
        //                        isValid = false;
        //                        msg_error = "Access to the file is denied.";
        //                    }
        //                }
        //            }
        //            #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


        //            // หากทุกอย่างผ่านการตรวจสอบ
        //            if (isValid)
        //            {
        //                // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
        //                file_fullpath_name = fullPath;
        //            }
        //            else { file_fullpath_name = ""; }
        //        }
        //        else { file_fullpath_name = ""; }

        //        #endregion call function export full report pdf 

        //        string doc_no = "";
        //        string doc_name = "";
        //        string reference_moc = "";

        //        string url = "";
        //        string step_text = "PHA Follow up Item";

        //        string to_displayname = "";
        //        string s_mail_to = "";
        //        string s_mail_cc = "";
        //        string s_mail_from = "";

        //        DataTable dt = new DataTable();

        //        // Query the Action Owner securely with parameters
        //        //string sqlstr = QueryActionOwner(seq, "", sub_software, ref parameters);

        //        sqlstr = @"select a.* from VW_EPHA_ACTION_OWNER a where a.responder_user_name is not null and a.seq = @seq  and a.sub_software = @sub_software ";

        //        var parameters = new List<SqlParameter>();
        //        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar) { Value = seq });
        //        parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar) { Value = sub_software });

        //        dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

        //        if (!string.IsNullOrEmpty(doc_no))
        //        {
        //            #region url  
        //            string plainText = $"seq={seq}&pha_no={doc_no}&step=3";
        //            string cipherText = EncryptString(plainText);

        //            url = $"{server_url_by_action(sub_software)}{cipherText}";
        //            #endregion url 

        //            #region mail to
        //            string msg = "";
        //            if (dt?.Rows.Count > 0)
        //            {
        //                string xbefor = "";
        //                string xafter = "";
        //                for (int i = 0; i < dt?.Rows.Count; i++)
        //                {
        //                    xbefor = dt.Rows[i]["user_displayname"]?.ToString() ?? "";
        //                    if (xbefor != xafter)
        //                    {
        //                        xafter = xbefor;
        //                    }
        //                    else
        //                    {
        //                        if (i != dt?.Rows.Count - 1) { continue; }
        //                    }

        //                    s_mail_cc = dt.Rows[i]["request_email"]?.ToString() ?? "";
        //                    s_mail_to = dt.Rows[i]["user_email"]?.ToString() ?? "";

        //                    doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
        //                    doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
        //                    reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
        //                    to_displayname = dt.Rows[i]["user_displayname"]?.ToString() ?? "";

        //                    int iTotal = Convert.ToInt32(dt.Rows[i]["total"].ToString());
        //                    int iOpen = Convert.ToInt32(dt.Rows[i]["open"].ToString());
        //                    int iClosed = Convert.ToInt32(dt.Rows[i]["closed"].ToString());

        //                    string s_subject = $"EPHA {doc_no}{(doc_name == "" ? "" : "")}, Please follow up item and update action.";

        //                    string s_body = $@"
        //        <html>
        //        <body>
        //            <font face='tahoma' size='2'>
        //                Dear {to_displayname},<br/><br/>
        //                <b>Step</b> : {step_text}<br/>
        //                {(reference_moc != "" ? $"<b>Reference MOC</b> : {reference_moc}<br/>" : "")}
        //                <b>Project Name</b> : {doc_name}<br/><br/>
        //                Items Status Total: {iTotal}, Open: {iOpen}, Closed: {iClosed}<br/><br/>
        //                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No.{doc_no}<br/>
        //                To see the detailed information, <font color='red'> please click <a href='{url}'>here</a></font><br/><br/>
        //                Best Regards,<br/>
        //                ePHA Online System<br/><br/><br/>
        //                Note that this message was automatically sent by ePHA Online System.
        //            </font>
        //        </body>
        //        </html>";

        //                    sendEmailModel data = new sendEmailModel
        //                    {
        //                        mail_subject = s_subject,
        //                        mail_body = s_body,
        //                        mail_to = s_mail_to,
        //                        mail_cc = s_mail_cc,
        //                        mail_from = s_mail_from
        //                    };

        //                    if (!string.IsNullOrEmpty(file_fullpath_name) && File.Exists(file_fullpath_name))
        //                    {
        //                        data.mail_attachments = file_fullpath_name;
        //                    }

        //                    msg = sendMail(data);
        //                    if (!string.IsNullOrEmpty(msg))
        //                    {
        //                        // Handle email sending error if needed
        //                    }
        //                }
        //            }
        //            #endregion mail to

        //            return msg;
        //        }
        //        else
        //        {
        //            return "";
        //        }

        //    }
        //    catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        //}


        public string MailToAdminCaseStudy(string seq, string sub_software)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            if (string.IsNullOrEmpty(sub_software)) { return "Invalid Sub Software."; }

            try
            {

                // Define a whitelist of allowed sub_software values
                var allowedSubSoftware = new HashSet<string> { "hazop", "whatif" };
                if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                {
                    return "Invalid sub_software value.";
                }

                string doc_no = "";
                string doc_name = "";
                string reference_moc = "";
                string url = "";
                string step_text = "Original Closed PHA.";

                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";
                string pha_sub_software = "";
                string meeting_date = "";
                string meeting_time = "";
                DataTable dt = new DataTable();

                string sqlstr = "";
                sqlstr = @" SELECT DISTINCT a.* FROM  VW_EPHA_DATA_CASESTUDY a WHERE a.seq = @seq  ORDER BY a.no";

                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });

                // Execute SQL query
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
                try
                {
                    _conn = new ClassConnectionDb(); _conn.OpenConnection(); try
                    {
                        var command = _conn.conn.CreateCommand();
                        command.CommandText = sqlstr;
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
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

                #region mail to
                if (dt?.Rows.Count > 0)
                {
                    doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                    doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                    reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                    pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                    pha_sub_software = (pha_sub_software == "hazop" || pha_sub_software == "jsea" || pha_sub_software == "hra") ? pha_sub_software.ToUpper() : pha_sub_software;
                    meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                    meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (i > 0) { s_mail_to += ";"; }
                        s_mail_to += dt.Rows[i]["user_email"]?.ToString() ?? "";
                    }
                }
                #endregion mail to

                #region mail cc 
                if (dt?.Rows.Count > 0)
                {
                    s_mail_cc = dt.Rows[0]["request_email"]?.ToString() ?? "";
                }
                #endregion mail cc


                if (!string.IsNullOrEmpty(doc_no))
                {
                    #region url encryption 
                    string plainText = $"seq={seq}&pha_no={doc_no}&step=2";
                    string cipherText = EncryptString(plainText);

                    url = $"{server_url_by_action(sub_software)}{cipherText}";
                    #endregion url encryption

                    string s_subject = $"EPHA : {pha_sub_software.ToUpper()}, Please Review data.";

                    string s_body = $@"
    <html>
    <body>
        <font face='tahoma' size='2'>
            Dear All,<br/><br/>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            Original Closed the {pha_sub_software.ToUpper()}, PHA No.{doc_no} as details mentioned below,<br/><br/>
            <b>Project Name</b> : {doc_name}<br/>
            <b>Date</b> : {meeting_date}<br/>
            <b>Time</b> : {meeting_time}<br/><br/>
            <b>Step</b> : {step_text}<br/>
            {(string.IsNullOrEmpty(reference_moc) ? "" : $"<b>Reference MOC</b> : {reference_moc}<br/>")}
            <b>Project Name</b> : {doc_name}<br/><br/>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No.{doc_no}<br/>
            To see the detailed information,<font color='red'> please click <a href='{url}'>here</a></font><br/><br/>
            More details of the study, <font color='red'; text-decoration: underline;><a href='{url}'> please click here</a></font><br/><br/>
            DISCLAIMER:<br/>
            This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s) of Thai Oil Public Company Limited.
        </font>
    </body>
    </html>";

                    sendEmailModel data = new sendEmailModel
                    {
                        mail_subject = s_subject,
                        mail_body = s_body,
                        mail_to = s_mail_to,
                        mail_cc = s_mail_cc,
                        mail_from = s_mail_from
                    };

                    return sendMail(data);
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }

        #endregion mail workflow

        #region mail workflow last version 
        public string MailNotificationWorkshopInvitation(string seq, string sub_software)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            if (string.IsNullOrEmpty(sub_software)) { return "Invalid Sub Software."; }

            try
            {
                // Define a whitelist of allowed sub_software values
                //var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };
                var allowedSubSoftware = new HashSet<string> { "hazop", "jsea" };
                if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                {
                    return "Invalid sub_software value.";
                }

                string doc_no = "";
                string doc_name = "";
                string reference_moc = "";
                string url = "";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                string pha_sub_software = sub_software;
                string meeting_date = "";
                string meeting_time = "";

                DataTable dt = new DataTable();

                string sqlstr = "";
                sqlstr = @" SELECT DISTINCT a.* FROM  VW_EPHA_DATA_CASESTUDY a WHERE a.seq = @seq ORDER BY a.no ";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "" });

                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
                try
                {
                    _conn = new ClassConnectionDb(); _conn.OpenConnection(); try
                    {
                        var command = _conn.conn.CreateCommand();
                        command.CommandText = sqlstr;
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
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

                #region mail to
                if (dt?.Rows.Count > 0)
                {
                    doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                    doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                    reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                    pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString()?.ToUpper() ?? "What If";
                    meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                    meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (i > 0) { s_mail_to += ";"; }
                        s_mail_to += dt.Rows[i]["user_email"]?.ToString() ?? "";
                    }
                }
                #endregion mail to

                #region mail cc 
                if (dt?.Rows.Count > 0)
                {
                    s_mail_cc = dt.Rows[0]["request_email"]?.ToString() ?? "";
                }
                #endregion mail cc

                if (!string.IsNullOrEmpty(doc_no))
                {
                    #region url encryption
                    string plainText = $"seq={seq}&pha_no={doc_no}&step=2";
                    string cipherText = EncryptString(plainText);

                    url = $"{server_url_by_action(sub_software)}{cipherText}";
                    #endregion url encryption

                    string s_subject = $"EPHA : {pha_sub_software} Workshop Invitation";

                    string s_body = $@"
                     <html>
                     <body>
                         <font face='tahoma' size='2'>
                             Dear All,<br/><br/>
                             &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                             All are invited to participate in the {pha_sub_software} Workshop, PHA No.{doc_no} as detailed below:<br/><br/>
                             <b>Project Name</b> : {doc_name}<br/>
                             <b>Date</b> : {meeting_date}<br/>
                             <b>Time</b> : {meeting_time}<br/>
                             {(string.IsNullOrEmpty(reference_moc) ? "" : $"<br/><b>Reference MOC</b> : {reference_moc}<br/>")}
                             <br/><br/>
                             More details of the study, <font color='red'><a href='{url}'>please click here</a></font><br/><br/>
                             DISCLAIMER:<br/>
                             This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s) of Thai Oil Public Company Limited.
                         </font>
                     </body>
                     </html>";

                    sendEmailModel data = new sendEmailModel
                    {
                        mail_subject = s_subject,
                        mail_body = s_body,
                        mail_to = s_mail_to,
                        mail_cc = s_mail_cc,
                        mail_from = s_mail_from
                    };

                    return sendMail(data);
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }

        public string MailNotificationMemberReview(string seq, string sub_software)
        {
            // ตรวจสอบว่า seq ไม่เป็นค่าว่างหรือ null
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            if (string.IsNullOrEmpty(sub_software)) { return "Invalid Sub Software."; }

            try
            {
                // ตรวจสอบ sub_software ว่ามีอยู่ใน whitelist
                var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };
                if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                {
                    return "Invalid sub_software.";
                }

                #region call function export full report pdf  
                var sameRamTypeSoftwares = new HashSet<string> { "hazop", "jsea", "whatif" };
                string _ram_type = sameRamTypeSoftwares.Contains(sub_software) ? "5" : "";

                ReportModel param = new ReportModel
                {
                    seq = seq,
                    export_type = "pdf",
                    sub_software = sub_software,
                    user_name = "",
                    ram_type = _ram_type
                };

                ClassExcel clsExcel = new ClassExcel();
                string file_fullpath_name = "";
                file_fullpath_name = clsExcel.export_full_report(param, true);

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = sub_software ?? "";
                    string msg_error = "";

                    //string folder = sub_software;
                    string fullPath = "";

                    #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

                    bool isValid = true;

                    // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
                    if (!Path.IsPathRooted(file_fullpath_def))
                    {
                        isValid = false;
                        msg_error = "File path is not an absolute path.";
                    }

                    // ตรวจสอบชื่อไฟล์
                    string safeFileName = Path.GetFileName(file_fullpath_def);
                    if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
                    {
                        isValid = false;
                        msg_error = "Invalid or potentially dangerous file name.";
                    }

                    // ตรวจสอบอักขระที่อนุญาต
                    char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                        .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                    // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                    string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
                    string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                    if (isValid && !allowedExtensions.Contains(fileExtension))
                    {
                        isValid = false;
                        msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
                    }
                    // ตรวจสอบค่า folder
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
                    {
                        isValid = false;
                        msg_error = "Invalid folder.";
                    }
                    if (isValid)
                    {
                        // สร้างพาธของโฟลเดอร์ที่อนุญาต
                        string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
                        if (isValid && !Directory.Exists(allowedDirectory))
                        {
                            isValid = false;
                            msg_error = "Folder directory not found.";
                        }

                        // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
                        string sourceFile = $"{safeFileName}";
                        if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
                                        string.IsNullOrWhiteSpace(sourceFile) ||
                                        sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                                        sourceFile.Contains("..") || sourceFile.Contains("\\")))
                        {
                            isValid = false;
                            msg_error = "Invalid fileName.";
                        }

                        // สร้างพาธเต็มของไฟล์
                        if (isValid)
                        {
                            fullPath = Path.Combine(allowedDirectory, sourceFile);
                            fullPath = Path.GetFullPath(fullPath);

                            // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
                            if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
                            {
                                isValid = false;
                                msg_error = "File is outside the allowed directory.";
                            }
                        }

                        // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        if (isValid)
                        {
                            FileInfo template = new FileInfo(fullPath);
                            if (!template.Exists)
                            {
                                isValid = false;
                                msg_error = "File not found.";
                            }

                            // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
                            if (isValid && template.IsReadOnly)
                            {
                                try
                                {
                                    template.IsReadOnly = false;
                                }
                                catch (Exception ex)
                                {
                                    isValid = false;
                                    msg_error = $"Failed to modify file attributes: {ex.Message}";
                                }
                            }
                        }

                        // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
                        if (isValid)
                        {
                            try
                            {
                                using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                                {
                                    // สามารถเปิดไฟล์ได้
                                }
                            }
                            catch (UnauthorizedAccessException)
                            {
                                isValid = false;
                                msg_error = "Access to the file is denied.";
                            }
                        }
                    }
                    #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


                    // หากทุกอย่างผ่านการตรวจสอบ
                    if (isValid)
                    {
                        // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
                        file_fullpath_name = fullPath;
                    }
                    else { file_fullpath_name = ""; }
                }
                else { file_fullpath_name = ""; }

                #endregion call function export full report pdf

                string url = "";
                string url_home_task = "";
                string step_text = "Outstanding Action Notification";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";
                string date_now = DateTime.Now.ToString("dd/MMM/yyyy");

                DataTable dt = new DataTable();
                DataTable dtUser = new DataTable();
                string mail_admin_group = get_mail_admin_group();

                ClassNoti classNoti = new ClassNoti();
                dtUser = classNoti.DataDailyByActionRequired_TeammMember("", seq, sub_software, true);
                dt = classNoti.DataDailyByActionRequired_TeammMember("", seq, sub_software, false);

                #region mail to
                s_mail_cc = mail_admin_group;

                string msg = "";
                if (dt?.Rows.Count > 0)
                {
                    for (int i = 0; i < dtUser?.Rows.Count; i++)
                    {
                        if (i > 0) { s_mail_to += ";"; }
                        s_mail_to += dtUser.Rows[i]["user_email"]?.ToString() ?? "";
                    }

                    if (true)
                    {
                        if (string.IsNullOrEmpty(url_home_task))
                        {
                            if (true)
                            {

                                string pha_no = "";

                                string plainText = $"seq={seq}&pha_no={pha_no}";
                                string cipherText = EncryptString(plainText);
                                url_home_task = $"{server_url_home_task}{cipherText}";
                            }
                        }

                        string s_subject = $"EPHA {step_text.ToUpper()}_{date_now}";

                        StringBuilder s_body = new StringBuilder();
                        s_body.Append("<html><body><font face='tahoma' size='2'>");
                        s_body.AppendFormat("Dear {0},", "All");
                        s_body.Append(@"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.");
                        s_body.Append(@"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please respond within five working days.</small>");

                        s_body.Append("</font></body></html>");

                        sendEmailModel data = new sendEmailModel
                        {
                            mail_subject = s_subject,
                            mail_body = s_body.ToString(),
                            mail_to = s_mail_to,
                            mail_cc = s_mail_cc,
                            mail_from = s_mail_from
                        };

                        if (!string.IsNullOrEmpty(file_fullpath_name) && File.Exists(file_fullpath_name))
                        {
                            data.mail_attachments = file_fullpath_name;
                        }

                        msg = sendMail(data);
                    }
                }
                #endregion mail to

                return msg;
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }

        public string MailNotificationApproverTA2eMOC(string seq, string sub_software)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            if (string.IsNullOrEmpty(sub_software)) { return "Invalid Sub Software."; }

            try
            {

                // ตรวจสอบค่าของ sub_software ด้วย whitelist
                var allowedSubSoftware = new HashSet<string> { "hazop", "jsea", "whatif" };
                if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                {
                    return ("Invalid sub_software value.");
                }

                #region call function export full report pdf  
                var sameRamTypeSoftwares = new HashSet<string> { "hazop", "jsea", "whatif" };
                string _ram_type = sameRamTypeSoftwares.Contains(sub_software) ? "5" : "";

                ReportModel param = new ReportModel
                {
                    seq = seq,
                    export_type = "pdf",
                    sub_software = sub_software,
                    user_name = "",
                    ram_type = _ram_type
                };
                ClassExcel clsExcel = new ClassExcel();
                string file_fullpath_name = "";
                file_fullpath_name = clsExcel.export_full_report(param, true);
                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = sub_software ?? "";
                    string msg_error = "";

                    //string folder = sub_software;
                    string fullPath = "";

                    #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

                    bool isValid = true;

                    // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
                    if (!Path.IsPathRooted(file_fullpath_def))
                    {
                        isValid = false;
                        msg_error = "File path is not an absolute path.";
                    }

                    // ตรวจสอบชื่อไฟล์
                    string safeFileName = Path.GetFileName(file_fullpath_def);
                    if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
                    {
                        isValid = false;
                        msg_error = "Invalid or potentially dangerous file name.";
                    }

                    // ตรวจสอบอักขระที่อนุญาต
                    char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                        .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                    // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                    string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
                    string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                    if (isValid && !allowedExtensions.Contains(fileExtension))
                    {
                        isValid = false;
                        msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
                    }
                    // ตรวจสอบค่า folder
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
                    {
                        isValid = false;
                        msg_error = "Invalid folder.";
                    }
                    if (isValid)
                    {
                        // สร้างพาธของโฟลเดอร์ที่อนุญาต
                        string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
                        if (isValid && !Directory.Exists(allowedDirectory))
                        {
                            isValid = false;
                            msg_error = "Folder directory not found.";
                        }

                        // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
                        string sourceFile = $"{safeFileName}";
                        if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
                                        string.IsNullOrWhiteSpace(sourceFile) ||
                                        sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                                        sourceFile.Contains("..") || sourceFile.Contains("\\")))
                        {
                            isValid = false;
                            msg_error = "Invalid fileName.";
                        }

                        // สร้างพาธเต็มของไฟล์
                        if (isValid)
                        {
                            fullPath = Path.Combine(allowedDirectory, sourceFile);
                            fullPath = Path.GetFullPath(fullPath);

                            // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
                            if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
                            {
                                isValid = false;
                                msg_error = "File is outside the allowed directory.";
                            }
                        }

                        // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        if (isValid)
                        {
                            FileInfo template = new FileInfo(fullPath);
                            if (!template.Exists)
                            {
                                isValid = false;
                                msg_error = "File not found.";
                            }

                            // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
                            if (isValid && template.IsReadOnly)
                            {
                                try
                                {
                                    template.IsReadOnly = false;
                                }
                                catch (Exception ex)
                                {
                                    isValid = false;
                                    msg_error = $"Failed to modify file attributes: {ex.Message}";
                                }
                            }
                        }

                        // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
                        if (isValid)
                        {
                            try
                            {
                                using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                                {
                                    // สามารถเปิดไฟล์ได้
                                }
                            }
                            catch (UnauthorizedAccessException)
                            {
                                isValid = false;
                                msg_error = "Access to the file is denied.";
                            }
                        }
                    }
                    #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


                    // หากทุกอย่างผ่านการตรวจสอบ
                    if (isValid)
                    {
                        // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
                        file_fullpath_name = fullPath;
                    }
                    else { file_fullpath_name = ""; }
                }
                else { file_fullpath_name = ""; }

                #endregion call function export full report pdf  

                string url = "";
                string url_home_task = "";
                string step_text = "TA2 Review and Approve for MOC";

                string to_displayname = "All";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";
                string date_now = DateTime.Now.ToString("dd/MMM/yyyy");

                DataTable dt = new DataTable();
                DataTable dtUser = new DataTable();
                string mail_admin_group = get_mail_admin_group();

                ClassNoti classNoti = new ClassNoti();
                //dtUser = classNoti.DataDailyByActionRequired_TeammMember_TA2eMOC("", seq, sub_software, true);
                //dt = classNoti.DataDailyByActionRequired_TeammMember_TA2eMOC("", seq, sub_software, false);
                List<SqlParameter> parameters = new List<SqlParameter>();
                sqlstr = "SELECT DISTINCT a.user_name, a.user_displayname, a.user_email FROM VW_EPHA_ACTION_TA2EMOC a where a.seq = @seq ORDER BY a.user_name";
                parameters = new List<SqlParameter>();
                if (!string.IsNullOrEmpty(seq))
                {
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                }
                else
                {
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = "-1" });
                }
                //dtUser = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                try
                {
                    _conn = new ClassConnectionDb(); _conn.OpenConnection(); try
                    {
                        var command = _conn.conn.CreateCommand();
                        command.CommandText = sqlstr;
                        command.Parameters.Add(parameters);
                        dtUser = new DataTable();
                        dtUser = _conn.ExecuteAdapter(command).Tables[0];
                        dtUser.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }

                sqlstr = "SELECT DISTINCT a.* FROM VW_EPHA_ACTION_TA2EMOC a  where a.seq = @seq ORDER BY a.user_name, a.action_sort, a.document_number, a.rev";
                parameters = new List<SqlParameter>();
                if (!string.IsNullOrEmpty(seq))
                {
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                }
                else
                {
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = "-1" });
                }

                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                try
                {
                    _conn = new ClassConnectionDb(); _conn.OpenConnection(); 
                    try
                    {
                        var command = _conn.conn.CreateCommand();
                        command.CommandText = sqlstr;
                        command.Parameters.Add(parameters);
                        dt = new DataTable();
                        dt = _conn.ExecuteAdapter(command).Tables[0];
                        dt.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }

                #region mail to 
                s_mail_cc = mail_admin_group;

                string msg = "";
                if (dt?.Rows.Count > 0)
                {
                    for (int i = 0; i < dtUser?.Rows.Count; i++)
                    {
                        if (i > 0) { s_mail_to += ";"; }
                        s_mail_to += dtUser.Rows[i]["user_email"]?.ToString() ?? "";

                        if (string.IsNullOrEmpty(url_home_task))
                        {
                            if (true)
                            {

                                string pha_no = ""; // (dtUser.Rows[i]["document_number"]?.ToString() ?? "");

                                string plainText = $"seq={seq}&pha_no={pha_no}";
                                string cipherText = EncryptString(plainText);
                                url_home_task = $"{server_url_home_task}{cipherText}";
                            }
                        }
                    }

                    if (true)
                    {
                        string s_subject = $"EPHA {step_text.ToUpper()}_{date_now}";

                        StringBuilder s_body = new StringBuilder();
                        s_body.Append("<html><body><font face='tahoma' size='2'>");
                        s_body.AppendFormat("Dear {0},", to_displayname);
                        s_body.Append(@"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.");
                        s_body.Append(@"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please respond by replying to this email within five working days prior to auto proceed to the next step.</small>");

                        // Table header
                        s_body.Append(@"<br/>
                    <table style='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   
                        <thead>    
                            <tr>
                                <td style='padding: 15px;'>Task</td>
                                <td style='padding: 15px;'>PHA Type</td>
                                <td style='padding: 15px;'>Action Required</td>
                                <td style='padding: 15px;'>Document Number</td>
                                <td style='padding: 15px;'>Document Title</td>
                                <td style='padding: 15px;'>Rev.</td>
                                <td style='padding: 15px;'>Originator</td>
                                <td style='padding: 15px;'>Received</td>
                                <td style='padding: 15px;'>Due Date</td>
                                <td style='padding: 15px;'>Remaining</td>
                                <td style='padding: 15px;'>Consolidator</td>
                            </tr>
                        </thead>
                        <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'>");

                        int iNo = 1;
                        foreach (DataRow row in dt.Rows)
                        {
                            string doc_no = row["document_number"]?.ToString() ?? "";
                            string background_color = "white";
                            string font_color = "black";
                            bool action_status_close = (row["remaining"]?.ToString() ?? "").ToLower() == "closed";

                            int remainingDays = Convert.ToInt32(row["remaining"]?.ToString() ?? "0");
                            if (remainingDays > 3)
                            {
                                background_color = "green"; font_color = "red";
                            }
                            else if (remainingDays <= 3 && remainingDays > 0 && !action_status_close)
                            {
                                background_color = "yellow";
                            }
                            else if (remainingDays <= 0 && !action_status_close)
                            {
                                background_color = "red"; font_color = "white";
                            }

                            #region Generate URL 
                            url = "";
                            if (true)
                            {
                                string plainText = $"seq={seq}&pha_no={doc_no}&step=3";
                                string cipherText = EncryptString(plainText);
                                url = $"{server_url_by_action(sub_software)}{cipherText}";
                            }
                            #endregion Generate URL 

                            // Add data to the table
                            s_body.Append("<tr>");
                            s_body.AppendFormat("<td style='padding: 15px;'>{0}</td>", iNo);
                            s_body.AppendFormat("<td style='padding: 15px;'>{0}</td>", row["pha_type"]);
                            s_body.AppendFormat("<td style='padding: 15px;'><a href='{0}'>{1}</a></td>", url, step_text);
                            s_body.AppendFormat("<td style='padding: 15px;'><a href='{0}'>{1}</a></td>", url, row["document_number"]);
                            s_body.AppendFormat("<td style='padding: 15px;'>{0}</td>", row["document_title"]);
                            s_body.AppendFormat("<td style='padding: 15px;'>{0}</td>", row["rev"]);
                            s_body.AppendFormat("<td style='padding: 15px;'>{0}</td>", row["originator"]);
                            s_body.AppendFormat("<td style='padding: 15px;'>{0}</td>", row["receivesd"]);
                            s_body.AppendFormat("<td style='padding: 15px;'>{0}</td>", row["due_date"]);
                            s_body.AppendFormat("<td style='padding: 15px; background-color:{0};color:{1};'>{2}</td>", background_color, font_color, remainingDays);
                            s_body.AppendFormat("<td style='padding: 15px;'>{0}</td>", row["consolidator"]);
                            s_body.Append("</tr>");
                            iNo++;
                        }

                        s_body.Append("</tbody></table>");

                        // Add disclaimer and send email
                        s_body.Append("<br/><br/>DISCLAIMER: ...");

                        sendEmailModel data = new sendEmailModel
                        {
                            mail_subject = s_subject,
                            mail_body = s_body.ToString(),
                            mail_to = s_mail_to,
                            mail_cc = s_mail_cc,
                            mail_from = s_mail_from
                        };

                        if (!string.IsNullOrEmpty(file_fullpath_name) && File.Exists(file_fullpath_name))
                        {
                            data.mail_attachments = file_fullpath_name;
                        }

                        msg = sendMail(data);
                    }
                }
                #endregion mail to 

                return msg;
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }

        public string MailNotificationApproverQMTSReviewer(string seq, string sub_software, string pha_status_def)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            if (string.IsNullOrEmpty(sub_software)) { return "Invalid Sub Software."; }
            if (string.IsNullOrEmpty(pha_status_def)) { return "Invalid Status."; }

            try
            {

                // Define a whitelist of allowed sub_software values
                var allowedSubSoftware = new HashSet<string> { "hazop", "jsea", "whatif", "hra" };

                // Check if sub_software is valid
                if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                {
                    return ("Invalid sub_software value");
                }

                var sameRamTypeSoftwares = new HashSet<string> { "hazop", "jsea", "whatif" };
                string _ram_type = sameRamTypeSoftwares.Contains(sub_software) ? "5" : "";

                ReportModel param = new ReportModel
                {
                    seq = seq,
                    export_type = "pdf",
                    sub_software = sub_software,
                    user_name = "",
                    ram_type = _ram_type
                };

                #region call function export full report pdf  
                ClassExcel clsExcel = new ClassExcel();
                string file_fullpath_name = "";
                file_fullpath_name = clsExcel.export_full_report(param, true);

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = sub_software ?? "";
                    string msg_error = "";

                    //string folder = sub_software;
                    string fullPath = "";

                    #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

                    bool isValid = true;

                    // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
                    if (!Path.IsPathRooted(file_fullpath_def))
                    {
                        isValid = false;
                        msg_error = "File path is not an absolute path.";
                    }

                    // ตรวจสอบชื่อไฟล์
                    string safeFileName = Path.GetFileName(file_fullpath_def);
                    if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
                    {
                        isValid = false;
                        msg_error = "Invalid or potentially dangerous file name.";
                    }

                    // ตรวจสอบอักขระที่อนุญาต
                    char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                        .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                    // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                    string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
                    string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                    if (isValid && !allowedExtensions.Contains(fileExtension))
                    {
                        isValid = false;
                        msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
                    }
                    // ตรวจสอบค่า folder
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
                    {
                        isValid = false;
                        msg_error = "Invalid folder.";
                    }
                    if (isValid)
                    {
                        // สร้างพาธของโฟลเดอร์ที่อนุญาต
                        string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
                        if (isValid && !Directory.Exists(allowedDirectory))
                        {
                            isValid = false;
                            msg_error = "Folder directory not found.";
                        }

                        // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
                        string sourceFile = $"{safeFileName}";
                        if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
                                        string.IsNullOrWhiteSpace(sourceFile) ||
                                        sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                                        sourceFile.Contains("..") || sourceFile.Contains("\\")))
                        {
                            isValid = false;
                            msg_error = "Invalid fileName.";
                        }

                        // สร้างพาธเต็มของไฟล์
                        if (isValid)
                        {
                            fullPath = Path.Combine(allowedDirectory, sourceFile);
                            fullPath = Path.GetFullPath(fullPath);

                            // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
                            if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
                            {
                                isValid = false;
                                msg_error = "File is outside the allowed directory.";
                            }
                        }

                        // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        if (isValid)
                        {
                            FileInfo template = new FileInfo(fullPath);
                            if (!template.Exists)
                            {
                                isValid = false;
                                msg_error = "File not found.";
                            }

                            // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
                            if (isValid && template.IsReadOnly)
                            {
                                try
                                {
                                    template.IsReadOnly = false;
                                }
                                catch (Exception ex)
                                {
                                    isValid = false;
                                    msg_error = $"Failed to modify file attributes: {ex.Message}";
                                }
                            }
                        }

                        // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
                        if (isValid)
                        {
                            try
                            {
                                using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                                {
                                    // สามารถเปิดไฟล์ได้
                                }
                            }
                            catch (UnauthorizedAccessException)
                            {
                                isValid = false;
                                msg_error = "Access to the file is denied.";
                            }
                        }
                    }
                    #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


                    // หากทุกอย่างผ่านการตรวจสอบ
                    if (isValid)
                    {
                        // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
                        file_fullpath_name = fullPath;
                    }
                    else { file_fullpath_name = ""; }
                }
                else { file_fullpath_name = ""; }

                #endregion call function export full report pdf 

                // Continue with the rest of the email generation and sending process
                string doc_no = "";
                string doc_name = "";
                string reference_moc = "";

                string url_home_task = "";
                string url = "";
                string url_approver = "";
                string url_reject_no_comment = "";
                string url_reject_comment = "";

                string pha_sub_software = "";
                string meeting_date = "";
                string meeting_time = "";

                string to_displayname = "All";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                string note_to_approver = "";

                DataTable dt = new DataTable();

                string sqlstr = "";
                sqlstr = @"select distinct a.* from VW_EPHA_DATA_TO_EMAIL a
                    where a.request_approver = 1 and a.id = @seq and isnull(a.approver_action_type, 0) < 2
                    order by a.no_approver";

                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "" });
                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
                try
                {
                    _conn = new ClassConnectionDb(); _conn.OpenConnection(); try
                    {
                        var command = _conn.conn.CreateCommand();
                        command.CommandText = sqlstr;
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(param);
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

                #region mail to
                if (dt?.Rows.Count > 0)
                {
                    to_displayname = dt.Rows[0]["user_displayname"]?.ToString() ?? "";
                    doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                    doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                    reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                    pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString()?.ToUpper() ?? "";
                    meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                    meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        if (i > 0) { s_mail_to += ";"; }
                        s_mail_to += dt.Rows[i]["user_email"]?.ToString() ?? "";
                    }

                    note_to_approver = dt.Rows[0]["note_to_approver"]?.ToString() ?? "";
                }
                #endregion mail to

                #region mail cc 
                if (dt?.Rows.Count > 0)
                {
                    s_mail_cc = dt.Rows[0]["request_email"]?.ToString() ?? "";
                }
                #endregion mail cc

                #region URL Encryption
                if (true)
                {
                    string pha_no = doc_no;
                    string plainText = $"seq={seq}&pha_no={pha_no}";
                    string cipherText = EncryptString(plainText);
                    url_home_task = $"{server_url_home_task}{cipherText}";
                }

                if (true)
                {

                    string plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject";
                    string cipherText = EncryptString(plainText);
                    url = $"{server_url_by_action(sub_software)}{cipherText}";

                    url_reject_comment = url;

                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject_no_comment";
                    cipherText = EncryptString(plainText);
                    url_reject_no_comment = $"{server_url_by_action(sub_software)}{cipherText}";

                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=approve";
                    cipherText = EncryptString(plainText);
                    url_approver = $"{server_url_by_action(sub_software)}{cipherText}";
                }
                #endregion URL Encryption

                StringBuilder s_body = new StringBuilder();
                s_body.Append("<html><body><font face='tahoma' size='2'>");
                s_body.AppendFormat("Dear {0},", to_displayname);

                s_body.Append("<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
                s_body.AppendFormat("Please approve the {0}, PHA No.{1} as details mentioned below,", pha_sub_software.ToUpper(), doc_no);

                s_body.AppendFormat("<br/><br/><b>Project Name</b> : {0}", doc_name);
                s_body.AppendFormat("<br/><b>Date</b> : {0}", meeting_date);
                s_body.AppendFormat("<br/><b>Time</b> : {0}", meeting_time);

                if (!string.IsNullOrEmpty(reference_moc)) { s_body.AppendFormat("<br/><b>Reference MOC</b> : {0}", reference_moc); }
                if (!string.IsNullOrEmpty(note_to_approver)) { s_body.AppendFormat("<br/><br/><b>Note to Approver</b> : {0}", note_to_approver); }

                s_body.Append("<br/><br/>");
                s_body.AppendFormat("More details of the study, <font color='red'; text-decoration: underline;><a href='{0}'> please click here</a></font>", url);

                s_body.Append("<br/><br/><b>Reply :</b>");
                s_body.AppendFormat("<br/><a style='border: none;background-color: #25b003; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block;' href='{0}'>Approve</a>", url_approver);
                s_body.AppendFormat("<a style='border: none;background-color: #d90476; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;' href='{0}'>Send back No Comment</a>", url_reject_no_comment);
                s_body.AppendFormat("<a style='border: none;background-color: #f64a8a; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;' href='{0}'>Send back with Comment</a>", url_reject_comment);

                s_body.Append("<br/><br/>DISCLAIMER:");
                s_body.Append("<br/>This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s) of Thai Oil Public Company Limited.");
                s_body.Append("</font></body></html>");

                sendEmailModel data = new sendEmailModel
                {
                    mail_subject = $"EPHA : {pha_sub_software.ToUpper()} Waiting QMTS Review",
                    mail_body = s_body.ToString(),
                    mail_to = s_mail_to,
                    mail_cc = s_mail_cc,
                    mail_from = s_mail_from
                };

                if (!string.IsNullOrEmpty(file_fullpath_name) && File.Exists(file_fullpath_name))
                {
                    data.mail_attachments = file_fullpath_name;
                }

                return sendMail(data);
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }
        public string MailNotificationApproverSafetyReviewer(string seq, string sub_software, string pha_status_def)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            if (string.IsNullOrEmpty(sub_software)) { return "Invalid Sub Software."; }
            if (string.IsNullOrEmpty(pha_status_def)) { return "Invalid Status."; }

            try
            {

                // Define a whitelist of allowed sub_software values
                var allowedSubSoftware = new HashSet<string> { "jsea", "whatif", "hra", "hazop" };

                // Check if sub_software is valid
                if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                {
                    return ("Invalid sub_software value");
                }

                if (!Regex.IsMatch(sub_software, @"^[a-zA-Z0-9_]+$"))
                {
                    return ("Invalid sub_software value.");
                }

                #region call function export full report pdf  
                var sameRamTypeSoftwares = new HashSet<string> { "hazop", "jsea", "whatif" };
                string _ram_type = sameRamTypeSoftwares.Contains(sub_software) ? "5" : "";

                ReportModel param = new ReportModel
                {
                    seq = seq,
                    export_type = "pdf",
                    sub_software = sub_software,
                    user_name = "",
                    ram_type = _ram_type
                };
                ClassExcel clsExcel = new ClassExcel();
                string file_fullpath_name = "";
                file_fullpath_name = clsExcel.export_full_report(param, true);

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = sub_software ?? "";
                    string msg_error = "";

                    //string folder = sub_software;
                    string fullPath = "";

                    #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

                    bool isValid = true;

                    // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
                    if (!Path.IsPathRooted(file_fullpath_def))
                    {
                        isValid = false;
                        msg_error = "File path is not an absolute path.";
                    }

                    // ตรวจสอบชื่อไฟล์
                    string safeFileName = Path.GetFileName(file_fullpath_def);
                    if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
                    {
                        isValid = false;
                        msg_error = "Invalid or potentially dangerous file name.";
                    }

                    // ตรวจสอบอักขระที่อนุญาต
                    char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                        .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                    // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                    string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
                    string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                    if (isValid && !allowedExtensions.Contains(fileExtension))
                    {
                        isValid = false;
                        msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
                    }
                    // ตรวจสอบค่า folder
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
                    {
                        isValid = false;
                        msg_error = "Invalid folder.";
                    }
                    if (isValid)
                    {
                        // สร้างพาธของโฟลเดอร์ที่อนุญาต
                        string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
                        if (isValid && !Directory.Exists(allowedDirectory))
                        {
                            isValid = false;
                            msg_error = "Folder directory not found.";
                        }

                        // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
                        string sourceFile = $"{safeFileName}";
                        if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
                                        string.IsNullOrWhiteSpace(sourceFile) ||
                                        sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                                        sourceFile.Contains("..") || sourceFile.Contains("\\")))
                        {
                            isValid = false;
                            msg_error = "Invalid fileName.";
                        }

                        // สร้างพาธเต็มของไฟล์
                        if (isValid)
                        {
                            fullPath = Path.Combine(allowedDirectory, sourceFile);
                            fullPath = Path.GetFullPath(fullPath);

                            // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
                            if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
                            {
                                isValid = false;
                                msg_error = "File is outside the allowed directory.";
                            }
                        }

                        // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        if (isValid)
                        {
                            FileInfo template = new FileInfo(fullPath);
                            if (!template.Exists)
                            {
                                isValid = false;
                                msg_error = "File not found.";
                            }

                            // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
                            if (isValid && template.IsReadOnly)
                            {
                                try
                                {
                                    template.IsReadOnly = false;
                                }
                                catch (Exception ex)
                                {
                                    isValid = false;
                                    msg_error = $"Failed to modify file attributes: {ex.Message}";
                                }
                            }
                        }

                        // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
                        if (isValid)
                        {
                            try
                            {
                                using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                                {
                                    // สามารถเปิดไฟล์ได้
                                }
                            }
                            catch (UnauthorizedAccessException)
                            {
                                isValid = false;
                                msg_error = "Access to the file is denied.";
                            }
                        }
                    }
                    #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


                    // หากทุกอย่างผ่านการตรวจสอบ
                    if (isValid)
                    {
                        // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
                        file_fullpath_name = fullPath;
                    }
                    else { file_fullpath_name = ""; }
                }
                else { file_fullpath_name = ""; }

                #endregion call function export full report pdf 

                string doc_no = "";
                string doc_name = "";
                string reference_moc = "";

                string url_home_task = "";
                string url = "";
                string url_approver = "";
                string url_reject_no_comment = "";
                string url_reject_comment = "";
                string pha_sub_software = "";
                string meeting_date = "";
                string meeting_time = "";

                string to_displayname = "All";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                string note_to_approver = "";

                DataTable dt = new DataTable();
                string sqlstr = "";
                sqlstr = @"SELECT DISTINCT h.pha_status, h.pha_no as pha_no, g.pha_request_name as pha_name, empre.user_email as request_email,
                                       ta2.no, FORMAT(s.meeting_date, 'dd MMM yyyy') as meeting_date,
                                       REPLACE(s.meeting_start_time,'1/1/1970 ','') + ' - ' + REPLACE(s.meeting_end_time,'1/1/1970 ','') as meeting_time,
                                       emp.user_displayname, emp.user_email, g.reference_moc, LOWER(h.pha_sub_software) as pha_sub_software,
                                       h.approver_user_name, h.approve_action_type, h.approve_status, h.approve_comment, s.note_to_approver
                                FROM epha_t_header h
                                INNER JOIN EPHA_T_GENERAL g ON LOWER(h.id) = LOWER(g.id_pha)
                                INNER JOIN EPHA_T_SESSION s ON LOWER(h.id) = LOWER(s.id_pha)
                                INNER JOIN EPHA_T_APPROVER ta2 ON LOWER(h.id) = LOWER(ta2.id_pha) AND s.seq = ta2.id_session
                                INNER JOIN VW_EPHA_PERSON_DETAILS emp ON LOWER(ta2.user_name) = LOWER(emp.user_name)
                                INNER JOIN VW_EPHA_PERSON_DETAILS empre ON LOWER(h.pha_request_by) = LOWER(empre.user_name)
                                INNER JOIN (SELECT MAX(id) as id_session, id_pha FROM EPHA_T_SESSION GROUP BY id_pha) s1 
                                            ON h.id = s1.id_pha AND s.id = s1.id_session AND s.id_pha = s1.id_pha
                                WHERE h.request_approver = 1 AND h.id = @seq AND ISNULL(ta2.approver_action_type, 0) < 2
                                ORDER BY CONVERT(INT, ta2.no)";

                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });

                //dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);
                #region Execute to Datable
                try
                {
                    _conn = new ClassConnectionDb(); _conn.OpenConnection();
                    try
                    {
                        var command = _conn.conn.CreateCommand();
                        command.CommandText = sqlstr;
                        //command.Parameters.Add(":costcenter", costcenter); 
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                        dt = new DataTable();
                        dt = _conn.ExecuteAdapter(command).Tables[0];
                        //dt.TableName = "Table1";
                        dt.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                #region mail to
                if (dt?.Rows.Count > 0)
                {
                    to_displayname = dt.Rows[0]["user_displayname"]?.ToString() ?? "";
                    doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                    doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                    reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                    pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString()?.ToUpper() ?? "";
                    meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                    meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";
                    note_to_approver = dt.Rows[0]["note_to_approver"]?.ToString() ?? "";

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (i > 0) { s_mail_to += ";"; }
                        s_mail_to += dt.Rows[i]["user_email"]?.ToString() ?? "";
                    }
                }
                #endregion mail to

                #region mail cc 
                if (dt?.Rows.Count > 0)
                {
                    s_mail_cc = dt.Rows[0]["request_email"]?.ToString() ?? "";
                }
                #endregion mail cc

                #region url generation

                string pha_no = doc_no;
                if (!string.IsNullOrEmpty(pha_no))
                {
                    string plainText = $"seq={seq}&pha_no={pha_no}";
                    string cipherText = EncryptString(plainText);
                    url_home_task = $"{server_url_home_task}{cipherText}";

                    // Generating URLs for different actions
                    string basePlainText = $"seq={seq}&pha_no={doc_no}&step=4";
                    cipherText = EncryptString($"{basePlainText}&approver_type=reject");
                    url_reject_comment = $"{server_url_by_action(sub_software)}{cipherText}";

                    cipherText = EncryptString($"{basePlainText}&approver_type=reject_no_comment");
                    url_reject_no_comment = $"{server_url_by_action(sub_software)}{cipherText}";

                    cipherText = EncryptString($"{basePlainText}&approver_type=approve");
                    url_approver = $"{server_url_by_action(sub_software)}{cipherText}";
                    #endregion url generation

                    // Preparing email body
                    StringBuilder s_body = new StringBuilder();
                    s_body.Append("<html><body><font face='tahoma' size='2'>");
                    s_body.AppendFormat("Dear {0},", to_displayname);
                    s_body.AppendFormat("<br/><br/>Please approve the {0}, PHA No.{1} as detailed below:", pha_sub_software, doc_no);
                    s_body.AppendFormat("<br/><br/><b>Project Name</b>: {0}", doc_name);
                    s_body.AppendFormat("<br/><b>Date</b>: {0}", meeting_date);
                    s_body.AppendFormat("<br/><b>Time</b>: {0}", meeting_time);
                    if (!string.IsNullOrEmpty(reference_moc)) s_body.AppendFormat("<br/><b>Reference MOC</b>: {0}", reference_moc);
                    if (!string.IsNullOrEmpty(note_to_approver)) s_body.AppendFormat("<br/><br/><b>Note to Approver</b>: {0}", note_to_approver);

                    s_body.AppendFormat("<br/><br/>More details of the study, <a href='{0}' style='color: red;'>please click here</a>", url);

                    s_body.Append("<br/><br/><b>Reply :</b>");
                    s_body.AppendFormat("<br/><a href='{0}' style='background-color: #25b003; padding: 14px 28px;'>Approve</a>", url_approver);
                    s_body.AppendFormat("<a href='{0}' style='background-color: #d90476; padding: 14px 28px; margin-left: 30px;'>Send back No Comment</a>", url_reject_no_comment);
                    s_body.AppendFormat("<a href='{0}' style='background-color: #f64a8a; padding: 14px 28px; margin-left: 30px;'>Send back with Comment</a>", url_reject_comment);

                    s_body.Append("<br/><br/>DISCLAIMER:");
                    s_body.Append("<br/>This email message is intended only for the personal use of the recipient(s) named above...");
                    s_body.Append("</font></body></html>");

                    sendEmailModel data = new sendEmailModel
                    {
                        mail_subject = $"EPHA : {pha_sub_software} Waiting Safety Reviewer Review",
                        mail_body = s_body.ToString(),
                        mail_to = s_mail_to,
                        mail_cc = s_mail_cc,
                        mail_from = s_mail_from
                    };

                    if (!string.IsNullOrEmpty(file_fullpath_name) && File.Exists(file_fullpath_name))
                    {
                        data.mail_attachments = file_fullpath_name;
                    }

                    return sendMail(data);
                }
                else
                {
                    return "";
                }

            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }

        public string MailNotificationApproverTA2(string seq, string sub_software, string pha_status_def, string approver_user_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            if (string.IsNullOrEmpty(sub_software)) { return "Invalid Sub Software."; }
            if (string.IsNullOrEmpty(pha_status_def)) { return "Invalid Status."; }
            if (string.IsNullOrEmpty(approver_user_name)) { return "Invalid Approver User Name."; }

            try
            {
                #region call function export full report pdf  
                var sameRamTypeSoftwares = new HashSet<string> { "hazop", "jsea", "whatif" };
                string _ram_type = sameRamTypeSoftwares.Contains(sub_software) ? "5" : "";

                ReportModel param = new ReportModel
                {
                    seq = seq,
                    export_type = "pdf",
                    sub_software = sub_software,
                    user_name = approver_user_name,
                    ram_type = _ram_type
                };
                ClassExcel clsExcel = new ClassExcel();
                string file_fullpath_name = "";
                file_fullpath_name = clsExcel.export_full_report(param, true);
                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = sub_software ?? "";
                    string msg_error = "";

                    //string folder = sub_software;
                    string fullPath = "";

                    #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

                    bool isValid = true;

                    // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
                    if (!Path.IsPathRooted(file_fullpath_def))
                    {
                        isValid = false;
                        msg_error = "File path is not an absolute path.";
                    }

                    // ตรวจสอบชื่อไฟล์
                    string safeFileName = Path.GetFileName(file_fullpath_def);
                    if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
                    {
                        isValid = false;
                        msg_error = "Invalid or potentially dangerous file name.";
                    }

                    // ตรวจสอบอักขระที่อนุญาต
                    char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                        .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                    // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                    string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
                    string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                    if (isValid && !allowedExtensions.Contains(fileExtension))
                    {
                        isValid = false;
                        msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
                    }
                    // ตรวจสอบค่า folder
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
                    {
                        isValid = false;
                        msg_error = "Invalid folder.";
                    }
                    if (isValid)
                    {
                        // สร้างพาธของโฟลเดอร์ที่อนุญาต
                        string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
                        if (isValid && !Directory.Exists(allowedDirectory))
                        {
                            isValid = false;
                            msg_error = "Folder directory not found.";
                        }

                        // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
                        string sourceFile = $"{safeFileName}";
                        if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
                                        string.IsNullOrWhiteSpace(sourceFile) ||
                                        sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                                        sourceFile.Contains("..") || sourceFile.Contains("\\")))
                        {
                            isValid = false;
                            msg_error = "Invalid fileName.";
                        }

                        // สร้างพาธเต็มของไฟล์
                        if (isValid)
                        {
                            fullPath = Path.Combine(allowedDirectory, sourceFile);
                            fullPath = Path.GetFullPath(fullPath);

                            // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
                            if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
                            {
                                isValid = false;
                                msg_error = "File is outside the allowed directory.";
                            }
                        }

                        // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        if (isValid)
                        {
                            FileInfo template = new FileInfo(fullPath);
                            if (!template.Exists)
                            {
                                isValid = false;
                                msg_error = "File not found.";
                            }

                            // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
                            if (isValid && template.IsReadOnly)
                            {
                                try
                                {
                                    template.IsReadOnly = false;
                                }
                                catch (Exception ex)
                                {
                                    isValid = false;
                                    msg_error = $"Failed to modify file attributes: {ex.Message}";
                                }
                            }
                        }

                        // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
                        if (isValid)
                        {
                            try
                            {
                                using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                                {
                                    // สามารถเปิดไฟล์ได้
                                }
                            }
                            catch (UnauthorizedAccessException)
                            {
                                isValid = false;
                                msg_error = "Access to the file is denied.";
                            }
                        }
                    }
                    #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


                    // หากทุกอย่างผ่านการตรวจสอบ
                    if (isValid)
                    {
                        // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
                        file_fullpath_name = fullPath;
                    }
                    else { file_fullpath_name = ""; }
                }
                else { file_fullpath_name = ""; }
                #endregion call function export full report pdf 

                string doc_no = "";
                string doc_name = "";
                string reference_moc = "";

                string url_home_task = "";
                string url = "";
                string url_approver = "";
                string url_reject_no_comment = "";
                string url_reject_comment = "";

                string pha_sub_software = "";
                string meeting_date = "";
                string meeting_time = "";

                string to_displayname = "All";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                string note_to_approver = "";

                DataTable dtOwner = new DataTable();
                var parameters = new List<SqlParameter>();
                string sqlstr = QueryActionOwnerUpperTA2(seq, approver_user_name, sub_software, ref parameters);

                //dtOwner = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                #region Execute to Datable
                try
                {
                    _conn = new ClassConnectionDb(); _conn.OpenConnection();
                    try
                    {
                        var command = _conn.conn.CreateCommand();
                        command.CommandText = sqlstr;
                        //command.Parameters.Add(":costcenter", costcenter); 
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                        dtOwner = new DataTable();
                        dtOwner = _conn.ExecuteAdapter(command).Tables[0];
                        //dt.TableName = "Table1";
                        dtOwner.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = @" SELECT * from VW_EPHA_DATA_TO_EMAIL a WHERE a.request_approver = 1 and isnull(a.approver_action_type, 0) < 2 AND a.id = @seq";
                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "" });

                if (!string.IsNullOrEmpty(sub_software))
                {
                    sqlstr += " and lower(a.pha_sub_software) = lower(@sub_software) ";
                    parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });
                }
                if (!string.IsNullOrEmpty(approver_user_name))
                {
                    sqlstr += " AND LOWER(a.approver_user_name) = LOWER(@user_approver_active)";
                    parameters.Add(new SqlParameter("@user_approver_active", SqlDbType.VarChar, 100) { Value = approver_user_name.ToLower() });
                }
                sqlstr += " ORDER BY a.no_approver ";

                DataTable dt = new DataTable();
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
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                        dt = new DataTable();
                        dt = _conn.ExecuteAdapter(command).Tables[0];
                        //dt.TableName = "Table1";
                        dt.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (dt?.Rows.Count == 0) { return ""; }

                #region mail to
                if (dt?.Rows.Count > 0)
                {
                    to_displayname = dt.Rows[0]["user_displayname"]?.ToString() ?? "";

                    doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                    doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                    reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                    pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                    pha_sub_software = (pha_sub_software == "hazop" || pha_sub_software == "jsea" || pha_sub_software == "hra") ? pha_sub_software.ToUpper() : pha_sub_software;
                    meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                    meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        if (i > 0) { s_mail_to += ";"; }
                        s_mail_to += dt.Rows[i]["user_email"]?.ToString() ?? "";
                    }

                    note_to_approver = dt.Rows[0]["note_to_approver"]?.ToString() ?? "";
                }
                #endregion mail to

                #region mail cc 
                if (dt?.Rows.Count > 0)
                {
                    s_mail_cc = dt.Rows[0]["request_email"]?.ToString() ?? "";
                }
                #endregion mail cc


                string pha_no = doc_no;
                if (!string.IsNullOrEmpty(pha_no))
                {
                    string plainText = $"seq={seq}&pha_no={pha_no}";
                    string cipherText = EncryptString(plainText);
                    url_home_task = $"{server_url_home_task}{cipherText}";

                    #region url  

                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject";
                    cipherText = EncryptString(plainText);

                    string approver_url = server_url_by_action("approve");

                    url = $"{server_url_by_action(sub_software)}{cipherText}";

                    url_reject_comment = url;

                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject_no_comment";
                    cipherText = EncryptString(plainText);
                    url_reject_no_comment = $"{server_url_by_action(sub_software)}{cipherText}";

                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=approve";
                    cipherText = EncryptString(plainText);
                    url_approver = $"{server_url_by_action(sub_software)}{cipherText}";

                    #endregion url 

                    StringBuilder s_body = new StringBuilder();
                    s_body.Append("<html><body><font face='tahoma' size='2'>");
                    s_body.AppendFormat("Dear {0},", to_displayname);

                    s_body.Append("<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
                    s_body.AppendFormat("Please approve the {0}, PHA No.{1} as details mentioned below,", pha_sub_software.ToUpper(), doc_no);

                    s_body.AppendFormat("<br/><br/><b>Project Name</b> : {0}", doc_name);
                    s_body.AppendFormat("<br/><b>Date</b> : {0}", meeting_date);
                    s_body.AppendFormat("<br/><b>Time</b> : {0}", meeting_time);

                    if (!string.IsNullOrEmpty(reference_moc)) { s_body.AppendFormat("<br/><b>Reference MOC</b> : {0}", reference_moc); }
                    s_body.AppendFormat("<br/><br/><b>Note to Approver</b> : {0}", note_to_approver);

                    s_body.Append("<br/><br/>");
                    s_body.AppendFormat("More details of the study, <font color='red'; text-decoration: underline;><a href='{0}'> please click here</a></font>", url);

                    if (dtOwner?.Rows.Count > 0)
                    {
                        s_body.Append(@"<br/><br/>
                <table style ='border-collapse: collapse;font-family: Tahoma, Geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>
                    <thead>
                     <tr>
	                     <td style ='padding: 15px;' rowspan='2'>SUB-SOFTWARE</td>
	                     <td style ='padding: 15px;' rowspan='2'>PHA NO.</td>
	                     <td style ='padding: 15px;' rowspan='2'>RESPONDER</td>
	                     <td style ='padding: 15px; text-align: center;' colspan='3'>ITEMS STATUS</td> 
                     </tr>
                        <tr>
                            <td style ='padding: 15px;'>TOTAL</td>
                            <td style ='padding: 15px;'>OPEN</td>
                            <td style ='padding: 15px;'>CLOSE</td>		
                        </tr>
                    </thead> ");

                        s_body.Append("<tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'>");
                        for (int o = 0; o < dtOwner?.Rows.Count; o++)
                        {
                            s_body.Append("<tr>");
                            s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", sub_software.ToUpper());
                            s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dtOwner.Rows[o]["pha_no"]);
                            s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dtOwner.Rows[o]["user_displayname"]);
                            s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dtOwner.Rows[o]["total"]);
                            s_body.AppendFormat("<td style ='padding: 15px; color: red'>{0}</td>", dtOwner.Rows[o]["open"]);
                            s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dtOwner.Rows[o]["closed"]);
                            s_body.Append("</tr>");
                        }
                        s_body.Append("</tbody>");
                        s_body.Append("</table>");
                    }

                    s_body.Append("<br/><br/><b>Reply :</b>");
                    s_body.AppendFormat("<br/><a style='border: none;background-color: #25b003; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block;' href='{0}'>Approve</a>", url_approver);
                    s_body.AppendFormat("<a style='border: none;background-color: #d90476; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;' href='{0}'>Send back No Comment</a>", url_reject_no_comment);
                    s_body.AppendFormat("<a style='border: none;background-color: #f64a8a; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;' href='{0}'>Send back with Comment</a>", url_reject_comment);

                    s_body.Append("<br/><br/>DISCLAIMER:");
                    s_body.Append("<br/>");
                    s_body.Append(@"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.");
                    s_body.Append("</font></body></html>");

                    sendEmailModel data = new sendEmailModel
                    {
                        mail_subject = s_subject,
                        mail_body = s_body.ToString(),
                        mail_to = s_mail_to,
                        mail_cc = s_mail_cc,
                        mail_from = s_mail_from
                    };

                    if (!string.IsNullOrEmpty(file_fullpath_name) && File.Exists(file_fullpath_name))
                    {
                        data.mail_attachments = file_fullpath_name;
                    }

                    return sendMail(data);
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }

        public string MailApprovByApprover(string seq, string sub_software, string user_approver_active)
        {
            if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(sub_software) || string.IsNullOrEmpty(user_approver_active))
            {
                return "Invalid input parameters.";
            }

            try
            {
                // Define a whitelist of allowed sub_software values
                var allowedSubSoftware = new HashSet<string> { "hazop", "jsea", "whatif", "hra" };
                if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                {
                    return "Invalid sub_software value.";
                }

                string file_fullpath_name = get_document_file_approver(seq, user_approver_active);

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = sub_software ?? "";
                    string msg_error = "";

                    //string folder = sub_software;
                    string fullPath = "";

                    #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

                    bool isValid = true;

                    // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
                    if (!Path.IsPathRooted(file_fullpath_def))
                    {
                        isValid = false;
                        msg_error = "File path is not an absolute path.";
                    }

                    // ตรวจสอบชื่อไฟล์
                    string safeFileName = Path.GetFileName(file_fullpath_def);
                    if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
                    {
                        isValid = false;
                        msg_error = "Invalid or potentially dangerous file name.";
                    }

                    // ตรวจสอบอักขระที่อนุญาต
                    char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                        .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                    // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                    string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
                    string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                    if (isValid && !allowedExtensions.Contains(fileExtension))
                    {
                        isValid = false;
                        msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
                    }
                    // ตรวจสอบค่า folder 
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder)
                        || folder.Any(c => !AllowedCharacters.Contains(c)) || folder.Contains("\\"))
                    {
                        //isValid = false;  msg_error = "Invalid folder.";
                        throw new ApplicationException("An unexpected error Folder directory not found.");
                    }
                    if (isValid)
                    {
                        // สร้างพาธของโฟลเดอร์ที่อนุญาต
                        string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
                        if (isValid && !Directory.Exists(allowedDirectory))
                        {
                            //isValid = false;  msg_error = "Folder directory not found."; 
                            throw new ApplicationException("An unexpected error Folder directory not found.");
                        }

                        // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
                        string sourceFile = $"{safeFileName}";
                        if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
                                        string.IsNullOrWhiteSpace(sourceFile) ||
                                        sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                                        sourceFile.Contains("..") || sourceFile.Contains("\\")))
                        {
                            isValid = false;
                            msg_error = "Invalid fileName.";
                        }

                        // สร้างพาธเต็มของไฟล์
                        if (isValid)
                        {
                            fullPath = Path.Combine(allowedDirectory, sourceFile);
                            fullPath = Path.GetFullPath(fullPath);

                            // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
                            if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
                            {
                                isValid = false;
                                msg_error = "File is outside the allowed directory.";
                            }
                        }

                        // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        if (isValid)
                        {
                            FileInfo template = new FileInfo(fullPath);
                            if (!template.Exists)
                            {
                                isValid = false;
                                msg_error = "File not found.";
                            }

                            // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
                            if (isValid && template.IsReadOnly)
                            {
                                try
                                {
                                    template.IsReadOnly = false;
                                }
                                catch (Exception ex)
                                {
                                    isValid = false;
                                    msg_error = $"Failed to modify file attributes: {ex.Message}";
                                }
                            }
                        }

                        // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
                        if (isValid)
                        {
                            try
                            {
                                using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                                {
                                    // สามารถเปิดไฟล์ได้
                                }
                            }
                            catch (UnauthorizedAccessException)
                            {
                                isValid = false;
                                msg_error = "Access to the file is denied.";
                            }
                        }
                    }
                    #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


                    // หากทุกอย่างผ่านการตรวจสอบ
                    if (isValid)
                    {
                        // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
                        file_fullpath_name = fullPath;
                    }
                    else { file_fullpath_name = ""; }
                }
                else { file_fullpath_name = ""; }

                string role_type = "";
                ClassHazop clshazop = new ClassHazop();
                clshazop.check_role_user_active(user_approver_active, ref role_type);

                string doc_no = "";
                string doc_name = "";
                string pha_sub_software = "";
                string meeting_date = "";
                string meeting_time = "";
                string reference_moc = "";
                string comment = "";
                string approve_status = "";
                string approver_displayname = "XXXXX (TOP-XX)";

                string url = "";

                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                #region mail to
                string mail_admin_group = get_mail_admin_group();
                s_mail_to = mail_admin_group;
                #endregion mail to

                DataTable dt = new DataTable();

                string sqlstr = "";
                sqlstr = @" SELECT * from VW_EPHA_DATA_TO_EMAIL a WHERE a.request_approver = 1 AND a.id = @seq";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                if (role_type != "admin")
                {
                    sqlstr += " AND LOWER(a.approver_user_name) = LOWER(@user_approver_active)";
                    parameters.Add(new SqlParameter("@user_approver_active", SqlDbType.VarChar, 100) { Value = user_approver_active.ToLower() });
                }
                sqlstr += " ORDER BY a.no_approver ";
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
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                        dt = new DataTable();
                        dt = _conn.ExecuteAdapter(command).Tables[0];
                        //dt.TableName = "Table1";
                        dt.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (dt?.Rows.Count > 0)
                {
                    doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                    doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                    reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                    comment = dt.Rows[0]["approve_comment"]?.ToString() ?? "";
                    approve_status = dt.Rows[0]["approve_action_review"]?.ToString() ?? "";
                    approver_displayname = dt.Rows[0]["user_displayname"]?.ToString() ?? "";

                    pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                    meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                    meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                    s_mail_cc += dt.Rows[0]["user_email"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(dt.Rows[0]["request_email"]?.ToString()))
                    {
                        s_mail_cc += ";" + dt.Rows[0]["request_email"];
                    }
                }

                if (!string.IsNullOrEmpty(doc_no))
                {

                    #region URL Encryption
                    if (true)
                    {

                        string plainText = $"seq={seq}&pha_no={doc_no}&step=2";
                        string cipherText = EncryptString(plainText);

                        url = $"{server_url_by_action(sub_software)}{cipherText}";
                    }
                    #endregion URL Encryption

                    StringBuilder s_body = new StringBuilder();
                    s_body.Append("<html><body><font face='tahoma' size='2'>");
                    s_body.Append("Dear All,");

                    s_body.Append("<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
                    s_body.AppendFormat("{0} has approved the {1}, PHA No.{2} as detailed below:", approver_displayname, pha_sub_software.ToUpper(), doc_no);

                    s_body.AppendFormat("<br/><br/><b>Project Name</b> : {0}", doc_name);
                    s_body.AppendFormat("<br/><b>Date</b> : {0}", meeting_date);
                    s_body.AppendFormat("<br/><b>Time</b> : {0}", meeting_time);

                    if (!string.IsNullOrEmpty(reference_moc))
                    {
                        s_body.AppendFormat("<br/><b>Reference MOC</b> : {0}", reference_moc);
                    }
                    if (!string.IsNullOrEmpty(comment))
                    {
                        s_body.AppendFormat("<br/><br/><b>Comment :</b> {0}", comment);
                    }

                    s_body.Append("<br/><br/>");
                    s_body.AppendFormat("More details of the study, <a href='{0}' style='color: red;'>please click here</a>", url);

                    s_body.Append("<br/><br/>DISCLAIMER:");
                    s_body.Append("<br/>This email message (including any attachment) is intended only for the personal use of the recipient(s) named above...");
                    s_body.Append("</font></body></html>");

                    sendEmailModel data = new sendEmailModel
                    {
                        mail_subject = $"EPHA : {pha_sub_software.ToUpper()}, Approver Approved.",
                        mail_body = s_body.ToString(),
                        mail_to = s_mail_to,
                        mail_cc = s_mail_cc,
                        mail_from = s_mail_from
                    };

                    if (!string.IsNullOrEmpty(file_fullpath_name) && File.Exists(file_fullpath_name))
                    {
                        data.mail_attachments = file_fullpath_name;
                    }

                    return sendMail(data);
                }
                else
                {
                    return "";
                }

            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }

        }

        public string MailRejectByApprover(string seq, string sub_software, string user_approver_active)
        {
            if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(sub_software) || string.IsNullOrEmpty(user_approver_active))
            {
                return "Invalid input parameters.";
            }

            // ตรวจสอบ sub_software ว่ามีอยู่ใน whitelist
            var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };
            if (!allowedSubSoftware.Contains(sub_software.ToLower()))
            {
                return "Invalid sub_software.";
            }

            try
            {

                string file_fullpath_name = get_document_file_approver(seq, user_approver_active);

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = sub_software ?? "";
                    string msg_error = "";

                    //string folder = sub_software;
                    string fullPath = "";

                    #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

                    bool isValid = true;

                    // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
                    if (!Path.IsPathRooted(file_fullpath_def))
                    {
                        isValid = false;
                        msg_error = "File path is not an absolute path.";
                    }

                    // ตรวจสอบชื่อไฟล์
                    string safeFileName = Path.GetFileName(file_fullpath_def);
                    if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
                    {
                        isValid = false;
                        msg_error = "Invalid or potentially dangerous file name.";
                    }

                    // ตรวจสอบอักขระที่อนุญาต
                    char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                        .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                    // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                    string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
                    string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                    if (isValid && !allowedExtensions.Contains(fileExtension))
                    {
                        isValid = false;
                        msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
                    }
                    // ตรวจสอบค่า folder 
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder)
                        || folder.Any(c => !AllowedCharacters.Contains(c)) || folder.Contains("\\"))
                    {
                        //isValid = false;  msg_error = "Invalid folder.";
                        throw new ApplicationException("An unexpected error Folder directory not found.");
                    }
                    if (isValid)
                    {
                        // สร้างพาธของโฟลเดอร์ที่อนุญาต
                        string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
                        if (isValid && !Directory.Exists(allowedDirectory))
                        {
                            //isValid = false;  msg_error = "Folder directory not found."; 
                            throw new ApplicationException("An unexpected error Folder directory not found.");
                        }

                        // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
                        string sourceFile = $"{safeFileName}";
                        if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
                                        string.IsNullOrWhiteSpace(sourceFile) ||
                                        sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                                        sourceFile.Contains("..") || sourceFile.Contains("\\")))
                        {
                            isValid = false;
                            msg_error = "Invalid fileName.";
                        }

                        // สร้างพาธเต็มของไฟล์
                        if (isValid)
                        {
                            fullPath = Path.Combine(allowedDirectory, sourceFile);
                            fullPath = Path.GetFullPath(fullPath);

                            // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
                            if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
                            {
                                isValid = false;
                                msg_error = "File is outside the allowed directory.";
                            }
                        }

                        // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        if (isValid)
                        {
                            FileInfo template = new FileInfo(fullPath);
                            if (!template.Exists)
                            {
                                isValid = false;
                                msg_error = "File not found.";
                            }

                            // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
                            if (isValid && template.IsReadOnly)
                            {
                                try
                                {
                                    template.IsReadOnly = false;
                                }
                                catch (Exception ex)
                                {
                                    isValid = false;
                                    msg_error = $"Failed to modify file attributes: {ex.Message}";
                                }
                            }
                        }

                        // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
                        if (isValid)
                        {
                            try
                            {
                                using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                                {
                                    // สามารถเปิดไฟล์ได้
                                }
                            }
                            catch (UnauthorizedAccessException)
                            {
                                isValid = false;
                                msg_error = "Access to the file is denied.";
                            }
                        }
                    }
                    #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


                    // หากทุกอย่างผ่านการตรวจสอบ
                    if (isValid)
                    {
                        // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
                        file_fullpath_name = fullPath;
                    }
                    else { file_fullpath_name = ""; }
                }
                else { file_fullpath_name = ""; }


                string role_type = "";
                ClassHazop clshazop = new ClassHazop();
                clshazop.check_role_user_active(user_approver_active, ref role_type);

                string doc_no = "";
                string doc_name = "";
                string pha_sub_software = "";
                string meeting_date = "";
                string meeting_time = "";
                string reference_moc = "";
                string comment = "";
                string approve_status = "";
                string approver_displayname = "XXXXX (TOP-XX)";

                string url = "";

                string to_displayname = "All";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                #region mail to
                string mail_admin_group = get_mail_admin_group();
                s_mail_to = mail_admin_group;
                #endregion mail to

                cls = new ClassFunctions();
                DataTable dt = new DataTable();

                string sqlstr = "";
                sqlstr = @" SELECT * from VW_EPHA_DATA_TO_EMAIL a WHERE a.request_approver = 1 AND a.seq = @seq";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.Int) { Value = seq });
                if (role_type != "admin")
                {
                    sqlstr += " AND LOWER(a.approver_user_name) = LOWER(@user_approver_active)";
                    parameters.Add(new SqlParameter("@user_approver_active", SqlDbType.VarChar, 100) { Value = user_approver_active.ToLower() });
                }
                sqlstr += " ORDER BY a.no_approver ";
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
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                        dt = new DataTable();
                        dt = _conn.ExecuteAdapter(command).Tables[0];
                        //dt.TableName = "Table1";
                        dt.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (dt?.Rows.Count > 0)
                {
                    doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                    doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                    reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                    comment = dt.Rows[0]["approve_comment"]?.ToString() ?? "";
                    approve_status = dt.Rows[0]["approve_action_review"]?.ToString() ?? "";
                    approver_displayname = dt.Rows[0]["user_displayname"]?.ToString() ?? "";

                    pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                    meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                    meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                    s_mail_cc += dt.Rows[0]["user_email"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(dt.Rows[0]["request_email"]?.ToString()))
                    {
                        s_mail_cc += ";" + dt.Rows[0]["request_email"];
                    }
                }

                if (!string.IsNullOrEmpty(doc_no))
                {

                    #region url  
                    if (true)
                    {
                        string plainText = $"seq={seq}&pha_no={doc_no}&step=2";
                        string cipherText = EncryptString(plainText);

                        url = $"{server_url_by_action(sub_software)}{cipherText}";
                    }
                    #endregion url 

                    StringBuilder s_body = new StringBuilder();
                    s_body.Append("<html><body><font face='tahoma' size='2'>");
                    s_body.Append("Dear All,");

                    s_body.Append("<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
                    s_body.AppendFormat("{0} has rejected the {1}, PHA No.{2} as detailed below:", approver_displayname, pha_sub_software.ToUpper(), doc_no);

                    s_body.AppendFormat("<br/><br/><b>Project Name</b> : {0}", doc_name);
                    s_body.AppendFormat("<br/><b>Date</b> : {0}", meeting_date);
                    s_body.AppendFormat("<br/><b>Time</b> : {0}", meeting_time);

                    if (!string.IsNullOrEmpty(reference_moc)) { s_body.AppendFormat("<br/><b>Reference MOC</b> : {0}", reference_moc); }

                    s_body.Append("<br/><br/><b>Send back with comment :</b>");
                    s_body.AppendFormat("<br/><b>{0}</b>", comment);

                    s_body.Append("<br/><br/>");
                    s_body.AppendFormat("More details of the study, <a href='{0}' style='color: red;'> please click here</a>", url);

                    s_body.Append("<br/><br/>DISCLAIMER:");
                    s_body.Append("<br/>This email message is intended only for the personal use of the recipient(s) named above...");
                    s_body.Append("</font></body></html>");

                    sendEmailModel data = new sendEmailModel
                    {
                        mail_subject = $"EPHA : {pha_sub_software.ToUpper()}, Approver Send back with Comment.",
                        mail_body = s_body.ToString(),
                        mail_to = s_mail_to,
                        mail_cc = s_mail_cc,
                        mail_from = s_mail_from
                    };

                    if (!string.IsNullOrEmpty(file_fullpath_name) && File.Exists(file_fullpath_name))
                    {
                        data.mail_attachments = file_fullpath_name;
                    }

                    return sendMail(data);
                }
                else
                {
                    return "";
                }


            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }

        public string MailNotificationOutstandingAction(string user_name_active, string seq, string sub_software)
        {
            if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(sub_software))
            {
                return "Invalid input parameters.";
            }
            try
            {
                string url = "";
                string url_home_task = "";

                string to_displayname = "All";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                string date_now = DateTime.Now.ToString("dd/MMM/yyyy");
                string file_fullpath_name = "";

                if (sub_software == "hra")
                {
                    try
                    {
                        ReportModel param = new ReportModel
                        {
                            export_type = "pdf",
                            sub_software = sub_software,
                            seq = seq
                        };
                        ClassExcel cls = new ClassExcel();
                        file_fullpath_name = cls.export_potential_health_checklist_template(param, true);

                        if (!string.IsNullOrEmpty(file_fullpath_name))
                        {
                            string file_fullpath_def = file_fullpath_name;
                            string folder = sub_software ?? "";
                            string msg_error = "";

                            //string folder = sub_software;
                            string fullPath = "";

                            #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

                            bool isValid = true;

                            // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
                            if (!Path.IsPathRooted(file_fullpath_def))
                            {
                                isValid = false;
                                msg_error = "File path is not an absolute path.";
                            }

                            // ตรวจสอบชื่อไฟล์
                            string safeFileName = Path.GetFileName(file_fullpath_def);
                            if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
                            {
                                isValid = false;
                                msg_error = "Invalid or potentially dangerous file name.";
                            }

                            // ตรวจสอบอักขระที่อนุญาต
                            char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                                .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                            // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                            string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
                            string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                            if (isValid && !allowedExtensions.Contains(fileExtension))
                            {
                                isValid = false;
                                msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
                            }
                            // ตรวจสอบค่า folder
                            if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
                            {
                                isValid = false;
                                msg_error = "Invalid folder.";
                            }
                            if (isValid)
                            {
                                // สร้างพาธของโฟลเดอร์ที่อนุญาต
                                string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
                                if (isValid && !Directory.Exists(allowedDirectory))
                                {
                                    isValid = false;
                                    msg_error = "Folder directory not found.";
                                }

                                // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
                                string sourceFile = $"{safeFileName}";
                                if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
                                                string.IsNullOrWhiteSpace(sourceFile) ||
                                                sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                                                sourceFile.Contains("..") || sourceFile.Contains("\\")))
                                {
                                    isValid = false;
                                    msg_error = "Invalid fileName.";
                                }

                                // สร้างพาธเต็มของไฟล์
                                if (isValid)
                                {
                                    fullPath = Path.Combine(allowedDirectory, sourceFile);
                                    fullPath = Path.GetFullPath(fullPath);

                                    // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
                                    if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
                                    {
                                        isValid = false;
                                        msg_error = "File is outside the allowed directory.";
                                    }
                                }

                                // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                                if (isValid)
                                {
                                    FileInfo template = new FileInfo(fullPath);
                                    if (!template.Exists)
                                    {
                                        isValid = false;
                                        msg_error = "File not found.";
                                    }

                                    // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
                                    if (isValid && template.IsReadOnly)
                                    {
                                        try
                                        {
                                            template.IsReadOnly = false;
                                        }
                                        catch (Exception ex)
                                        {
                                            isValid = false;
                                            msg_error = $"Failed to modify file attributes: {ex.Message}";
                                        }
                                    }
                                }

                                // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
                                if (isValid)
                                {
                                    try
                                    {
                                        using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                                        {
                                            // สามารถเปิดไฟล์ได้
                                        }
                                    }
                                    catch (UnauthorizedAccessException)
                                    {
                                        isValid = false;
                                        msg_error = "Access to the file is denied.";
                                    }
                                }
                            }
                            #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


                            // หากทุกอย่างผ่านการตรวจสอบ
                            if (isValid)
                            {
                                // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
                                file_fullpath_name = fullPath;
                            }
                            else { file_fullpath_name = ""; }
                        }
                        else { file_fullpath_name = ""; }


                    }
                    catch { file_fullpath_name = ""; }
                }

                DataTable dt = new DataTable();
                string mail_admin_group = get_mail_admin_group();

                cls_conn = new ClassConnectionDb();
                DataTable dtOwner = new DataTable();
                ClassNoti classNoti = new ClassNoti();

                dtOwner = classNoti.DataDailyByActionRequired(user_name_active, seq, sub_software, true, false);
                dt = classNoti.DataDailyByActionRequired(user_name_active, seq, sub_software, false, false);

                #region mail to
                if (!string.IsNullOrEmpty(mail_admin_group))
                {
                    s_mail_cc = mail_admin_group;
                }

                string msg = "";
                if (dt?.Rows.Count > 0)
                {
                    foreach (DataRow ownerRow in dtOwner.Rows)
                    {
                        to_displayname = ownerRow["user_displayname"].ToString() ?? "";
                        s_mail_to = ownerRow["user_email"].ToString() ?? "";
                        string responder_user_name = ownerRow["user_name"].ToString() ?? "";

                        string pha_no = "";
                        //string pha_no = dt.Select($"user_name='{responder_user_name}'").FirstOrDefault()?["document_number"].ToString() ?? "";
                        var filterParameters = new Dictionary<string, object>();
                        filterParameters.Add("user_name", responder_user_name);
                        var (drDoc, iMerge) = FilterDataTable(dt, filterParameters);
                        if (drDoc != null)
                        {
                            if (drDoc?.Length > 0)
                            {
                                pha_no = drDoc.FirstOrDefault()?["document_number"].ToString() ?? "";
                            }
                        }

                        string plainText = $"seq={seq}&pha_no={pha_no}";
                        string cipherText = EncryptString(plainText);
                        url_home_task = $"{server_url_home_task}{cipherText}";

                        StringBuilder s_body = new StringBuilder();
                        s_body.Append("<html><body><font face='tahoma' size='2'>");
                        s_body.AppendFormat("Dear {0},", to_displayname);
                        s_body.Append(@"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.");
                        s_body.Append(@"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please response by reply this email within five working days prior auto proceed to next step.</small>");
                        s_body.Append(@"<br/>
                            <table style ='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   
                            <thead>    
                                <tr>
                                    <td style ='padding: 15px;' rowspan='1'>Task</td>
                                    <td style ='padding: 15px;' rowspan='1'>PHA Type</td>
                                    <td style ='padding: 15px;' rowspan='1'>Action Required</td>
                                    <td style ='padding: 15px;' rowspan='1'>Document Number</td>
                                    <td style ='padding: 15px;' rowspan='1'>Document Title</td>
                                    <td style ='padding: 15px;' rowspan='1'>Rev.</td>
                                    <td style ='padding: 15px;' rowspan='1'>Originator</td>
                                    <td style ='padding: 15px;' rowspan='1'>Received</td>
                                    <td style ='padding: 15px;' rowspan='1'>Due Date</td>
                                    <td style ='padding: 15px;' rowspan='1'>Remaining</td> 
                                    <td style ='padding: 15px;' rowspan='1'>Consolidator</td> 
                                </tr>
                            </thead>
                            <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'>");

                        int iNo = 1;

                        var filterParametersUser = new Dictionary<string, object>();
                        filterParametersUser.Add("user_name", responder_user_name);
                        var (drUser, iUser) = FilterDataTable(dt, filterParametersUser);
                        if (drUser != null)
                        {
                            if (drUser?.Length > 0)
                            {

                                for (int j = 0; j < drUser.Length; j++)
                                {
                                    string url_reject_comment = "";
                                    string url_reject = "";
                                    string url_approver = "";

                                    string doc_no = drUser[j]["document_number"]?.ToString() ?? "";

                                    string background_color = "white";
                                    string font_color = "black";
                                    int iRemaining = 0;
                                    bool action_status_close = (drUser[j]["remaining"]?.ToString() ?? "").ToLower() == "closed";

                                    try
                                    {
                                        iRemaining = Convert.ToInt32(drUser[j]["remaining"].ToString());
                                        if (iRemaining > 3)
                                        {
                                            background_color = "green"; font_color = "red";
                                        }
                                        else if (iRemaining > 0 && iRemaining < 3 && !action_status_close)
                                        {
                                            background_color = "yellow";
                                        }
                                        else if (iRemaining <= 0 && !action_status_close)
                                        {
                                            background_color = "red"; font_color = "white";
                                        }
                                    }
                                    catch { }

                                    if (true)
                                    {
                                        plainText = $"seq={seq}&pha_no={doc_no}&step=3";
                                        cipherText = EncryptString(plainText);
                                        if (drUser[j]["pha_status"].ToString() == "13")
                                        {
                                            plainText = $"seq={seq}&pha_no={doc_no}&step=3";
                                            cipherText = EncryptString(plainText);
                                            url = $"{server_url_by_action(sub_software)}{cipherText}";
                                        }
                                        else if (drUser[j]["pha_status"].ToString() == "21")
                                        {
                                            plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=required";
                                            cipherText = EncryptString(plainText);
                                            url = $"{server_url_by_action(sub_software)}{cipherText}";

                                            plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject";
                                            cipherText = EncryptString(plainText);
                                            url_reject = $"{server_url_by_action(sub_software)}{cipherText}";

                                            plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject_no_comment";
                                            cipherText = EncryptString(plainText);
                                            url_reject_comment = $"{server_url_by_action(sub_software)}{cipherText}";

                                            plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=approve";
                                            cipherText = EncryptString(plainText);
                                            url_approver = $"{server_url_by_action(sub_software)}{cipherText}";
                                        }
                                    }

                                    s_body.Append("<tr>");
                                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", iNo);
                                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", drUser[j]["pha_type"]);
                                    s_body.AppendFormat("<td style ='padding: 15px;'><a href='{0}'>{1}</a></td>", url, drUser[j]["action_required"]);
                                    s_body.AppendFormat("<td style ='padding: 15px;'><a href='{0}'>{1}</a></td>", url, drUser[j]["document_number"]);
                                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", drUser[j]["document_title"]);
                                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", drUser[j]["rev"]);
                                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", drUser[j]["originator"]);
                                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", drUser[j]["receivesd"]);
                                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", drUser[j]["due_date"]);
                                    s_body.AppendFormat("<td style ='padding: 15px; background-color:{0};color:{1};'>{2}</td>", background_color, font_color, drUser[j]["remaining"]);
                                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", drUser[j]["consolidator"]);
                                    s_body.Append("</tr>");
                                    iNo++;
                                }
                            }
                        }

                        s_body.Append("</tbody></table>");
                        s_body.Append("<br/><br/>The message of color assignment is as follow:");
                        s_body.Append("<table><tbody>");
                        s_body.Append("<tr><td style='width: 120px;padding:4px;background-color:green; color:red'><label>Green Color</label></td><td> : &gt; 3 days; this document has more than 3 days to complete your task</td></tr>");
                        s_body.Append("<tr><td style='width: 120px;padding:4px;background-color:yellow;'><label>Yellow Color</label></td><td> : &lt; 3 days; this document has less than 3 days to complete your task</td></tr>");
                        s_body.Append("<tr><td style='width: 120px;padding:4px;background-color:Red; color : white'><label>Red Color</label></td><td> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label></td></tr>");
                        s_body.Append("</tbody></table>");
                        s_body.AppendFormat("<br/><br/><a href='{0}'>Click here to access your Overall Tasks Window</a>", url_home_task);

                        s_body.Append("<br/><br/>DISCLAIMER:<br/>");
                        s_body.Append(@"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.");
                        s_body.Append("</font></body></html>");

                        sendEmailModel data = new sendEmailModel
                        {
                            mail_subject = $"EPHA OUTSTANDING ACTION NOTIFICATION_{to_displayname}_{date_now}",
                            mail_body = s_body.ToString(),
                            mail_to = s_mail_to,
                            mail_cc = s_mail_cc,
                            mail_from = s_mail_from
                        };

                        if (!string.IsNullOrEmpty(file_fullpath_name) && File.Exists(file_fullpath_name))
                        {
                            data.mail_attachments = file_fullpath_name;
                        }

                        msg = sendMail(data);
                        if (!string.IsNullOrEmpty(msg))
                        {
                            // Handle error
                        }

                    }
                }
                #endregion mail to

                return msg;
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }

        }

        public string MailNotificationReviewerReviewFollowup(string seq, string responder_user_name, string sub_software, Boolean responder_close_all)
        {
            if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(sub_software) || string.IsNullOrEmpty(responder_user_name))
            {
                return "Invalid input parameters.";
            }

            try
            {
                string url = "";
                string url_home_task = "";
                //string step_text = "Outstanding Action Notification";

                string to_displayname = "All";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                string date_now = DateTime.Now.ToString("dd/MMM/yyyy");

                DataTable dt = new DataTable();
                string mail_admin_group = get_mail_admin_group();

                cls_conn = new ClassConnectionDb();
                ClassNoti classNoti = new ClassNoti();
                dt = classNoti.DataDailyByActionRequired_ReviewApprove(seq, responder_user_name, sub_software, false, responder_close_all);

                #region mail to
                s_mail_to = mail_admin_group;
                s_mail_cc = "";

                string msg = "";
                if (dt?.Rows.Count > 0)
                {
                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        s_mail_cc += (dt.Rows[i]["user_email"] + ";");

                        if (url_home_task == "")
                        {
                            string pha_no = (dt.Rows[i]["document_number"] + "");

                            if (!string.IsNullOrEmpty(pha_no))
                            {
                                //insert keyBase64 to db 
                                string plainText = "seq=" + seq + "&pha_no=" + pha_no;
                                string cipherText = EncryptString(plainText);
                                url_home_task = $"{server_url_home_task}{cipherText}";

                            }
                        }
                    }

                    s_subject = "EPHA " + ("Outstanding Action Notification").ToString().ToUpper() + "_" + date_now;

                    s_body = "<html><body><font face='tahoma' size='2'>";
                    s_body += "Dear " + to_displayname + ",";

                    s_body += @"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.";
                    s_body += @"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please response by reply this email within five working days prior auto proceed to next step.</small>";


                    s_body += @"<br/>
                                <table style ='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   <thead>    
                                    <tr>
                                        <td style ='padding: 15px;' rowspan='1'>Task</td>
                                        <td style ='padding: 15px;' rowspan='1'>PHA Type</td>
                                        <td style ='padding: 15px;' rowspan='1'>Action Required</td>
                                        <td style ='padding: 15px;' rowspan='1'>Document Number</td>
                                        <td style ='padding: 15px;' rowspan='1'>Document Title</td>
                                        <td style ='padding: 15px;' rowspan='1'>Rev.</td>
                                        <td style ='padding: 15px;' rowspan='1'>Originator</td>
                                        <td style ='padding: 15px;' rowspan='1'>Received</td>
                                        <td style ='padding: 15px;' rowspan='1'>Due Date</td>
                                        <td style ='padding: 15px;' rowspan='1'>Remaining</td> 
                                        <td style ='padding: 15px;' rowspan='1'>Consolidator</td> 
                                    </tr>
                                </thead>
                                <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'> ";

                    int iNo = 1;
                    DataRow[] dr = dt.Select();
                    for (int a = 0; a < dr.Length; a++)
                    {
                        string doc_no = (dr[a]["document_number"] + "");

                        string background_color = "white";
                        string font_color = "black";
                        int iRemaining = 0;
                        Boolean action_status_close = (dr[a]["remaining"] + "").ToLower() == "closed";

                        try
                        {
                            iRemaining = Convert.ToInt32(dr[a]["remaining"] + "");
                            if (iRemaining > 3)
                            {
                                background_color = "green"; font_color = "red";
                            }
                            else if ((iRemaining > 0 && iRemaining < 3) && action_status_close == false)
                            {
                                background_color = "yellow";
                            }
                            else if (iRemaining <= 0 && action_status_close == false)
                            { background_color = "red"; font_color = "white"; }
                        }
                        catch { }

                        #region url  
                        url = "";
                        string url_def = "";
                        string url_approver = "";
                        string url_reject = "";
                        string url_reject_comment = "";
                        if (true)
                        {
                            //insert keyBase64 to db 
                            string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                            string cipherText = EncryptString(plainText);
                            if ((dr[a]["pha_status"] + "") == "13")
                            {
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                                cipherText = EncryptString(plainText);
                                url = $"{server_url_by_action(sub_software)}{cipherText}";
                            }
                            else if ((dr[a]["pha_status"] + "") == "21")
                            {
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=required";
                                cipherText = EncryptString(plainText);
                                url = $"{server_url_by_action(sub_software)}{cipherText}";

                                //reject 
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject";
                                cipherText = EncryptString(plainText);
                                url_reject = $"{server_url_by_action(sub_software)}{cipherText}";


                                //reject no comment
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject_no_comment";
                                cipherText = EncryptString(plainText);
                                url_reject_comment = $"{server_url_by_action(sub_software)}{cipherText}";

                                //approve
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=approve";
                                cipherText = EncryptString(plainText);
                                url_approver = $"{server_url_by_action(sub_software)}{cipherText}";
                            }

                        }
                        #endregion url 

                        s_body += "<tr>";
                        s_body += "<td style ='padding: 15px;'>" + (iNo) + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["pha_type"] + "</td>";//hazop
                        if ((dr[a]["action_required"] + "").ToLower() == "recommendation closing"
                            || (dr[a]["action_required"] + "").ToLower() == "review"
                            || (dr[a]["action_required"] + "").ToLower() == "review & approve")
                        {
                            url_def = url;
                        }
                        else if ((dr[a]["action_required"] + "").ToLower() == "approve")
                        {
                            url_def = url;
                        }
                        s_body += "<td style ='padding: 15px;'><a href='" + url_def + "'>" + dr[a]["action_required"] + "</a></td>";//Recommendation Closing, Review, Approve

                        s_body += "<td style ='padding: 15px;'><a href='" + url_def + "'>" + dr[a]["document_number"] + "</a></td>";//hazop-2023-0000023
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["document_title"] + "</td>";//xxmoc0003

                        s_body += "<td style ='padding: 15px;'>" + dr[a]["rev"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["originator"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["receivesd"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["due_date"] + "</td>";
                        s_body += "<td style ='padding: 15px; background-color:" + background_color + ";color:" + font_color + "; '>" + dr[a]["remaining"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["consolidator"] + "</td>";
                        s_body += "</tr>";
                        iNo += 1;

                    }

                    s_body += "</tbody>";
                    s_body += "</table>";

                    s_body += "<br/><br/>The message of color assignment is as follow:";

                    s_body += "<table>";
                    s_body += "<tbody>";
                    s_body += "<tr>";
                    s_body += "<td style='width: 120px;padding:4px;background-color:green; color:red'><label>Green Color</label></td>";
                    s_body += "<td> : &gt; 3 days; this document has more than 3 days to complete your task</td>";
                    s_body += "</tr>";
                    s_body += "<tr>";
                    s_body += "<td style='width: 120px;padding:4px;background-color:yellow;'><label>Yellow Color</label></td>";
                    s_body += "<td> : &lt; 3 days; this document has less than 3 days to complete your task</td>";
                    s_body += "</tr>";
                    s_body += "<tr>";
                    s_body += "<td style='width: 120px;padding:4px;background-color:Red; color : white'><label>Red Color</label></td>";
                    s_body += "<td> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label></td>";
                    s_body += "</tr>";
                    s_body += "</tbody>";
                    s_body += "</table>";
                    //s_body += "<br/><label style='width: 120px;padding:4px;background-color:green; color:red'>Green Color</label> : &gt; 3 days; this document has more than 3 days to complete your task";
                    //s_body += "<br/><label style='width: 120px;padding:4px;background-color:yellow;'>Yellow Color</label> : &lt; 3 days; this document has less than 3 days to complete your task";
                    //s_body += "<br/><label style='width: 130px;padding:4px;background-color:Red; color : white'>Red Color</label> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label>";

                    s_body += "<br/><br/><a href='" + url_home_task + "'>Click here to access your Overall Tasks Window</a>";

                    s_body += "<br/><br/>DISCLAIMER:";
                    s_body += "<br/>";
                    s_body += @"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.";
                    s_body += "</font></body></html>";


                    sendEmailModel data = new sendEmailModel();
                    data.mail_subject = s_subject;
                    data.mail_body = s_body;
                    data.mail_to = s_mail_to;
                    data.mail_cc = s_mail_cc;
                    data.mail_from = s_mail_from;

                    msg = sendMail(data);
                    if (msg != "") { }

                }
                #endregion mail to

                return msg;
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }

        }
        public string MailNotificationReviewerClosedAll(string seq, string sub_software)
        {
            if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(sub_software))
            {
                return "Invalid input parameters.";
            }
            try
            {
                #region call function export full report pdf  
                var sameRamTypeSoftwares = new HashSet<string> { "hazop", "jsea", "whatif" };
                string _ram_type = sameRamTypeSoftwares.Contains(sub_software) ? "5" : "";

                ReportByWorksheetModel param = new ReportByWorksheetModel
                {
                    seq = seq,
                    export_type = "pdf",
                    sub_software = sub_software,
                    user_name = "",
                    ram_type = _ram_type
                };
                ClassExcel clsExcel = new ClassExcel();
                string file_fullpath_name = "";
                file_fullpath_name = clsExcel.export_report_recommendation(param, false, true, true);

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = sub_software ?? "";
                    string msg_error = "";

                    //string folder = sub_software;
                    string fullPath = "";

                    #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

                    bool isValid = true;

                    // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
                    if (!Path.IsPathRooted(file_fullpath_def))
                    {
                        isValid = false;
                        msg_error = "File path is not an absolute path.";
                    }

                    // ตรวจสอบชื่อไฟล์
                    string safeFileName = Path.GetFileName(file_fullpath_def);
                    if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
                    {
                        isValid = false;
                        msg_error = "Invalid or potentially dangerous file name.";
                    }

                    // ตรวจสอบอักขระที่อนุญาต
                    char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                        .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                    // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                    string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
                    string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                    if (isValid && !allowedExtensions.Contains(fileExtension))
                    {
                        isValid = false;
                        msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
                    }
                    // ตรวจสอบค่า folder
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
                    {
                        isValid = false;
                        msg_error = "Invalid folder.";
                    }
                    if (isValid)
                    {
                        // สร้างพาธของโฟลเดอร์ที่อนุญาต
                        string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
                        if (isValid && !Directory.Exists(allowedDirectory))
                        {
                            isValid = false;
                            msg_error = "Folder directory not found.";
                        }

                        // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
                        string sourceFile = $"{safeFileName}";
                        if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
                                        string.IsNullOrWhiteSpace(sourceFile) ||
                                        sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                                        sourceFile.Contains("..") || sourceFile.Contains("\\")))
                        {
                            isValid = false;
                            msg_error = "Invalid fileName.";
                        }

                        // สร้างพาธเต็มของไฟล์
                        if (isValid)
                        {
                            fullPath = Path.Combine(allowedDirectory, sourceFile);
                            fullPath = Path.GetFullPath(fullPath);

                            // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
                            if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
                            {
                                isValid = false;
                                msg_error = "File is outside the allowed directory.";
                            }
                        }

                        // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        if (isValid)
                        {
                            FileInfo template = new FileInfo(fullPath);
                            if (!template.Exists)
                            {
                                isValid = false;
                                msg_error = "File not found.";
                            }

                            // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
                            if (isValid && template.IsReadOnly)
                            {
                                try
                                {
                                    template.IsReadOnly = false;
                                }
                                catch (Exception ex)
                                {
                                    isValid = false;
                                    msg_error = $"Failed to modify file attributes: {ex.Message}";
                                }
                            }
                        }

                        // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
                        if (isValid)
                        {
                            try
                            {
                                using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                                {
                                    // สามารถเปิดไฟล์ได้
                                }
                            }
                            catch (UnauthorizedAccessException)
                            {
                                isValid = false;
                                msg_error = "Access to the file is denied.";
                            }
                        }
                    }
                    #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


                    // หากทุกอย่างผ่านการตรวจสอบ
                    if (isValid)
                    {
                        // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
                        file_fullpath_name = fullPath;
                    }
                    else { file_fullpath_name = ""; }
                }
                else { file_fullpath_name = ""; }

                #endregion call function export full report pdf 

                DataTable dt = new DataTable();
                string mail_admin_group = get_mail_admin_group();

                string url = "";
                string url_home_task = "";
                //string step_text = "Outstanding Action Notification";

                string to_displayname = "All";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                string date_now = DateTime.Now.ToString("dd/MMM/yyyy");

                cls_conn = new ClassConnectionDb();
                //dt = classNoti.DataDailyByActionRequired_Closed(seq, sub_software, false, false);
                Boolean group_by_user = false;
                Boolean task_noti = false;
                List<SqlParameter> parameters = new List<SqlParameter>();

                string sqlstr = "";
                if (group_by_user)
                {
                    sqlstr = "SELECT DISTINCT a.user_name, a.user_displayname, a.user_email FROM VW_EPHA_ACTION_CLOSED a WHERE 1=1";
                }
                else
                {
                    sqlstr = "SELECT DISTINCT a.* FROM VW_EPHA_ACTION_CLOSED t WHERE 1=1";
                }

                // Add extra query if task_noti is true
                if (!task_noti)
                {
                    sqlstr += @" and a.task_noti = @task_noti";
                    parameters.Add(new SqlParameter("@task_noti", SqlDbType.Int) { Value = 0 });
                }

                // Add parameters for id_pha
                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " AND LOWER(a.seq) = LOWER(@seq)";
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                }
                if (group_by_user)
                {
                    sqlstr += " ORDER BY a.user_name, a.user_displayname, a.user_email";
                }
                else
                {
                    sqlstr += " ORDER BY a.user_name, a.action_sort, a.document_number, a.rev";
                }
                dt = new DataTable();
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
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                        dt = new DataTable();
                        dt = _conn.ExecuteAdapter(command).Tables[0];
                        //dt.TableName = "Table1";
                        dt.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                #region mail to
                s_mail_to = mail_admin_group;
                s_mail_cc = "";

                string msg = "";
                if (dt?.Rows.Count > 0)
                {
                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        s_mail_cc += (dt.Rows[i]["user_email"] + ";");

                        if (url_home_task == "")
                        {
                            string pha_no = (dt.Rows[i]["document_number"] + "");

                            if (!string.IsNullOrEmpty(pha_no))
                            {
                                //insert keyBase64 to db 
                                string plainText = $"seq={seq}&pha_no={pha_no}";
                                string cipherText = EncryptString(plainText);
                                url_home_task = $"{server_url_home_task}{cipherText}";
                            }
                        }
                    }

                    s_subject = "EPHA " + ("Outstanding Action Notification").ToString().ToUpper() + "_" + date_now;

                    s_body = "<html><body><font face='tahoma' size='2'>";
                    s_body += "Dear " + to_displayname + ",";

                    s_body += @"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.";
                    s_body += @"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please response by reply this email within five working days prior auto proceed to next step.</small>";


                    s_body += @"<br/>
                                <table style ='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   <thead>    
                                    <tr>
                                        <td style ='padding: 15px;' rowspan='1'>Task</td>
                                        <td style ='padding: 15px;' rowspan='1'>PHA Type</td>
                                        <td style ='padding: 15px;' rowspan='1'>Action Required</td>
                                        <td style ='padding: 15px;' rowspan='1'>Document Number</td>
                                        <td style ='padding: 15px;' rowspan='1'>Document Title</td>
                                        <td style ='padding: 15px;' rowspan='1'>Rev.</td>
                                        <td style ='padding: 15px;' rowspan='1'>Originator</td>
                                        <td style ='padding: 15px;' rowspan='1'>Received</td>
                                        <td style ='padding: 15px;' rowspan='1'>Due Date</td>
                                        <td style ='padding: 15px;' rowspan='1'>Remaining</td> 
                                        <td style ='padding: 15px;' rowspan='1'>Consolidator</td> 
                                    </tr>
                                </thead>
                                <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'> ";

                    int iNo = 1;
                    DataRow[] dr = dt.Select();
                    for (int a = 0; a < dr.Length; a++)
                    {
                        string doc_no = (dr[a]["document_number"] + "");

                        string background_color = "white";
                        string font_color = "black";
                        int iRemaining = 0;
                        Boolean action_status_close = (dr[a]["remaining"] + "").ToLower() == "closed";


                        try
                        {
                            iRemaining = Convert.ToInt32(dr[a]["remaining"] + "");
                            if (iRemaining > 3)
                            {
                                background_color = "green"; font_color = "red";
                            }
                            else if ((iRemaining > 0 && iRemaining < 3) && action_status_close == false)
                            {
                                background_color = "yellow";
                            }
                            else if (iRemaining <= 0 && action_status_close == false)
                            { background_color = "red"; font_color = "white"; }
                        }
                        catch { }

                        #region url  
                        url = "";
                        string url_def = "";
                        string url_approver = "";
                        string url_reject = "";
                        string url_reject_comment = "";
                        if (true)
                        {
                            //insert keyBase64 to db 
                            string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                            string cipherText = EncryptString(plainText);
                            if ((dr[a]["pha_status"] + "") == "13")
                            {
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                                cipherText = EncryptString(plainText);
                                url = $"{server_url_by_action(sub_software)}{cipherText}";
                            }
                            else if ((dr[a]["pha_status"] + "") == "21")
                            {
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=required";
                                cipherText = EncryptString(plainText);
                                url = $"{server_url_by_action(sub_software)}{cipherText}";

                                //reject 
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject";
                                cipherText = EncryptString(plainText);
                                url_reject = $"{server_url_by_action(sub_software)}{cipherText}";


                                //reject no comment
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject_no_comment";
                                cipherText = EncryptString(plainText);
                                url_reject_comment = $"{server_url_by_action(sub_software)}{cipherText}";

                                //approve
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=approve";
                                cipherText = EncryptString(plainText);
                                url_approver = $"{server_url_by_action(sub_software)}{cipherText}";
                            }

                        }
                        #endregion url 

                        s_body += "<tr>";
                        s_body += "<td style ='padding: 15px;'>" + (iNo) + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["pha_type"] + "</td>";//hazop

                        if ((dr[a]["action_required"] + "").ToLower() == "recommendation closing"
                            || (dr[a]["action_required"] + "").ToLower() == "review"
                            || (dr[a]["action_required"] + "").ToLower() == "review & approve")
                        {
                            url_def = url;
                        }
                        else if ((dr[a]["action_required"] + "").ToLower() == "approve")
                        {
                            url_def = url;
                        }
                        s_body += "<td style ='padding: 15px;'><a href='" + url_def + "'>" + dr[a]["action_required"] + "</a></td>";//Recommendation Closing, Review, Approve

                        s_body += "<td style ='padding: 15px;'><a href='" + url_def + "'>" + dr[a]["document_number"] + "</a></td>";//hazop-2023-0000023
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["document_title"] + "</td>";//xxmoc0003

                        s_body += "<td style ='padding: 15px;'>" + dr[a]["rev"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["originator"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["receivesd"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["due_date"] + "</td>";
                        s_body += "<td style ='padding: 15px; background-color:" + background_color + ";color:" + font_color + "; '>" + dr[a]["remaining"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["consolidator"] + "</td>";
                        s_body += "</tr>";
                        iNo += 1;

                    }

                    s_body += "</tbody>";
                    s_body += "</table>";

                    s_body += "<br/><br/>The message of color assignment is as follow:";
                    s_body += "<br/><label style='width: 120px;padding:4px;background-color:green; color:red'>Green Color</label> : &gt; 3 days; this document has more than 3 days to complete your task";
                    s_body += "<br/><label style='width: 120px;padding:4px;background-color:yellow;'>Yellow Color</label> : &lt; 3 days; this document has less than 3 days to complete your task";
                    s_body += "<br/><label style='width: 130px;padding:4px;background-color:Red; color : white'>Red Color</label> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label>";

                    s_body += "<br/><br/><a href='" + url_home_task + "'>Click here to access your Overall Tasks Window</a>";

                    s_body += "<br/><br/>DISCLAIMER:";
                    s_body += "<br/>";
                    s_body += @"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.";
                    s_body += "</font></body></html>";


                    sendEmailModel data = new sendEmailModel();
                    data.mail_subject = s_subject;
                    data.mail_body = s_body;
                    data.mail_to = s_mail_to;
                    data.mail_cc = s_mail_cc;
                    data.mail_from = s_mail_from;

                    if (file_fullpath_name != "")
                    {
                        if (File.Exists(file_fullpath_name))
                        {
                            data.mail_attachments = file_fullpath_name;
                        }
                    }
                    msg = sendMail(data);
                    if (msg != "")
                    {
                        //data.mail_attachments = null;
                        msg = sendMail(data);
                    }

                }
                #endregion mail to

                return msg;
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }

        public string MailNotificationApproverTA3(string seq, string sub_software, string seq_approver, string user_approver_active)
        {
            if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(sub_software) || string.IsNullOrEmpty(seq_approver) || string.IsNullOrEmpty(user_approver_active))
            {
                return "Invalid input parameters.";
            }
            try
            {
                #region call function export full report pdf  
                var sameRamTypeSoftwares = new HashSet<string> { "hazop", "whatif" };
                string _ram_type = sameRamTypeSoftwares.Contains(sub_software) ? "5" : "";

                string file_fullpath_name = get_document_file_approver(seq_approver, user_approver_active);

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = sub_software ?? "";
                    string msg_error = "";

                    //string folder = sub_software;
                    string fullPath = "";

                    #region ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ

                    bool isValid = true;

                    // ตรวจสอบว่าพาธเป็น Absolute path หรือไม่
                    if (!Path.IsPathRooted(file_fullpath_def))
                    {
                        isValid = false;
                        msg_error = "File path is not an absolute path.";
                    }

                    // ตรวจสอบชื่อไฟล์
                    string safeFileName = Path.GetFileName(file_fullpath_def);
                    if (isValid && (string.IsNullOrEmpty(safeFileName) || safeFileName.Contains("..")))
                    {
                        isValid = false;
                        msg_error = "Invalid or potentially dangerous file name.";
                    }

                    // ตรวจสอบอักขระที่อนุญาต
                    char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                        .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                    // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                    string fileExtension = Path.GetExtension(safeFileName).ToLowerInvariant();
                    string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                    if (isValid && !allowedExtensions.Contains(fileExtension))
                    {
                        isValid = false;
                        msg_error = "Invalid file type. Allowed types: Excel, PDF, Word, PNG, JPG, etc.";
                    }
                    // ตรวจสอบค่า folder 
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder)
                        || folder.Any(c => !AllowedCharacters.Contains(c)) || folder.Contains("\\"))
                    {
                        //isValid = false;  msg_error = "Invalid folder.";
                        throw new ApplicationException("An unexpected error Folder directory not found.");
                    }
                    if (isValid)
                    {
                        // สร้างพาธของโฟลเดอร์ที่อนุญาต
                        string allowedDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", folder);
                        if (isValid && !Directory.Exists(allowedDirectory))
                        {
                            //isValid = false;  msg_error = "Folder directory not found."; 
                            throw new ApplicationException("An unexpected error Folder directory not found.");
                        }

                        // ตรวจสอบอักขระของชื่อไฟล์อีกครั้ง
                        string sourceFile = $"{safeFileName}";
                        if (isValid && (sourceFile.Any(c => !AllowedCharacters.Contains(c)) ||
                                        string.IsNullOrWhiteSpace(sourceFile) ||
                                        sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                                        sourceFile.Contains("..") || sourceFile.Contains("\\")))
                        {
                            isValid = false;
                            msg_error = "Invalid fileName.";
                        }

                        // สร้างพาธเต็มของไฟล์
                        if (isValid)
                        {
                            fullPath = Path.Combine(allowedDirectory, sourceFile);
                            fullPath = Path.GetFullPath(fullPath);

                            // ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์ที่อนุญาต
                            if (!fullPath.StartsWith(allowedDirectory, StringComparison.OrdinalIgnoreCase))
                            {
                                isValid = false;
                                msg_error = "File is outside the allowed directory.";
                            }
                        }

                        // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        if (isValid)
                        {
                            FileInfo template = new FileInfo(fullPath);
                            if (!template.Exists)
                            {
                                isValid = false;
                                msg_error = "File not found.";
                            }

                            // ตรวจสอบสถานะ Read-Only และเปลี่ยนถ้าจำเป็น
                            if (isValid && template.IsReadOnly)
                            {
                                try
                                {
                                    template.IsReadOnly = false;
                                }
                                catch (Exception ex)
                                {
                                    isValid = false;
                                    msg_error = $"Failed to modify file attributes: {ex.Message}";
                                }
                            }
                        }

                        // ลองเปิดไฟล์เพื่อยืนยันว่าไฟล์สามารถเข้าถึงได้
                        if (isValid)
                        {
                            try
                            {
                                using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                                {
                                    // สามารถเปิดไฟล์ได้
                                }
                            }
                            catch (UnauthorizedAccessException)
                            {
                                isValid = false;
                                msg_error = "Access to the file is denied.";
                            }
                        }
                    }
                    #endregion ตรวจสอบว่าไฟล์มีอยู่หรือไม่ อีกรอบ


                    // หากทุกอย่างผ่านการตรวจสอบ
                    if (isValid)
                    {
                        // กำหนดค่า file_ResponseSheet ให้เป็น fullPath
                        file_fullpath_name = fullPath;
                    }
                    else { file_fullpath_name = ""; }
                }
                else { file_fullpath_name = ""; }


                #endregion call function export full report pdf  

                string doc_no = "";
                string doc_name = "";
                string pha_sub_software = "";
                string meeting_date = "";
                string meeting_time = "";
                string reference_moc = "";
                string comment = "";
                string approve_status = "";
                string approver_displayname = "XXXXX (TOP-XX)";

                string url = "";

                string to_displayname = "All";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                string mail_admin_group = get_mail_admin_group();

                ClassFunctions cls = new ClassFunctions();
                DataTable dt = new DataTable();
                List<SqlParameter> parameters = new List<SqlParameter>();

                if (sub_software == "hazop" || sub_software == "whatif")
                {
                    // Approver ที่ Active
                    sqlstr = @" select distinct h.pha_status, h.pha_no as pha_no,g.pha_request_name as pha_name,empre.user_email as request_email
                         , ta3.no 
                         , format(a.meeting_date, 'dd MMM yyyy') as meeting_date
                         , replace(a.meeting_start_time,'1/1/1970 ','') +' - '+ replace(a.meeting_end_time,'1/1/1970 ','') as meeting_time
                         , emp.user_displayname, emp.user_email, g.reference_moc
                         , lower(h.pha_sub_software) as pha_sub_software
                         , lower(ta3.user_name) as approver_user_name, ta3.approver_action_type 
                         from epha_t_header h
                         inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                         inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha)  
                         inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and a.seq = ta2.id_session  
                         inner join EPHA_T_APPROVER_TA3 ta3 on lower(h.id) = lower(ta3.id_pha) and a.seq = ta3.id_session and ta2.id = ta3.id_approver    
                         inner join VW_EPHA_PERSON_DETAILS emp on lower(ta3.user_name)  = lower(emp.user_name) 
                         inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name) 
                         inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s on h.id = s.id_pha and a.id = s.id_session and a.id_pha = s.id_pha  
                         where coalesce(ta3.approver_action_type,0) = 0 and h.id = @seq";
                    sqlstr += @" and lower(ta2.user_name) like lower(@user_approver_active)";
                    if (!string.IsNullOrEmpty(seq))
                    {
                        parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar)
                        { Value = cls.ChkSqlNum(seq, "N")?.ToString() });
                    }
                    if (!string.IsNullOrEmpty(user_approver_active))
                    {
                        parameters.Add(new SqlParameter("@user_approver_active", SqlDbType.VarChar)
                        { Value = cls.ChkSqlStr(user_approver_active ?? "", 100) });
                    }
                    sqlstr += @" order by convert(int,ta3.no) ";
                }

                dt = new DataTable();
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
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                        dt = new DataTable();
                        dt = _conn.ExecuteAdapter(command).Tables[0];
                        //dt.TableName = "Table1";
                        dt.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (dt?.Rows.Count > 0)
                {
                    doc_no = (dt.Rows[0]["pha_no"] + "");
                    doc_name = (dt.Rows[0]["pha_name"] + "");
                    reference_moc = (dt.Rows[0]["reference_moc"] + "");

                    approver_displayname = (dt.Rows[0]["user_displayname"] + "");

                    pha_sub_software = (dt.Rows[0]["pha_sub_software"] + "");
                    meeting_date = (dt.Rows[0]["meeting_date"] + "");
                    meeting_time = (dt.Rows[0]["meeting_time"] + "");

                    s_mail_to = (dt.Rows[0]["user_email"] + "");

                    s_mail_cc = mail_admin_group;
                    //cc originator
                    if ((dt.Rows[0]["request_email"] + "") != "")
                    {
                        s_mail_cc += ";" + (dt.Rows[0]["request_email"] + "");
                    }
                }
                else { return "data is null"; }

                if (!string.IsNullOrEmpty(doc_no))
                {
                    //insert keyBase64 to db 
                    string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                    string cipherText = EncryptString(plainText);

                    url = $"{server_url_by_action(sub_software)}{cipherText}";

                }

                s_subject = "EPHA : " + pha_sub_software.ToUpper() + ",Approval Responsibility Assigned to TA3.";

                s_body = "<html><body><font face='tahoma' size='2'>";
                s_body += "Dear " + approver_displayname + ",";

                s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp";
                s_body += " Approval Responsibility Assigned to TA3";
                s_body += ", PHA No." + doc_no + " as details mentioned below,";

                s_body += "<br/><br/><b>Project Name</b> : " + doc_name;
                s_body += "<br/><b>Date</b> : " + meeting_date;
                s_body += "<br/><b>Time</b> : " + meeting_time;

                if (reference_moc != "") { s_body += "<br/><b>Reference MOC</b> : " + reference_moc; }


                s_body += "<br/><br/>";
                s_body += "More details of the study, <font color='red'; text-decoration: underline;><a href='" + url + "'> please click here</a></font>";

                s_body += "<br/><br/>DISCLAIMER:";
                s_body += "<br/>";
                s_body += @"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.";
                s_body += "</font></body></html>";

                sendEmailModel data = new sendEmailModel();
                data.mail_subject = s_subject;
                data.mail_body = s_body;
                data.mail_to = s_mail_to;
                data.mail_cc = s_mail_cc;
                data.mail_from = s_mail_from;


                if (!string.IsNullOrEmpty(file_fullpath_name) && File.Exists(file_fullpath_name))
                {
                    data.mail_attachments = file_fullpath_name;
                }

                return sendMail(data);
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }
        }
        #endregion mail workflow last version


        #region file step approver

        private string get_document_file_approver(string seq, string user_approver_active)
        {
            if (string.IsNullOrEmpty(seq)) { return ""; }
            if (string.IsNullOrEmpty(user_approver_active)) { return ""; }

            string file_ResponseSheet = "";
            try
            {
                List<SqlParameter> parameters = new List<SqlParameter>();

                #region call function  export excel 
                sqlstr = @" select distinct da.document_file_path, lower(h.pha_sub_software) as sub_software
                         from epha_t_header h
                         inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                         inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha)  
                         inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and a.seq = ta2.id_session  
                         inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s on h.id = s.id_pha and a.id = s.id_session and a.id_pha = s.id_pha  
                         inner join EPHA_T_DRAWING_APPROVER da on lower(h.id) = lower(da.id_pha) and ta2.id_session = da.id_session and ta2.seq = da.id_approver 
						 where h.request_approver = 1 and isnull(da.document_file_name,'') <>''
                         and h.id = @seq and lower(ta2.user_name) like lower(@user_approver_active) ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                parameters.Add(new SqlParameter("@user_approver_active", SqlDbType.VarChar, 100) { Value = user_approver_active ?? "" });

                DataTable dtDrawing = new DataTable();
                //dtDrawing = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                #region Execute to Datable
                try
                {
                    _conn = new ClassConnectionDb(); _conn.OpenConnection();
                    try
                    {
                        var command = _conn.conn.CreateCommand();
                        command.CommandText = sqlstr;
                        //command.Parameters.Add(":costcenter", costcenter); 
                        foreach (var _param in parameters)
                        {
                            if (_param != null && !command.Parameters.Contains(_param.ParameterName))
                            {
                                command.Parameters.Add(_param);
                            }
                        }
                        dtDrawing = new DataTable();
                        dtDrawing = _conn.ExecuteAdapter(command).Tables[0];
                        dtDrawing.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (dtDrawing != null)
                {
                    if (dtDrawing?.Rows.Count > 0)
                    {
                        file_ResponseSheet = dtDrawing.Rows[0]["document_file_path"]?.ToString() ?? "";
                    }
                }

                //for (int i = 0; i < dtDrawing?.Rows.Count; i++)
                //{
                //    //https://localhost:7098/AttachedFileTemp/hazop/HAZOP-2024-0000004-DRAWING-202401102132.PDF
                //    //to
                //    //D:\dotnet6-epha-api\dotnet6-epha-api/wwwroot/AttachedFileTemp/Hazop/HAZOP Report 202311281602.pdf

                //    string sub_software = dtDrawing.Rows[i]["sub_software"]?.ToString() ?? "";
                //    string fileTemp = dtDrawing.Rows[i]["document_file_path"]?.ToString() ?? "";
                //    if (!string.IsNullOrEmpty(sub_software) || !string.IsNullOrEmpty(fileTemp))
                //    {
                //        string file_fullpath_name = "";
                //        string file_download_name = "";
                //        string msg = ClassFile.check_file_other(sub_software, ref file_fullpath_name, ref file_download_name, fileTemp);
                //        //if (string.IsNullOrEmpty(msg) && !string.IsNullOrEmpty(file_fullpath_name))
                //        //{
                //        // ??? ต้องปรับเป็นรวมไฟล์เดียวแล้วส่งเข้า mail
                //        //    if (file_ResponseSheet != "") { file_ResponseSheet += "|"; }
                //        //    file_ResponseSheet += file_fullpath_name;
                //        //}
                //    }
                //}
                #endregion call function  export excel 
            }
            catch { file_ResponseSheet = ""; }

            return file_ResponseSheet;
        }
        #endregion file step approver
        public string MailNotificationChangeActionOwner(string seq, string sub_software, string seq_worksheet_list)
        {
            if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(sub_software))
            {
                return "Invalid input parameters.";
            }

            try
            {
                // ตรวจสอบว่า sub_software อยู่ใน whitelist ที่อนุญาต
                var allowedSubSoftware = new HashSet<string> { "jsea", "whatif", "hra", "hazop" };
                if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                {
                    return "Invalid sub_software value.";
                }

                // ตรวจสอบว่า seq และ seq_worksheet_list เป็นค่าที่เหมาะสม (ตัวเลข)
                if (!Regex.IsMatch(seq, @"^\d+$") || (!string.IsNullOrEmpty(seq_worksheet_list) && !Regex.IsMatch(seq_worksheet_list, @"^\d+(,\d+)*$")))
                {
                    return "Invalid seq or seq_worksheet_list value.";
                }

                string msg = "";
                string url = "";
                string url_home_task = "";
                string step_text = "Outstanding Action Notification (Change Action Owner)";
                string to_displayname = "All";
                string s_mail_to = "";
                string s_mail_cc = "";
                string s_mail_from = "";

                string date_now = DateTime.Now.ToString("dd/MMM/yyyy");

                //// ใช้ parameterized query เพื่อป้องกัน SQL Injection
                //var parameters = new List<SqlParameter>();
                //parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                //parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software }); 
                //if (!string.IsNullOrEmpty(seq_worksheet_list))
                //{
                //    parameters.Add(new SqlParameter("@seq_worksheet_list", SqlDbType.VarChar, 50) { Value = seq_worksheet_list });
                //}

                DataTable dt = new DataTable();
                string mail_admin_group = get_mail_admin_group();

                cls_conn = new ClassConnectionDb();
                ClassNoti classNoti = new ClassNoti();
                dt = classNoti.DataDailyByActionRequired_Responder(seq, sub_software, false);

                // ตรวจสอบว่ามีข้อมูลใน dt หรือไม่
                if (dt?.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string user_name = dt.Rows[i]["user_name"]?.ToString() ?? "";
                        string user_email = dt.Rows[i]["user_email"]?.ToString() ?? "";
                        string seq_worksheet = dt.Rows[i]["seq_worksheet"]?.ToString() ?? "";

                        if (!string.IsNullOrEmpty(seq_worksheet_list))
                        {
                            HashSet<string> seqWorksheetSet = new HashSet<string>(seq_worksheet_list.Split(',').Select(s => s.Trim()));
                            if (!seqWorksheetSet.Contains(seq_worksheet)) { continue; }
                        }

                        to_displayname = dt.Rows[i]["user_displayname"]?.ToString() ?? "";

                        s_mail_to = string.Join(";", dt.AsEnumerable().Select(row => row["user_email"].ToString()).ToArray());
                        s_mail_cc = mail_admin_group;

                        if (string.IsNullOrEmpty(url_home_task))
                        {
                            if (true)
                            {
                                string pha_no = dt.Rows[i]["document_number"]?.ToString() ?? "";
                                string plainText = $"seq={seq}&pha_no={pha_no}";
                                string cipherText = EncryptString(plainText);
                                url_home_task = $"{server_url_home_task}{cipherText}";
                            }
                        }

                        // สร้างเนื้อหา email
                        StringBuilder s_body = new StringBuilder();
                        s_body.Append("<html><body><font face='tahoma' size='2'>");
                        s_body.AppendFormat("Dear {0},", to_displayname);
                        s_body.Append(@"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.");
                        s_body.Append(@"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please respond by replying to this email within five working days.</small>");

                        s_body.Append(@"<br/>
                    <table style ='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   
                    <thead>    
                    <tr>
                        <td style ='padding: 15px;'>Task</td>
                        <td style ='padding: 15px;'>PHA Type</td>
                        <td style ='padding: 15px;'>Action Required</td>
                        <td style ='padding: 15px;'>Document Number</td>
                        <td style ='padding: 15px;'>Document Title</td>
                        <td style ='padding: 15px;'>Rev.</td>
                        <td style ='padding: 15px;'>Originator</td>
                        <td style ='padding: 15px;'>Received</td>
                        <td style ='padding: 15px;'>Due Date</td>
                        <td style ='padding: 15px;'>Remaining</td> 
                        <td style ='padding: 15px;'>Consolidator</td> 
                    </tr>
                    </thead>
                    <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'> ");

                        foreach (DataRow row in dt.Rows)
                        {
                            s_body.Append("<tr>");
                            s_body.AppendFormat("<td>{0}</td>", row["task"]);
                            s_body.AppendFormat("<td>{0}</td>", row["pha_type"]);
                            s_body.AppendFormat("<td>{0}</td>", row["action_required"]);
                            s_body.AppendFormat("<td>{0}</td>", row["document_number"]);
                            s_body.AppendFormat("<td>{0}</td>", row["document_title"]);
                            s_body.AppendFormat("<td>{0}</td>", row["rev"]);
                            s_body.AppendFormat("<td>{0}</td>", row["originator"]);
                            s_body.AppendFormat("<td>{0}</td>", row["received"]);
                            s_body.AppendFormat("<td>{0}</td>", row["due_date"]);
                            s_body.AppendFormat("<td>{0}</td>", row["remaining"]);
                            s_body.AppendFormat("<td>{0}</td>", row["consolidator"]);
                            s_body.Append("</tr>");
                        }

                        s_body.Append("</tbody></table>");
                        s_body.AppendFormat("<br/><br/><a href='{0}'>Click here to access your Overall Tasks Window</a>", url_home_task);
                        s_body.Append("</font></body></html>");

                        sendEmailModel data = new sendEmailModel
                        {
                            mail_subject = step_text.ToUpper() + "_" + date_now,
                            mail_body = s_body.ToString(),
                            mail_to = s_mail_to,
                            mail_cc = s_mail_cc,
                            mail_from = s_mail_from
                        };

                        msg = sendMail(data);
                        if (msg != "") { return msg; }
                    }
                }

                return msg;
            }
            catch (Exception ex_mail) { return ex_mail.Message.ToString(); }

        }

    }
}
