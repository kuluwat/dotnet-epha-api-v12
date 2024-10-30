using Class;
using Microsoft.AspNetCore.Mvc;
using Model;
using dotnet_epha_api.Class;
using Microsoft.Exchange.WebServices.Data;
using System.Diagnostics;
using System.Data;
//using attributes;
//using services.interfaces;
//using jwts;
//using System.Diagnostics; 

namespace Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    //[IgnoreAntiforgeryToken] // ข้ามการตรวจสอบ CSRF
    public class FlowController : ControllerBase
    {
        //private readonly IAuthenticationService _authenticationService;
        //public FlowController(IAuthenticationService authenticationService)
        //{
        //    _authenticationService = authenticationService;
        //}
        //private string InsertTransactionLog(string module, string sub_software, Object param, ref string token_log)
        //{
        //    try
        //    {
        //        string param_text = JsonConvert.SerializeObject(param, Formatting.Indented) ?? "";
        //        int maxLength = 4000; // Set this to the maximum length your column supports
        //        if (param_text.Length > maxLength)
        //        {
        //            param_text = param_text.Substring(0, maxLength);
        //        }

        //        //insert log    
        //        ClassTransactionLog clstranlog = new ClassTransactionLog();
        //        string msg = clstranlog.insert_log(module, sub_software, param_text, ref token_log);
        //        if ((msg?.ToString() == ""))
        //        {
        //            msg = "error insert log : " + token_log;
        //        }
        //        return msg;
        //    }
        //    catch (Exception ex) { return ex.Message.ToString(); }
        //}

        #region uploadfile

        //[HttpPost("export_excel_to_pdf", Name = "export_excel_to_pdf")]
        //public string export_excel_to_pdf()
        //{
        //    // Paths to input and output files
        //    string inputFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp", "HAZOP Report Template.xlsx");
        //    string outputDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "AttachedFileTemp");

        //    // Path to LibreOffice in the project
        //    string libreOfficePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "LibreOffice", "program", "soffice.exe");

        //    try
        //    {
        //        // ตรวจสอบว่าไฟล์ LibreOffice อยู่ในที่ถูกต้อง
        //        if (!System.IO.File.Exists(libreOfficePath))
        //        {
        //            throw new FileNotFoundException("LibreOffice executable not found.");
        //        }

        //        // ตรวจสอบว่าไฟล์ input มีอยู่
        //        if (!System.IO.File.Exists(inputFilePath))
        //        {
        //            throw new FileNotFoundException("Input Excel file not found.");
        //        }

        //        // ตรวจสอบว่าโฟลเดอร์ output มีอยู่
        //        if (!Directory.Exists(outputDirectory))
        //        {
        //            throw new DirectoryNotFoundException("Output directory not found.");
        //        }

        //        // Set up the process start info
        //        var startInfo = new ProcessStartInfo
        //        {
        //            FileName = libreOfficePath,
        //            Arguments = $"--headless --convert-to pdf \"{inputFilePath}\" --outdir \"{outputDirectory}\"",
        //            RedirectStandardOutput = true,
        //            RedirectStandardError = true,
        //            UseShellExecute = false,
        //            CreateNoWindow = true
        //        };

        //        // Start the process
        //        using (var process = Process.Start(startInfo))
        //        {
        //            // Read output and error (if any)
        //            string output = process.StandardOutput.ReadToEnd();
        //            string error = process.StandardError.ReadToEnd();

        //            process.WaitForExit();

        //            if (process.ExitCode == 0)
        //            {
        //                Console.WriteLine("PDF conversion successful.");
        //                Console.WriteLine($"Output PDF saved to: {outputDirectory}");
        //            }
        //            else
        //            {
        //                Console.WriteLine($"Error during conversion: {error}");
        //                return $"Error during conversion: {error}";
        //            }
        //        }
        //    }
        //    catch (FileNotFoundException fnfEx)
        //    {
        //        Console.WriteLine(fnfEx.Message);
        //        return fnfEx.Message;
        //    }
        //    catch (DirectoryNotFoundException dnEx)
        //    {
        //        Console.WriteLine(dnEx.Message);
        //        return dnEx.Message;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Exception: {ex.Message}");
        //        return $"Exception: {ex.Message}";
        //    }

        //    return Path.Combine(outputDirectory, "HAZOP AttendeeSheet Template.pdf"); // คืนค่าพาธของไฟล์ PDF ที่ถูกสร้างขึ้น
        //}

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("importfile_data_jsea", Name = "importfile_data_jsea")]
        public string importfile_data_jsea([FromForm] uploadFile param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("importfile_data_jsea", "jsea", param, ref token_log);
            try
            {
                // ตรวจสอบค่า param เพื่อป้องกันปัญหา Dereference หลังจาก null check
                if (param == null)
                {
                    return ClassJSON.SetJSONresultRef(ClassFile.refMsg("Error", "Invalid parameters."));
                }

                // กำหนดชนิดไฟล์ที่อนุญาต
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };

                // ตรวจสอบไฟล์แต่ละไฟล์ที่อัพโหลด
                foreach (var file in param.file_obj)
                {
                    var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
                    if (!allowedExtensions.Contains(extension))
                    {
                        // ถ้าไฟล์มีชนิดที่ไม่ได้รับอนุญาต ให้คืนค่าข้อความผิดพลาด
                        return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave("false", "File type not allowed.", "", "", "", ""));
                    }
                }

                string sub_software = "jsea";
                try { sub_software = param?.sub_software ?? ""; } catch { }

                ClassHazopSet cls = new ClassHazopSet();
                return cls.importfile_data_jsea(param, sub_software);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("uploadfile_data", Name = "uploadfile_data")]
        public string uploadfile_data([FromForm] uploadFile param)
        {
            string sub_software = "hazop";
            try { sub_software = param?.sub_software ?? ""; } catch { }

            string token_log = "";
            string msg = "";//InsertTransactionLog("uploadfile_data", sub_software, param, ref token_log);
            try
            {
                // กำหนดชนิดไฟล์ที่อนุญาต
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };

                // ตรวจสอบไฟล์แต่ละไฟล์ที่อัพโหลด
                foreach (var file in param.file_obj)
                {
                    var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
                    if (!allowedExtensions.Contains(extension))
                    {
                        // ถ้าไฟล์มีชนิดที่ไม่ได้รับอนุญาต ให้คืนค่าข้อความผิดพลาด
                        return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave("false", "File type not allowed.", "", "", "", ""));
                    }
                }

                ClassHazopSet cls = new ClassHazopSet();
                return cls.uploadfile_data(param, sub_software);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("uploadfile_data_followup", Name = "uploadfile_data_followup")]
        public string uploadfile_data_followup([FromForm] uploadFile param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("uploadfile_data_followup", "followup", param, ref token_log);

            try
            {
                // กำหนดชนิดไฟล์ที่อนุญาต
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };

                // ตรวจสอบไฟล์แต่ละไฟล์ที่อัพโหลด
                foreach (var file in param.file_obj)
                {
                    var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
                    if (!allowedExtensions.Contains(extension))
                    {
                        // ถ้าไฟล์มีชนิดที่ไม่ได้รับอนุญาต ให้คืนค่าข้อความผิดพลาด
                        return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave("false", "File type not allowed.", "", "", "", ""));
                    }
                }

                // เรียกใช้ฟังก์ชัน uploadfile_data ถ้าผ่านการตรวจสอบแล้ว
                ClassHazopSet cls = new ClassHazopSet();
                return cls.uploadfile_data(param, "followup");
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("copy_pdf_file", Name = "copy_pdf_file")]
        public string copy_pdf_file(CopyFileModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("", "copy_pdf_file", param, ref token_log);
            try
            {
                string _file_fullpath_name = param.file_path ?? "";
                string _folder = param.sub_software ?? "hazop";
                if (string.IsNullOrEmpty(_file_fullpath_name) || string.IsNullOrEmpty(_folder))
                {
                    msg = "Invalid file path/folder.";
                }
                else
                {
                    string _file_fullpath_name_new = "";
                    if (string.IsNullOrEmpty(ClassFile.check_file_on_server(_folder, _file_fullpath_name, ref _file_fullpath_name_new)))
                    {
                        ClassExcel classExcel = new ClassExcel();
                        return classExcel.copy_pdf_file(param);
                    }
                    else { msg = "The file is not within the allowed directory."; }
                }
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        #endregion uploadfile

        #region Function All

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("edit_worksheet", Name = "edit_worksheet")]
        public string edit_worksheet(SetDocWorksheetModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("edit_worksheet", (param?.sub_software ?? "all"), param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.edit_worksheet(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_approve", Name = "set_approve")]
        public string set_approve(SetDocApproveModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("set_approve", (param?.sub_software ?? "all"), param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_approve(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_transfer_monitoring", Name = "set_transfer_monitoring")]
        public string set_transfer_monitoring(SetDocTransferMonitoringModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("set_transfer_monitoring", (param?.sub_software ?? "all"), param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_transfer_monitoring(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_approve_ta3", Name = "set_approve_ta3")]
        public string set_approve_ta3(SetDocApproveTa3Model param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("set_approve_ta3", (param?.sub_software ?? "all"), param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_approve_ta3(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        #endregion Function All

        #region Mail

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("MailToPHAConduct", Name = "MailToPHAConduct")]
        public string MailToPHAConduct(string seq, string sub_software)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("MailToPHAConduct", "all", new object(), ref token_log);
            try
            {
                ClassEmail cls = new ClassEmail();
                return cls.MailNotificationWorkshopInvitation(seq, sub_software);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }


        //[ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        //[HttpPost("MailToActionOwner", Name = "MailToActionOwner")]
        //public string MailToActionOwner(string seq, string sub_software)
        //{
        //    string token_log = "";
        //    string msg = "";//InsertTransactionLog("MailToActionOwner", "all", new object(), ref token_log);
        //    try
        //    {
        //        if (string.IsNullOrEmpty(seq))
        //        {
        //            msg = "Invalid file seq.";
        //        }
        //        else
        //        {
        //            if (string.IsNullOrWhiteSpace(sub_software))
        //            {
        //                msg = "Invalid file sub software.";
        //            }
        //            else
        //            {
        //                // ตรวจสอบ sub_software ว่ามีอยู่ใน whitelist
        //                var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra" };
        //                if (!allowedSubSoftware.Contains(pha_sub_software.ToLower()))
        //                {
        //                    return ClassJSON.SetJSONresultRef(ClassFile.refMsg("Error", "Invalid sub_software."));
        //                }

        //                ClassEmail cls = new ClassEmail();
        //                return cls.MailToActionOwner(seq, sub_software);
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
        //    }
        //    return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        //}


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("load_notification", Name = "load_notification")]
        public string load_notification(LoadDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("load_notification", "all", new object(), ref token_log);
            try
            {
                ClassHazop cls = new ClassHazop();
                return cls.get_notification(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("send_notification_member_review", Name = "send_notification_member_review")]
        public string send_notification_member_review(SetDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("send_notification_member_review", "all", new object(), ref token_log);
            try
            {
                ClassEmail cls = new ClassEmail();
                string id_session = "";
                string ret = cls.MailNotificationMemberReview((param.pha_seq + ""), (param.sub_software + ""));

                return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave(ret, "Msg :" + (ret == "" ? "true" : ret) + ",Session Last :" + id_session, "", "", "", ""));
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("send_notification_daily", Name = "send_notification_daily")]
        public string send_notification_daily(SetDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("send_notification_daily", "all", new object(), ref token_log);
            try
            {
                string user_name = param.user_name?.ToString() ?? "";
                string role_type = ClassLogin.GetUserRoleFromDb(user_name);
                string pha_seq = param.pha_seq?.ToString() ?? "";
                string sub_software = param.sub_software?.ToString() ?? "";

                ClassEmail cls = new ClassEmail();
                string ret = cls.MailNotificationOutstandingAction("", pha_seq, sub_software);

                return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave(ret, "Msg :" + (ret == "" ? "true" : ret), "", "", "", ""));
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("MailNotificationDaily", Name = "MailNotificationDaily")]
        public string MailNotificationDaily(LoadDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("MailNotificationDaily", "noti", param, ref token_log);
            try
            {
                string user_name = (param.user_name + "");
                string role_type = ClassLogin.GetUserRoleFromDb(user_name);
                string seq = (param.token_doc + "");
                string sub_software = (param.sub_software + "");

                ClassEmail classEmail = new ClassEmail();
                return classEmail.MailNotificationOutstandingAction("", seq, sub_software);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        #endregion Mail

        #region Hazop

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_hazop_details", Name = "get_hazop_details")]
        public string get_hazop_details(LoadDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("get_hazop_details", "hazop", param, ref token_log);
            try
            {
                param.sub_software = "hazop";
                ClassHazop cls = new ClassHazop();
                return cls.get_details(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));


        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_hazop", Name = "set_hazop")]
        public string set_hazop(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("set_hazop", "hazop", param, ref token_log);

            try
            {
                // ตรวจสอบว่า flow_action อยู่ในค่าที่เรายอมรับได้
                var allowedFlowActions = new HashSet<string>
            {
                "save",
                "submit",
                "submit_register",
                "submit_without",
                "submit_moc",
                "submit_generate_full_report",
                "edit_worksheet",
                "confirm_submit_generate",
                "confirm_submit_generate_without",
                "confirm_submit_register_without",
                "confirm_submit_register",
                "change_action_owner",
                "change_approver"
            };
                if (!allowedFlowActions.Contains(param.flow_action))
                {
                    return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave("false", $"Invalid flow_action.{param.flow_action}", "", "", "", ""));
                }

                ClassHazopSet cls = new ClassHazopSet();
                param.sub_software = "hazop"; // ตั้งค่า sub_software ให้เป็น "hazop"

                // ตรวจสอบ flow_action เพื่อเลือกการดำเนินการที่ถูกต้อง
                if (param.flow_action == "change_action_owner")
                {
                    return cls.set_workflow_change_employee(param);
                }
                else if (param.flow_action == "change_approver")
                {
                    return cls.set_workflow_change_employee(param);
                }
                else
                {
                    return cls.set_workflow(param);
                }
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
                return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave("false", msg, "", "", "", ""));
            }
        }

        #endregion Hazop


        #region Jsea

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_jsea_details", Name = "get_jsea_details")]
        public string get_jsea_details(LoadDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("get_jsea_details", "jsea", param, ref token_log);
            try
            {
                param.sub_software = "jsea";
                ClassHazop cls = new ClassHazop();
                return cls.get_details(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_jsea", Name = "set_jsea")]
        public string set_jsea(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("set_jsea", "jsea", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                param.sub_software = "jsea";
                return cls.set_workflow(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        #endregion Jsea


        #region Whatif

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_whatif_details", Name = "get_whatif_details")]
        public string get_whatif_details(LoadDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("get_whatif_details", "whatif", param, ref token_log);
            try
            {
                param.sub_software = "whatif";
                ClassHazop cls = new ClassHazop();
                return cls.get_details(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_whatif", Name = "set_whatif")]
        public string set_whatif(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("set_whatif", "whatif", param, ref token_log);

            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                param.sub_software = "whatif";

                //20240419 เพิ่ม flow_action : change_action_owner --> เป็นการแก้ไขรายชื่อ action owner 
                if (param.flow_action == "change_action_owner")
                {
                    return cls.set_workflow_change_employee(param);
                }

                //20240516 เพิ่ม flow_action : change_approver --> เป็นการแก้ไขรายชื่อ approver 
                if (param.flow_action == "change_approver")
                {
                    return cls.set_workflow_change_employee(param);
                }

                return cls.set_workflow(param);
            }
            catch (Exception e)
            {
                // เพิ่มการบันทึกข้อผิดพลาดลงระบบบันทึก (logging) 
                msg += $" method error: {e.Message} -> token log: {token_log}";
            }

            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        #endregion Whatif

        #region Hra


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_hra_details", Name = "get_hra_details")]
        public string get_hra_details(LoadDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("get_hra_details", "hra", param, ref token_log);
            try
            {
                param.sub_software = "hra";
                ClassHazop cls = new ClassHazop();
                return cls.get_details(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_hra", Name = "set_hra")]
        public string set_hra(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("set_hra", "hra", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                param.sub_software = "hra";

                //20240419 เพิ่ม flow_action : change_action_owner --> เป็นการแก้ไขรายชื่อ action owner 
                if (param.flow_action == "change_action_owner") { return cls.set_workflow_change_employee(param); }

                //20240516 เพิ่ม flow_action : change_approver --> เป็นการแก้ไขรายชื่อ approver 
                if (param.flow_action == "change_approver") { return cls.set_workflow_change_employee(param); }

                return cls.set_workflow(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        #endregion Hra

        #region Page Search

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("load_page_search_details", Name = "load_page_search_details")]
        public string load_page_search_details(LoadDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("load_page_search_details", "all", param, ref token_log);
            try
            {
                ClassHazop cls = new ClassHazop();
                return cls.get_search_details(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        #endregion Page Search

        #region Master RAM 

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_ram", Name = "set_master_ram")]
        public string set_master_ram(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("set_master_ram", "all", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_master_ram(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        #endregion Master RAM 

        #region follow up  

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("load_follow_up", Name = "load_follow_up")]
        public string load_follow_up(LoadDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("load_follow_up", "follow_up", param, ref token_log);
            try
            {
                ClassHazop cls = new ClassHazop();
                return cls.get_followup(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("load_follow_up_details", Name = "load_follow_up_details")]
        public string load_follow_up_details(LoadDocFollowModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("load_follow_up_details", "follow_up", param, ref token_log);
            try
            {
                ClassHazop cls = new ClassHazop();
                return cls.get_followup_detail(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_follow_up", Name = "set_follow_up")]
        public string set_follow_up(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("set_follow_up", "follow_up", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_follow_up(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_follow_up_review", Name = "set_follow_up_review")]
        public string set_follow_up_review(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("set_follow_up_review", "follow_up_review", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_follow_up_review(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        #endregion follow up 

        #region home tasks

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("load_home_tasks", Name = "load_home_tasks")]
        public string load_home_tasks(LoadDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("load_home_tasks", "home_tasks", param, ref token_log);
            try
            {
                ClassHazop cls = new ClassHazop();
                return cls.get_hometasks(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));


        }
        #endregion home tasks

        #region export hazop 

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_hazop_report", Name = "export_hazop_report")]
        public string export_hazop_report(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_hazop_report", "hazop", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_hazop_worksheet", Name = "export_hazop_worksheet")]
        public string export_hazop_worksheet(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_hazop_worksheet", "hazop", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                //return cls.export_hazop_worksheet(param);
                return cls.export_full_report(param, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_hazop_recommendation", Name = "export_hazop_recommendation")]
        public string export_hazop_recommendation(ReportByWorksheetModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_hazop_recommendation", "hazop", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_report_recommendation(param, false, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_hazop_ram", Name = "export_hazop_ram")]
        public string export_hazop_ram(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_hazop_ram", "hazop", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_template_ram(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_hazop_guidewords", Name = "export_hazop_guidewords")]
        public string export_hazop_guidewords(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_hazop_guidewords", "hazop", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_hazop_guidewords(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        #endregion export hazop

        #region export what's if 

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_template_whatif", Name = "export_template_whatif")]
        public string export_template_whatif(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_template_whatif", "whatif", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_whatif_report", Name = "export_whatif_report")]
        public string export_whatif_report(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_whatif_report", "whatif", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));


        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_whatif_worksheet", Name = "export_whatif_worksheet")]
        public string export_whatif_worksheet(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_whatif_worksheet", "whatif", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_whatif_recommendation", Name = "export_whatif_recommendation")]
        public string export_whatif_recommendation(ReportByWorksheetModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_whatif_recommendation", "whatif", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_report_recommendation(param, false, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_whatif_ram", Name = "export_whatif_ram")]
        public string export_whatif_ram(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_whatif_ram", "whatif", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_template_ram(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        #endregion export what's if

        #region export jsea

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_template_jsea", Name = "export_template_jsea")]
        public string export_template_jsea(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_template_jsea", "jsea", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                //return cls.export_template_jsea(param);
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_jsea_report", Name = "export_jsea_report")]
        public string export_jsea_report(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_jsea_report", "jsea", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));


        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_jsea_worksheet", Name = "export_jsea_worksheet")]
        public string export_jsea_worksheet(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_jsea_worksheet", "jsea", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                //return cls.export_jsea_worksheet(param);
                return cls.export_full_report(param, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        #endregion export jsea

        #region export hra 

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_hra_report", Name = "export_hra_report")]
        public string export_hra_report(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_hra_report", "hra", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_hra_worksheet", Name = "export_hra_worksheet")]
        public string export_hra_worksheet(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_hra_worksheet", "hra", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_potential_health_checklist_template(param, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_hra_recommendation", Name = "export_hra_recommendation")]
        public string export_hra_recommendation(ReportByWorksheetModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_hra_recommendation", "hra", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_report_recommendation(param, false, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_hra_template_moc", Name = "export_hra_template_moc")]
        public string export_hra_template_moc(ReportModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_hra_template_moc", "hra", param, ref token_log);
            try
            {
                if (param != null)
                {
                    param.export_type = "template";

                    ClassExcel cls = new ClassExcel();
                    return cls.export_full_report(param, false);
                }
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_recommendation_by_action_owner", Name = "export_recommendation_by_action_owner")]
        public string export_recommendation_by_action_owner(ReportByWorksheetModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_recommendation_by_action_owner", "all", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                msg = cls.export_report_recommendation(param, true, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("export_recommendation_by_item", Name = "export_recommendation_by_item")]
        public string export_recommendation_by_item(ReportByWorksheetModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("export_recommendation_by_item", "all", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                msg = cls.export_report_recommendation(param, false, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        #endregion export hra

        #region Function Search

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("employees_search", Name = "employees_search")]
        public string employees_search(EmployeeModel param)
        {
            //param.max_rows = (param.max_rows == null ? "10" : param.max_rows); 
            ClassHazop cls = new ClassHazop();
            return cls.employees_search(param);

        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("employees_list", Name = "employees_list")]
        public string employees_list(EmployeeListModel param)
        {
            ClassHazop cls = new ClassHazop();
            return cls.employees_list(param);
        }
        #endregion  Function Search

        #region Function Manage Document

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("manage_document_copy", Name = "manage_document_copy")]
        public string manage_document_copy(ManageDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("manage_document_copy", "all", param, ref token_log);
            try
            {
                ClassManage cls = new ClassManage();
                string ret = cls.DocumentCopy(param);

                return ret;
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("manage_document_cancel", Name = "manage_document_cancel")]
        public string manage_document_cancel(ManageDocModel param)
        {
            string token_log = "";
            string msg = "";//InsertTransactionLog("manage_document_cancel", "all", param, ref token_log);
            try
            {
                ClassManage cls = new ClassManage();
                string ret = cls.DocumentCancel(param);

                return ret;
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        #endregion Function Manage Document

    }
}
