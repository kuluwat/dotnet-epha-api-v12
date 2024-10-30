
using Model;
using dotnet6_epha_api.Class;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;

using dotnet_epha_api.Class;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PdfSharpCore.Pdf.IO;
using PdfSharpCore.Pdf;
using System.Diagnostics;
using Microsoft.Exchange.WebServices.Data;
using System.Reflection.Metadata;
using Microsoft.AspNetCore.Http;

namespace Class
{
    public class ClassExcel
    {
        string server_url = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["WebServer_ePHA_Index"] ?? "";
        string sqlstr = "";
        string jsper = "";
        string ret = "";
        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb _conn = new ClassConnectionDb();

        string[] sMonth = ("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC").Split(',');

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
        #region export excel all
        static void MergeCell(ExcelWorksheet worksheet, string address_start, string address_end)
        {
            var startCell = worksheet.Cells[address_start];
            var endCell = worksheet.Cells[address_end];
            var mergeRange = worksheet.Cells[startCell.Address + ":" + endCell.Address];
            // Merge the cells
            mergeRange.Merge = true;
        }
        static void DrawTableBorders(ExcelWorksheet worksheet, int startRow, int startCol, int endRow, int endCol)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
            }
        }

        private void fnSplit_Data_No(string val, ref string no_ref, ref string desc_ref)
        {
            no_ref = "";
            desc_ref = "";
            //index ? 0 : ข้อมูล No,  1 : ข้อมูล Description
            try
            {
                //1.1.1. xxx
                string[] xSplit = val.Split(new string[] { ". " }, StringSplitOptions.None);
                if (xSplit.Length > 0)
                {

                    if (!(xSplit[0].ToString() == ""))
                    {
                        //string xval_no = xSplit[0].ToString() + ". ";
                        string xval_no = xSplit[0].ToString();

                        xSplit = val.Split(new string[] { xval_no }, StringSplitOptions.None);
                        if (xSplit.Length > 1)
                        {
                            string xref = "";
                            foreach (string item in xSplit)
                            {
                                if (xref != "") { xref += xval_no; }
                                xref += item;
                            }
                            desc_ref = xref;
                            no_ref = xval_no;
                        }
                    }
                }
            }
            catch { }
        }
        public int FindRowWithAttendees(ExcelWorksheet worksheet, string cellText, string searchText)
        {
            var cell = worksheet.Cells[cellText + ":" + cellText]
                .FirstOrDefault(c => c.Text == searchText);

            if (cell != null)
            {
                // พบข้อมูลที่ต้องการในเซลล์ที่ cell อ้างถึง
                return cell.Start.Row;
            }

            // ถ้าไม่พบข้อมูลที่ต้องการในเซลล์
            return 1; // หรือส่งค่าอื่นๆ ตามที่ต้องการ เช่น 0, -999 เป็นต้น
        }
        //Report 
        private void _delay_time(string filePath)
        {

            int maxRetryAttempts = 5;
            int retryDelayMilliseconds = 1000;

            for (int i = 0; i < maxRetryAttempts; i++)
            {
                try
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        // Your code to work with the file
                    }
                    break; // If successful, break out of the loop
                }
                catch (IOException ex)
                {
                    // Handle the exception or log it
                    //Console.WriteLine($"Attempt {i + 1}: {ex.Message}");
                    System.Threading.Thread.Sleep(retryDelayMilliseconds); // Wait before retrying
                }
            }
        }

        #endregion export excel all

        #region function all module 
        public string export_full_report(ReportModel param, Boolean report_all, Boolean res_fullpath = false)
        {
            string msg_error = "";
            DataSet _dsData = new DataSet();
            DataTable dtdef = (DataTable)ClassFile.DatatableFile();
            string _file_name = "";
            string _file_download_name = "";
            string _file_fullpath_name = "";
            try
            {
                if (param == null) { msg_error = $"The specified file does not param."; }
                else
                {
                    string seq = param?.seq ?? "";
                    string export_type = param?.export_type ?? "";
                    string sub_software = param.sub_software ?? "";

                    string file_part = (export_type == "template" ? "Template" : "Report");
                    if (export_type == "template") { export_type = "excel"; }

                    // ตรวจสอบ sub_software ว่ามีอยู่ใน whitelist
                    var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };
                    if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                    {
                        return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid sub_software."));
                    }
                    string folder = sub_software ?? "";

                    if (!string.IsNullOrEmpty(seq) && !string.IsNullOrEmpty(sub_software))
                    {
                        //copy template to new file report 
                        ClassFile.copy_file_excel_template(ref _file_name, ref _file_download_name, ref _file_fullpath_name, folder, file_part, "");
                        if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(_file_name) || string.IsNullOrEmpty(_file_download_name) || string.IsNullOrEmpty(_file_fullpath_name))
                        { msg_error = "Invalid folder."; }
                        else
                        {

                            if (!string.IsNullOrEmpty(_file_fullpath_name))
                            {
                                string file_fullpath_def = _file_fullpath_name;
                                // string folder = sub_software ?? "";
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
                                    _file_fullpath_name = fullPath;
                                }
                                else { _file_fullpath_name = ""; }

                            }

                            if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(_file_fullpath_name))
                            {
                                if (file_part == "Template")
                                {
                                    switch (sub_software.ToLower())
                                    {
                                        case "jsea":
                                            msg_error = excle_template_data_jsea(seq, _file_fullpath_name, true);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }
                                            break;

                                        case "whatif":
                                            msg_error = excle_template_data_whatif(seq, _file_fullpath_name, true);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }
                                            break;

                                        case "hra":
                                            msg_error = excel_potential_health_checklist_template(seq, _file_fullpath_name);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }
                                            break;
                                    }
                                }
                                else
                                {
                                    #region set data worksheet / import file data to database 
                                    switch (sub_software.ToLower())
                                    {
                                        case "hazop":

                                            //Study Objective and Work Scope, Drawing & Reference, Node List
                                            msg_error = excel_hazop_general(seq, _file_fullpath_name);
                                            if (!string.IsNullOrEmpty(msg_error)) { msg_error = $"excel_hazop_general:{msg_error}"; goto Next_Line; }

                                            //HAZOP Attendee Sheet 
                                            msg_error = excel_hazop_atendeesheet(seq, _file_fullpath_name);
                                            if (!string.IsNullOrEmpty(msg_error)) { msg_error = $"excel_hazop_atendeesheet:{msg_error}"; goto Next_Line; }

                                            //HAZOP Recommendation
                                            msg_error = excel_hazop_recommendation(seq, _file_fullpath_name, true, "", "");
                                            if (!string.IsNullOrEmpty(msg_error)) { msg_error = $"excel_hazop_recommendation:{msg_error}"; goto Next_Line; }

                                            msg_error = excel_hazop_worksheet(seq, _file_fullpath_name, true);
                                            if (!string.IsNullOrEmpty(msg_error)) { msg_error = $"excel_hazop_worksheet:{msg_error}"; goto Next_Line; }

                                            msg_error = excel_hazop_guidewords(seq, _file_fullpath_name, true);
                                            if (!string.IsNullOrEmpty(msg_error)) { msg_error = $"excel_hazop_guidewords:{msg_error}"; goto Next_Line; }

                                            break;
                                        case "jsea":

                                            //Study Objective and Work Scope, Drawing & Reference, Node List
                                            msg_error = excel_jsea_general(seq, _file_fullpath_name);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }

                                            //JSEA Attendee Sheet 
                                            msg_error = excel_jsea_atendeesheet(seq, _file_fullpath_name);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }

                                            msg_error = excel_jsea_worksheet(seq, _file_fullpath_name, true);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }
                                            break;
                                        case "whatif":

                                            //Study Objective and Work Scope, Drawing & Reference, Node List
                                            msg_error = excel_whatif_general(seq, _file_fullpath_name);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }

                                            //Whatif Attendee Sheet 
                                            msg_error = excel_whatif_atendeesheet(seq, _file_fullpath_name);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }

                                            //Whatif Recommendation
                                            msg_error = excel_whatif_recommendation(seq, _file_fullpath_name, true, "", "");
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }

                                            msg_error = excel_whatif_worksheet(seq, _file_fullpath_name, true);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }

                                            break;
                                        case "hra":

                                            //Cover Page, Drawing & Reference 
                                            msg_error = excel_hra_general(seq, _file_fullpath_name);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }

                                            //Recommendation, ,HRA Worksheet and Associate Worksheet
                                            msg_error = excel_hra_worksheet(seq, _file_fullpath_name, true, false);
                                            if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }

                                            break;
                                    }

                                    if (folder.ToLower().Contains("hazop") || folder.ToLower().Contains("jsea") || folder.ToLower().Contains("whatif"))
                                    {
                                        //msg_error = excel_master_ram(seq, _file_fullpath_name, true, folder);
                                        msg_error = excel_master_ram(seq, _file_fullpath_name, true);
                                        if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }
                                    }

                                    if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(_file_fullpath_name))
                                    {
                                        try
                                        {
                                            FileInfo template = new FileInfo(_file_fullpath_name);
                                            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                                            using (ExcelPackage excelPackage = new ExcelPackage(template))
                                            {
                                                string SheetName_befor = excelPackage.Workbook.Worksheets[excelPackage.Workbook.Worksheets.Count - 1].Name;
                                                string SheetName = "Drawing PIDs & PFDs";

                                                excelPackage.Workbook.Worksheets.MoveAfter(SheetName, SheetName_befor);

                                                // Save changes
                                                excelPackage.Save();
                                            }
                                        }
                                        catch (Exception ex_error) { msg_error = ex_error.Message.ToString(); }
                                        if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }
                                    }


                                    #endregion set data worksheet / import file data to database 
                                }

                                // ตรวจสอบการมีอยู่ของไดเรกทอรี wwwroot
                                string templateWwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                                if (!Directory.Exists(templateWwwRootDir))
                                {
                                    return "Folder directory 'wwwroot' not found.";
                                }

                                // ตรวจสอบการมีอยู่ของไดเรกทอรี AttachedFileTemp
                                string templateRootDir = Path.Combine(templateWwwRootDir, "AttachedFileTemp");
                                if (!Directory.Exists(templateRootDir))
                                {
                                    return "Folder directory 'AttachedFileTemp' not found.";
                                }

                                // ตรวจสอบการมีอยู่ของไดเรกทอรี module (ที่ถูกอ้างอิงจาก folder)
                                string moduleFolderPath = Path.Combine(templateRootDir, folder);
                                if (!Directory.Exists(moduleFolderPath))
                                {
                                    //return $"Folder directory '{folder}' not found.";
                                    return $"Folder directory folder not found.";
                                }

                                // ตรวจสอบการมีอยู่ของ LibreOffice ใน tools directory
                                string libreOfficePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "LibreOffice", "program", "soffice.exe");
                                if (!File.Exists(libreOfficePath))
                                {
                                    return "LibreOffice executable not found in project directory.";
                                }

                                if (export_type == "pdf")
                                {
                                    try
                                    {
                                        FileInfo template = new FileInfo(_file_fullpath_name);
                                        if (!template.Exists || template.IsReadOnly)
                                        {
                                            if (template.IsReadOnly)
                                            {
                                                template.IsReadOnly = false;
                                            }
                                            else
                                            {
                                                msg_error = "File permissions are not correctly set.";
                                                goto Next_Line;
                                            }
                                        }

                                        // ใช้ LibreOffice ในการแปลงไฟล์
                                        string _file_fullpath_name_pdf = _file_fullpath_name.Replace(".xlsx", ".pdf");

                                        var process = new System.Diagnostics.Process();
                                        process.StartInfo.FileName = libreOfficePath;
                                        process.StartInfo.Arguments = $"--headless --convert-to pdf --outdir \"{Path.GetDirectoryName(_file_fullpath_name)}\" \"{_file_fullpath_name}\"";
                                        process.StartInfo.RedirectStandardOutput = true;
                                        process.StartInfo.RedirectStandardError = true;
                                        process.StartInfo.UseShellExecute = false;
                                        process.StartInfo.CreateNoWindow = true;
                                        process.Start();

                                        bool exited = process.WaitForExit(30000); // รอให้กระบวนการทำงานเสร็จภายใน 30 วินาที (30000 มิลลิวินาที)
                                        if (!exited)
                                        {
                                            // ถ้ากระบวนการไม่เสร็จภายในเวลาที่กำหนด ให้บังคับปิด
                                            process.Kill();
                                            throw new Exception("Process timed out and was killed.");
                                        }

                                        if (process.ExitCode == 0)
                                        {
                                            // ตรวจสอบการสร้างไฟล์ PDF
                                            msg_error = ClassFile.check_file_other(_file_fullpath_name_pdf, ref _file_fullpath_name_pdf, ref _file_download_name, folder);

                                            if (string.IsNullOrEmpty(msg_error))
                                            {
                                                // เพิ่มไฟล์เข้าไปใน appendix
                                                msg_error = add_drawing_to_appendix(seq, _file_fullpath_name_pdf, folder);
                                            }
                                            else
                                            {
                                                msg_error = $"Failed to create PDF file: {_file_fullpath_name_pdf}";
                                            }
                                        }
                                        else
                                        {
                                            msg_error = "Failed to convert file using LibreOffice.";
                                        }

                                        // เปลี่ยนเส้นทาง fullpath filename
                                        _file_fullpath_name = _file_fullpath_name_pdf;
                                        if (!string.IsNullOrEmpty(_file_fullpath_name))
                                        {
                                            msg_error = ClassFile.check_format_file_name(_file_fullpath_name);
                                            if (!string.IsNullOrEmpty(msg_error))
                                            {
                                                msg_error = $"Failed to change path Excel to PDF file ";
                                            }
                                        }
                                        else
                                        {
                                            _file_name = (_file_name?.ToLower() ?? "").Replace(".xlsx", ".pdf");
                                            if (!string.IsNullOrEmpty(_file_name))
                                            {
                                                msg_error = ClassFile.check_format_file_name(_file_name);
                                                if (!string.IsNullOrEmpty(msg_error))
                                                {
                                                    msg_error = $"Failed to replace PDF file: {_file_name}";
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex_pdf)
                                    {
                                        msg_error = ex_pdf.Message.ToString();
                                    }
                                }

                            }

                        }
                    }

                }

            Next_Line:;

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }
            if (res_fullpath)
            {
                return _file_fullpath_name;
            }
            else
            {
                if (dtdef != null)
                {
                    ClassFile.AddRowToDataTable(ref dtdef, _file_name, _file_download_name, msg_error);
                    DataSet dsData = new DataSet();
                    _dsData.Tables.Add(dtdef.Copy());
                }
                return JsonConvert.SerializeObject(_dsData, Formatting.Indented);
            }
        }
        public string export_template_ram(ReportModel param)
        {
            string msg_error = "";
            DataSet _dsData = new DataSet();
            DataTable dtdef = (DataTable)ClassFile.DatatableFile();
            string _file_name = "";
            string _file_download_name = "";
            string _file_fullpath_name = "";

            try
            {
                if (param == null) { msg_error = $"The specified file does not param."; }
                else
                {
                    string seq = param?.seq ?? "";
                    string export_type = param?.export_type ?? "";
                    string sub_software = param.sub_software ?? "";
                    string user_name = (param?.user_name ?? "");

                    if (!string.IsNullOrEmpty(seq) && !string.IsNullOrEmpty(sub_software))
                    {
                        string folder = sub_software ?? "";
                        string file_part = (export_type == "template" ? "Template" : "Report");
                        if (export_type == "template") { export_type = "excel"; }

                        //copy template to new file report 
                        ClassFile.copy_file_excel_template(ref _file_name, ref _file_download_name, ref _file_fullpath_name, folder, file_part, "");
                        if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(_file_name) || string.IsNullOrEmpty(_file_download_name) || string.IsNullOrEmpty(_file_fullpath_name))
                        { msg_error = "Invalid folder."; }
                        else
                        {
                            //ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                            if (!File.Exists(_file_fullpath_name))
                            {
                                msg_error = $"The specified file does not exist.{_file_fullpath_name}";
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(_file_fullpath_name))
                                {
                                    string file_fullpath_def = _file_fullpath_name;
                                    // string folder = sub_software ?? "";
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
                                        _file_fullpath_name = fullPath;
                                    }
                                    else { _file_fullpath_name = ""; }
                                }

                                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(_file_fullpath_name))
                                {
                                    //msg_error = excel_master_ram(seq, _file_fullpath_name, false, folder);
                                    msg_error = excel_master_ram(seq, _file_fullpath_name, false);
                                    if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }

                                    // Save the workbook as PDF
                                    // ตรวจสอบการมีอยู่ของ LibreOffice ใน tools directory
                                    string libreOfficePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "LibreOffice", "program", "soffice.exe");
                                    if (!File.Exists(libreOfficePath))
                                    {
                                        return "LibreOffice executable not found in project directory.";
                                    }

                                    if (export_type == "pdf")
                                    {
                                        try
                                        {
                                            FileInfo template = new FileInfo(_file_fullpath_name);
                                            if (!template.Exists || template.IsReadOnly)
                                            {
                                                if (template.IsReadOnly)
                                                {
                                                    template.IsReadOnly = false;
                                                }
                                                else
                                                {
                                                    msg_error = "File permissions are not correctly set.";
                                                    goto Next_Line;
                                                }
                                            }

                                            // ใช้ LibreOffice ในการแปลงไฟล์
                                            string _file_fullpath_name_pdf = _file_fullpath_name.Replace(".xlsx", ".pdf");

                                            var process = new System.Diagnostics.Process();
                                            process.StartInfo.FileName = libreOfficePath;
                                            process.StartInfo.Arguments = $"--headless --convert-to pdf --outdir \"{Path.GetDirectoryName(_file_fullpath_name)}\" \"{_file_fullpath_name}\"";
                                            process.StartInfo.CreateNoWindow = true;
                                            process.StartInfo.UseShellExecute = false;
                                            process.StartInfo.RedirectStandardOutput = true;
                                            process.Start();


                                            bool exited = process.WaitForExit(30000); // รอให้กระบวนการทำงานเสร็จภายใน 30 วินาที (30000 มิลลิวินาที)
                                            if (!exited)
                                            {
                                                // ถ้ากระบวนการไม่เสร็จภายในเวลาที่กำหนด ให้บังคับปิด
                                                process.Kill();
                                                throw new Exception("Process timed out and was killed.");
                                            }

                                            if (process.ExitCode == 0)
                                            {
                                                // ตรวจสอบการสร้างไฟล์ PDF
                                                msg_error = ClassFile.check_file_other(_file_fullpath_name_pdf, ref _file_fullpath_name_pdf, ref _file_download_name, folder);

                                                if (string.IsNullOrEmpty(msg_error))
                                                {
                                                    // เพิ่มไฟล์เข้าไปใน appendix
                                                    msg_error = add_drawing_to_appendix(seq, _file_fullpath_name_pdf, folder);
                                                }
                                                else
                                                {
                                                    msg_error = $"Failed to create PDF file: {_file_fullpath_name_pdf}";
                                                }
                                            }
                                            else
                                            {
                                                msg_error = "Failed to convert file using LibreOffice.";
                                            }

                                            // เปลี่ยนเส้นทาง fullpath filename
                                            _file_fullpath_name = _file_fullpath_name_pdf;
                                            if (!string.IsNullOrEmpty(_file_fullpath_name))
                                            {
                                                msg_error = ClassFile.check_format_file_name(_file_fullpath_name);
                                                if (!string.IsNullOrEmpty(msg_error))
                                                {
                                                    msg_error = $"Failed to change path Excel to PDF file ";
                                                }
                                            }
                                            else
                                            {
                                                _file_name = (_file_name?.ToLower() ?? "").Replace(".xlsx", ".pdf");
                                                if (!string.IsNullOrEmpty(_file_name))
                                                {
                                                    msg_error = ClassFile.check_format_file_name(_file_name);
                                                    if (!string.IsNullOrEmpty(msg_error))
                                                    {
                                                        msg_error = $"Failed to replace PDF file: {_file_name}";
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex_pdf)
                                        {
                                            msg_error = ex_pdf.Message.ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }

                Next_Line:;
                }
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            if (dtdef != null)
            {
                ClassFile.AddRowToDataTable(ref dtdef, _file_name, _file_download_name, msg_error);
                DataSet dsData = new DataSet();
                _dsData.Tables.Add(dtdef.Copy());
            }
            return JsonConvert.SerializeObject(_dsData, Formatting.Indented);
        }
        public string add_drawing_to_appendix(string seq, string file_fullpath_name, string sub_software)
        {
            //file_fullpath_name -> *.pdf only
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            if (string.IsNullOrEmpty(file_fullpath_name)) { return "Invalid File."; }
            if (string.IsNullOrEmpty(sub_software)) { return "Invalid Module."; }

            // ตรวจสอบ sub_software ว่ามีอยู่ใน whitelist
            var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };
            if (!allowedSubSoftware.Contains(sub_software.ToLower()))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid sub_software."));
            }

            string msg_error = "";

            #region Get Data 
            //Drawing PIDs & PFDs 
            sqlstr = @" select distinct d.no, d.document_name, d.document_no, d.document_file_name, d.descriptions, h.pha_sub_software as sub_software
                        , d.document_file_path
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                        inner join EPHA_T_DRAWING d on h.id = d.id_pha    
                        where h.seq = @seq and isnull(d.document_file_name,'') <>'' order by convert(int,d.no) ";

            List<SqlParameter> parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
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
                    dtDrawing = new DataTable();
                    dtDrawing = _conn.ExecuteAdapter(command).Tables[0];
                    dtDrawing.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            #endregion Get Data

            if (dtDrawing != null)
            {
                if (dtDrawing?.Rows.Count > 0)
                {
                    //File ต้นทาง
                    string sourceFilePath = file_fullpath_name;
                    if (!string.IsNullOrEmpty(sourceFilePath))
                    {
                        string file_fullpath_def = sourceFilePath;
                        string folder = sub_software ?? "";

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
                            sourceFilePath = fullPath;
                        }
                        else { sourceFilePath = ""; }
                    }
                    if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(sourceFilePath))
                    {

                        for (int i = 0; i < dtDrawing?.Rows.Count; i++)
                        {
                            //https://localhost:7098/AttachedFileTemp/hazop/HAZOP-2024-0000004-DRAWING-202401102132.PDF
                            string _document_file_name = "";
                            try
                            {
                                //string[] xfile = (dtDrawing.Rows[i]["document_file_path"] + "").Split("/");
                                //_document_file_name = xfile[xfile.Length - 1];
                                _document_file_name = dtDrawing.Rows[i]["document_file_path"]?.ToString() ?? "";
                            }
                            catch { }

                            if (!string.IsNullOrEmpty(_document_file_name))
                            {
                                string pdfToBeAddedFilePath = "";
                                string file_download_name = "";
                                msg_error = ClassFile.check_file_other(_document_file_name, ref pdfToBeAddedFilePath, ref file_download_name, "");

                                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(pdfToBeAddedFilePath))
                                {

                                    if (!string.IsNullOrEmpty(pdfToBeAddedFilePath))
                                    {
                                        string file_fullpath_def = pdfToBeAddedFilePath;
                                        string folder = sub_software ?? "";

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
                                            pdfToBeAddedFilePath = fullPath;
                                        }
                                        else { pdfToBeAddedFilePath = ""; }


                                        if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(sourceFilePath) && !string.IsNullOrEmpty(pdfToBeAddedFilePath))
                                        {
                                            //// Create a FileStream for the output PDF 
                                            //using (FileStream outputStream = new FileStream(sourceFilePath, FileMode.Create))
                                            //{
                                            //    //// Create a Document object
                                            //    //using (iTextSharp.text.Document document = new iTextSharp.text.Document())
                                            //    //{
                                            //    //    // Create a PdfCopy object that will merge the PDFs
                                            //    //    using (PdfCopy copy = new PdfCopy(document, outputStream))
                                            //    //    {
                                            //    //        document.Open();

                                            //    //        // Open the PDF to be added
                                            //    //        using (PdfReader pdfToBeAddedReader = new PdfReader(pdfToBeAddedFilePath))
                                            //    //        {
                                            //    //            // Add the pages from the PDF to be added to the new PDF
                                            //    //            for (int pageNum = 1; pageNum <= pdfToBeAddedReader.NumberOfPages; pageNum++)
                                            //    //            {
                                            //    //                copy.AddPage(copy.GetImportedPage(pdfToBeAddedReader, pageNum));
                                            //    //            }
                                            //    //        }
                                            //    //    }
                                            //    //} 
                                            //}

                                            // สร้าง PdfDocument ใหม่สำหรับการรวมไฟล์ PDF
                                            using (PdfDocument outputDocument = new PdfDocument())
                                            {
                                                // เปิด PDF ไฟล์แรกเพื่อเพิ่มหน้า PDF เข้าไป
                                                using (PdfDocument pdfToBeAdded = PdfReader.Open(pdfToBeAddedFilePath, PdfDocumentOpenMode.Import))
                                                {
                                                    // เพิ่มหน้าจากไฟล์ PDF ที่ต้องการนำเข้ามาใน outputDocument
                                                    for (int pageNum = 0; pageNum < pdfToBeAdded.PageCount; pageNum++)
                                                    {
                                                        // เพิ่มหน้าทีละหน้า
                                                        outputDocument.AddPage(pdfToBeAdded.Pages[pageNum]);
                                                    }
                                                }

                                                // บันทึก PDF ไฟล์ที่รวมแล้ว
                                                outputDocument.Save(sourceFilePath);
                                            }
                                        }
                                    }

                                }
                            }
                        }

                    }

                }
            }
            return msg_error;
        }
        public string excel_master_ram(string seq, string file_fullpath_name, Boolean report_all)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            //// ตรวจสอบ sub_software ว่ามีอยู่ใน whitelist
            //var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };
            //if (!allowedSubSoftware.Contains(pha_sub_software.ToLower()))
            //{
            //    return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid sub_software."));
            //}

            string msg_error = "";

            try
            {
                string sub_software = "";
                List<SqlParameter> parameters = new List<SqlParameter>();
                sqlstr = @" select distinct a.name as ram_type, a.descriptions, a.document_file_name, h.pha_sub_software 
                             from epha_m_ram a 
                             inner join epha_t_general g on a.id = g.id_ram
                             inner join epha_t_header h on h.id = g.id_pha
                             where a.active_type = 1 and g.id_pha = @seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
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


                if (dt != null) { return ""; }
                if (dt?.Rows.Count == 0) { return ""; }
                else
                {
                    sub_software = dt?.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                }
                if (!string.IsNullOrEmpty(sub_software)) { return "Invalid sub_software."; }

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = sub_software ?? "";
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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                    {
                        if (!report_all)
                        {
                            var sheetsToDelete = excelPackage.Workbook.Worksheets
                              .Where(sheet => sheet.Name != "Risk Assessment Matrix")
                              .ToList(); // ใช้ ToList เพื่อหลีกเลี่ยงปัญหา Collection Modified

                            foreach (var sheet in sheetsToDelete)
                            {
                                excelPackage.Workbook.Worksheets.Delete(sheet);
                            }
                        }

                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Risk Assessment Matrix"];
                        // Define picture dimensions and position
                        int left = 2;     // column index
                        int top = 2;      // row index
                        int width = 400;  // width in pixels
                        int height = 400; // height in pixels 

                        // Define the picture file path 
                        string _file_fullpath_name = "";
                        string _file_download_name = "";
                        string pictureFilePath = "";
                        if (dt != null)
                        {
                            if (dt?.Rows.Count > 0) { pictureFilePath = dt.Rows[0]["document_file_name"]?.ToString() ?? ""; }
                        }

                        if (!string.IsNullOrEmpty(pictureFilePath))
                        {
                            msg_error = ClassFile.check_file_other(pictureFilePath, ref file_fullpath_name, ref _file_download_name, "");
                            if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                            {
                                if (!string.IsNullOrEmpty(_file_fullpath_name))
                                {
                                    string file_fullpath_def = file_fullpath_name;
                                    string folder = sub_software ?? "";


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

                                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                                {
                                    // Insert the picture
                                    var picture = worksheet.Drawings.AddPicture("RAM", new FileInfo(file_fullpath_name));
                                    picture.From.Column = left;
                                    picture.From.Row = top;

                                    if ((dt.Rows[0]["ram_type"] + "") == "4x4") { width = 300; height = 300; }
                                    else if ((dt.Rows[0]["ram_type"] + "") == "5x5") { width = 400; height = 400; }
                                    else if ((dt.Rows[0]["ram_type"] + "") == "6x6") { width = 400; height = 400; }
                                    else if ((dt.Rows[0]["ram_type"] + "") == "7x7") { width = 400; height = 400; }
                                    else if ((dt.Rows[0]["ram_type"] + "") == "8x8") { width = 400; height = 400; }

                                    picture.SetSize(width, height);

                                    //descriptions  
                                    int startRows = 27;
                                    worksheet.Cells["A" + (startRows)].Value = dt.Rows[0]["descriptions"].ToString();

                                    excelPackage.Save();
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }

        public string copy_pdf_file(CopyFileModel param)
        {
            string msg_error = "";
            DataSet _dsData = new DataSet();
            DataTable dtdef = (DataTable)ClassFile.DatatableFile();
            string _file_name = "";
            string _file_download_name = "";
            string _file_fullpath_name = "";

            if (param == null) { msg_error = $"The specified file does not param."; goto Next_Line; }
            try
            {
                string sub_software = param.sub_software ?? "";
                string file_name = param.file_name ?? "";
                string file_path = param.file_path ?? "";
                string page_start_first = (param.page_start_first == null ? "" : param.page_start_first).Replace("null", "");
                string page_start_second = (param.page_start_second == null ? "" : param.page_start_second).Replace("null", "");
                string page_end_first = (param.page_end_first == null ? "" : param.page_end_first).Replace("null", "");
                string page_end_second = (param.page_end_second == null ? "" : param.page_end_second).Replace("null", "");

                // @"D:\dotnet6-epha-api\dotnet6-epha-api\wwwroot\AttachedFileTemp\Hazop\ebook_def.pdf";  // Replace with the path to the source PDF
                // @"D:\dotnet6-epha-api\dotnet6-epha-api\wwwroot\AttachedFileTemp\Hazop\ebook_v1.pdf"; // Replace with the path to the target PDF

                if (!string.IsNullOrEmpty(file_name) && !string.IsNullOrEmpty(sub_software))
                {
                    string folder = sub_software ?? "";
                    if (string.IsNullOrEmpty(file_name))
                    {
                        msg_error = "Invalid folder.";
                    }
                    else
                    {
                        msg_error = ClassFile.copy_file_duplicate(file_name, ref _file_name, ref _file_download_name, ref _file_fullpath_name, folder);
                        if (string.IsNullOrEmpty(msg_error) || string.IsNullOrEmpty(_file_name) || string.IsNullOrEmpty(_file_fullpath_name))
                        {
                            msg_error = "Invalid file name.";
                        }
                        else
                        {
                            //if (!string.IsNullOrEmpty(_file_name))
                            //if (!string.IsNullOrEmpty(_file_fullpath_name))

                            if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(_file_name))
                            {
                                string file_fullpath_def = _file_name;
                                // string folder = sub_software ?? "";


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
                                    _file_name = fullPath;
                                }
                                else { _file_name = ""; }
                            }
                            if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(_file_name) && !string.IsNullOrEmpty(_file_fullpath_name))
                            {
                                string file_fullpath_def = _file_fullpath_name;
                                // string folder = sub_software ?? "";


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
                                    _file_fullpath_name = fullPath;
                                }
                                else { _file_fullpath_name = ""; }
                            }

                            //// ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                            //if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(_file_name) && !string.IsNullOrEmpty(_file_fullpath_name))
                            //{
                            //    //using (var sourcePdfReader = new PdfReader(_file_download_name))
                            //    using (var sourcePdfReader = new PdfReader(_file_name))//file name = full path name 
                            //    using (var targetPdfStream = new FileStream(_file_fullpath_name, FileMode.Open, FileAccess.ReadWrite))
                            //    using (var targetPdfDoc = new iTextSharp.text.Document())
                            //    using (var targetPdfWriter = new PdfCopy(targetPdfDoc, targetPdfStream))
                            //    {
                            //        targetPdfDoc.Open();

                            //        int startPagePart1 = (page_start_first == "" ? 1 : Convert.ToInt32(page_start_first));  // Replace with the start page number
                            //        int endPagePart1 = (page_end_first == "" ? 100 : Convert.ToInt32(page_end_first)); ;    // Replace with the end page number
                            //        int startPagePart2 = (page_start_second == "" ? 0 : Convert.ToInt32(page_start_second)); ;  // Replace with the start page number
                            //        int endPagePart2 = (page_end_second == "" ? 0 : Convert.ToInt32(page_end_second)); ;    // Replace with the end page number

                            //        if (startPagePart1 > 0)
                            //        {
                            //            for (int pageNumber = startPagePart1; pageNumber <= endPagePart1; pageNumber++)
                            //            {
                            //                var page = targetPdfWriter.GetImportedPage(sourcePdfReader, pageNumber);
                            //                targetPdfWriter.AddPage(page);
                            //            }
                            //        }
                            //        if (startPagePart2 > 0)
                            //        {
                            //            for (int pageNumber = startPagePart2; pageNumber <= endPagePart2; pageNumber++)
                            //            {
                            //                var page = targetPdfWriter.GetImportedPage(sourcePdfReader, pageNumber);
                            //                targetPdfWriter.AddPage(page);
                            //            }
                            //        }

                            //        targetPdfDoc.Close();
                            //        msg_error = "";
                            //    }
                            //}
                            // ตรวจสอบว่าชื่อไฟล์และเส้นทางถูกต้อง
                            if (!string.IsNullOrEmpty(_file_name) && !string.IsNullOrEmpty(_file_fullpath_name))
                            {
                                // เปิด PDF ไฟล์ต้นฉบับ
                                using (PdfDocument sourcePdf = PdfReader.Open(_file_name, PdfDocumentOpenMode.Import))
                                {
                                    // สร้าง PDF ไฟล์ปลายทาง
                                    using (PdfDocument targetPdf = new PdfDocument())
                                    {
                                        int startPagePart1 = (page_start_first == "" ? 1 : Convert.ToInt32(page_start_first));  // Replace with the start page number
                                        int endPagePart1 = (page_end_first == "" ? 100 : Convert.ToInt32(page_end_first)); ;    // Replace with the end page number
                                        int startPagePart2 = (page_start_second == "" ? 0 : Convert.ToInt32(page_start_second)); ;  // Replace with the start page number
                                        int endPagePart2 = (page_end_second == "" ? 0 : Convert.ToInt32(page_end_second)); ;    // Replace with the end page number

                                        // ส่วนที่ 1: คัดลอกหน้าจาก startPagePart1 ถึง endPagePart1
                                        if (startPagePart1 > 0)
                                        {
                                            for (int pageNumber = startPagePart1 - 1; pageNumber < endPagePart1; pageNumber++)
                                            {
                                                if (pageNumber < sourcePdf.PageCount)
                                                {
                                                    // เพิ่มหน้าจากไฟล์ต้นฉบับไปยังไฟล์ปลายทาง
                                                    targetPdf.AddPage(sourcePdf.Pages[pageNumber]);
                                                }
                                            }
                                        }

                                        // ส่วนที่ 2: คัดลอกหน้าจาก startPagePart2 ถึง endPagePart2
                                        if (startPagePart2 > 0)
                                        {
                                            for (int pageNumber = startPagePart2 - 1; pageNumber < endPagePart2; pageNumber++)
                                            {
                                                if (pageNumber < sourcePdf.PageCount)
                                                {
                                                    // เพิ่มหน้าจากไฟล์ต้นฉบับไปยังไฟล์ปลายทาง
                                                    targetPdf.AddPage(sourcePdf.Pages[pageNumber]);
                                                }
                                            }
                                        }

                                        // บันทึกไฟล์ PDF ที่รวมแล้ว
                                        targetPdf.Save(_file_fullpath_name);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }
        Next_Line:;

            if (dtdef != null)
            {
                ClassFile.AddRowToDataTable(ref dtdef, _file_name, _file_fullpath_name, msg_error);
                DataSet dsData = new DataSet();
                _dsData.Tables.Add(dtdef.Copy());
            }
            return JsonConvert.SerializeObject(_dsData, Formatting.Indented);
        }

        #endregion function all module

        #region export excel hazop

        public string excel_hazop_general(string seq, string file_fullpath_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            try
            {
                #region get data
                List<SqlParameter> parameters = new List<SqlParameter>();
                sqlstr = @" select g.work_scope, h.pha_no, h.pha_version_text as pha_version, g.descriptions, format(g.target_start_date, 'dd MMM yyyy ') as target_start_date 
                         ,  case when ums.user_name is null  then h.request_user_displayname else case when ums.departments is null  then  ums.user_displayname else  ums.user_displayname + ' (' + ums.departments +')' end end request_user_displayname
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         left join VW_EPHA_PERSON_DETAILS ums on lower(h.request_user_name) = lower(ums.USER_NAME)
                         where h.seq = @seq ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtWorkScope = new DataTable();
                //dtWorkScope = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorkScope = new DataTable();
                        dtWorkScope = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorkScope.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable



                sqlstr = @" select distinct d.no, d.document_name, d.document_no, d.document_file_name, d.descriptions 
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                        inner join EPHA_T_DRAWING d on h.id = d.id_pha    
                        where h.seq = @seq and d.document_name is not null order by convert(int,d.no) ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
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
                        dtDrawing = new DataTable();
                        dtDrawing = _conn.ExecuteAdapter(command).Tables[0];
                        dtDrawing.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable



                sqlstr = @" select nl.no, nl.node, nl.design_intent, nl.design_conditions, nl.operating_conditions, nl.node_boundary
                         , d.document_no
                         , isnull(replace(replace( convert(char,nd.page_start_first) + (case when isnull(nd.page_start_first,'') ='' then '' else
                         (case when isnull(nd.page_end_first,'') ='' then '' else 'to'end)  end) 
                         + convert(char,nd.page_end_first)  ,' ',''),'to',' to '),'All') as  document_page
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                         left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                         left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                         where h.seq = @seq  and nl.node is not null order by convert(int,nl.no), convert(int,nd.no) ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtNode = new DataTable();
                //dtNode = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtNode = new DataTable();
                        dtNode = _conn.ExecuteAdapter(command).Tables[0];
                        dtNode.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                //MAJOR ACCIDENT EVENT  (Y/N) ให้ดึงที่เป็น Y -> running no , node , cause, R ของ  UNMITIGATED RISK ASSESSMENT MATRIX
                sqlstr = @" select 0 as no, nl.node, nw.causes, nw.causes_no, nw.ram_befor_risk
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_NODE nl on h.id = nl.id_pha  
                         left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node 
                         and lower(isnull(nw.major_accident_event,'')) = lower('Y') 
                         where h.seq = @seq order by convert(int,nw.no) ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtMajor = new DataTable();
                //dtMajor = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtMajor = new DataTable();
                        dtMajor = _conn.ExecuteAdapter(command).Tables[0];
                        dtMajor.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                #endregion get data

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "hazop";// sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                    {
                        ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];  // Replace "SourceSheet" with the actual source sheet name
                        ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                        //HAZOP Cover Page
                        worksheet = excelPackage.Workbook.Worksheets["HAZOP Cover Page"];

                        ClassHazop clshazop = new ClassHazop();
                        worksheet.Cells["A14"].Value = clshazop.convert_revision_text(dtWorkScope.Rows[0]["pha_version"] + "");
                        worksheet.Cells["B14"].Value = (dtWorkScope.Rows[0]["target_start_date"] + "");
                        worksheet.Cells["C14"].Value = (dtWorkScope.Rows[0]["request_user_displayname"] + "");
                        worksheet.Cells["D14"].Value = (dtWorkScope.Rows[0]["descriptions"] + "");

                        //Study Objective and Work Scope
                        worksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];
                        worksheet.Cells["A2"].Value = (dtWorkScope.Rows[0]["work_scope"] + "");

                        //Description/Remarks 
                        if ((dtWorkScope.Rows[0]["descriptions"] + "") == "")
                        {
                            //remove rows A4, A5
                            worksheet.DeleteRow(5); worksheet.DeleteRow(4);
                        }
                        else
                        {
                            worksheet.Cells["A5"].Value = (dtWorkScope.Rows[0]["descriptions"] + "");
                        }

                        //Drawing & Reference
                        #region Drawing & Reference
                        if (true)
                        {
                            worksheet = excelPackage.Workbook.Worksheets["Drawing & Reference"];

                            int startRows = 3;
                            int icol_end = 6;
                            int ino = 1;
                            for (int i = 0; i < dtDrawing?.Rows.Count; i++)
                            {
                                //No.	Document Name	Drawing No	Document File	Comment
                                worksheet.InsertRow(startRows, 1);
                                worksheet.Cells["A" + (startRows)].Value = (i + 1); ;
                                worksheet.Cells["B" + (startRows)].Value = (dtDrawing.Rows[i]["document_name"] + "");
                                worksheet.Cells["C" + (startRows)].Value = (dtDrawing.Rows[i]["document_no"] + "");
                                worksheet.Cells["D" + (startRows)].Value = (dtDrawing.Rows[i]["document_file_name"] + "");
                                worksheet.Cells["E" + (startRows)].Value = (dtDrawing.Rows[i]["descriptions"] + "");
                                startRows++;
                            }
                            // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                            DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                            //var eRange = worksheet.Cells[worksheet.Cells["A3"].Address + ":" + worksheet.Cells["D" + startRows].Address];
                            //eRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            //eRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        #endregion Drawing & Reference

                        //Node List
                        #region Node List
                        if (true)
                        {
                            worksheet = excelPackage.Workbook.Worksheets["Node List"];

                            int startRows = 3;
                            int icol_end = 9;
                            for (int i = 0; i < dtNode?.Rows.Count; i++)
                            {
                                //No.	Node	Design Intent	Design Conditions	Operating Conditions	Node Boundary	Drawing	Drawing Page (From-To)
                                worksheet.InsertRow(startRows, 1);
                                worksheet.Cells["A" + (startRows)].Value = (i + 1);
                                worksheet.Cells["B" + (startRows)].Value = (dtNode.Rows[i]["node"] + "");
                                worksheet.Cells["C" + (startRows)].Value = (dtNode.Rows[i]["design_intent"] + "");
                                worksheet.Cells["D" + (startRows)].Value = (dtNode.Rows[i]["design_conditions"] + "");
                                worksheet.Cells["E" + (startRows)].Value = (dtNode.Rows[i]["operating_conditions"] + "");
                                worksheet.Cells["F" + (startRows)].Value = (dtNode.Rows[i]["node_boundary"] + "");
                                worksheet.Cells["G" + (startRows)].Value = (dtNode.Rows[i]["document_no"] + "");
                                worksheet.Cells["H" + (startRows)].Value = (dtNode.Rows[i]["document_page"] + "");

                                startRows++;
                            }
                            // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                            DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);
                        }
                        #endregion Node List

                        // Major Accident Event (MAE),
                        #region Major Accident Event (MAE)
                        if (true)
                        {
                            worksheet = excelPackage.Workbook.Worksheets["Major Accident Event (MAE)"];

                            int startRows = 3;
                            int icol_end = 4;
                            for (int i = 0; i < dtMajor?.Rows.Count; i++)
                            {
                                //No.	 nl.node, nw.causes, nw.causes_no, nw.ram_befor_risk
                                worksheet.InsertRow(startRows, 1);
                                worksheet.Cells["A" + (startRows)].Value = (i + 1);
                                worksheet.Cells["B" + (startRows)].Value = (dtMajor.Rows[i]["node"] + "");
                                worksheet.Cells["C" + (startRows)].Value = (dtMajor.Rows[i]["causes"] + "");
                                worksheet.Cells["D" + (startRows)].Value = (dtMajor.Rows[i]["ram_befor_risk"] + "");

                                startRows++;
                            }
                            // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                            DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);
                        }
                        #endregion Node List

                        //Study Objective and Work Scope
                        #region Study Objective and Work Scope
                        if (true)
                        {
                            worksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];
                            worksheet.Cells["A2"].Value = (dtWorkScope.Rows[0]["work_scope"] + "");

                            //Description/Remarks 
                            worksheet.Cells["A5"].Value = (dtWorkScope.Rows[0]["descriptions"] + "");
                        }
                        #endregion Study Objective and Work Scope

                        excelPackage.Save();
                    }
                }

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }
            return msg_error;
        }
        public string excel_hazop_atendeesheet(string seq, string file_fullpath_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            try
            {
                List<SqlParameter> parameters = new List<SqlParameter>();
                sqlstr = @" select distinct a.id as id_pha, a.pha_status, ta2.no
                     , isnull(emp.user_name,'') as user_name, emp.user_displayname, emp.user_email
                     , case when ta2.action_review = 2 then (case when ta2.action_status = 'approve' then 'Approve' else
						(case when ta2.action_status = 'reject' and ta2.comment is not null  then 'Send Back' else 'Send Back with comment' end) 
					  end) else '' end action_status
					  
                     from epha_t_header a  
                     inner join EPHA_T_GENERAL g on a.id = g.id_pha   
                     inner join EPHA_T_SESSION s on a.id = s.id_pha 
                     inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s2 on a.id = s2.id_pha and s.id = s2.id_session 
                     inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha ) t2 on a.id = t2.id_pha and s2.id_session = t2.id_session 
                     inner join EPHA_T_APPROVER ta2 on a.id = ta2.id_pha and t2.id_pha = ta2.id_pha and t2.id_session = ta2.id_session 
                     left join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name) 
					 inner join  (select max(seq)as seq, pha_no from epha_t_header group by pha_no) hm on a.seq = hm.seq and a.pha_no = hm.pha_no
                     where a.request_approver = 1 and a.seq = @seq  order by ta2.no";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtTA2 = new DataTable();
                //dtTA2 = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtTA2 = new DataTable();
                        dtTA2 = _conn.ExecuteAdapter(command).Tables[0];
                        dtTA2.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable



                sqlstr = $@"select t.* from  (
                         select s.id_pha, s.seq, s.seq as seq_session, s.no as session_no
                         , convert(varchar,s.meeting_date,106) as meeting_date
                         , mt.no as member_no, isnull(mt.user_name,'') as user_name, emp.user_displayname
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_SESSION s on h.id = s.id_pha 
                         left join EPHA_T_MEMBER_TEAM mt on h.id = mt. id_pha and mt.id_session = s.id
                         left join VW_EPHA_PERSON_DETAILS emp on lower(emp.user_name) = lower(mt.user_name)
                         where lower(mt.user_name) is not null
                         )t where t.seq = @seq  ";
                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtAll = new DataTable();
                //dtAll = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtAll = new DataTable();
                        dtAll = _conn.ExecuteAdapter(command).Tables[0];
                        dtAll.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                sqlstr = $@" select distinct 0 as no, t.user_name, t.user_displayname, '' as company_text from (
                         select s.id_pha, s.seq, s.seq as seq_session, s.no as session_no
                         , convert(varchar,s.meeting_date,106) as meeting_date
                         , mt.no as member_no, isnull(mt.user_name,'') as user_name, emp.user_displayname
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_SESSION s on h.id = s.id_pha 
                         left join EPHA_T_MEMBER_TEAM mt on h.id = mt. id_pha and mt.id_session = s.id
                         left join VW_EPHA_PERSON_DETAILS emp on lower(emp.user_name) = lower(mt.user_name)
                         where lower(mt.user_name) is not null
                         )t where t.seq = @seq and t.user_name <> '' order by t.user_name";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtMember = new DataTable();
                //dtMember = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtMember = new DataTable();
                        dtMember = _conn.ExecuteAdapter(command).Tables[0];
                        dtMember.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = $@" select distinct t.seq_session, t.session_no, t.meeting_date from (
                         select s.id_pha, s.seq, s.seq as seq_session, s.no as session_no
                         , convert(varchar,s.meeting_date,106) as meeting_date
                         , mt.no as member_no, isnull(mt.user_name,'') as user_name, emp.user_displayname
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_SESSION s on h.id = s.id_pha 
                         left join EPHA_T_MEMBER_TEAM mt on h.id = mt. id_pha and mt.id_session = s.id
                         left join VW_EPHA_PERSON_DETAILS emp on lower(emp.user_name) = lower(mt.user_name)
                         where lower(mt.user_name) is not null
                         )t where t.seq = @seq order by t.session_no ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtSession = new DataTable();
                //dtSession = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtSession = new DataTable();
                        dtSession = _conn.ExecuteAdapter(command).Tables[0];
                        dtSession.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (dtAll != null)
                {
                    if (dtAll?.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(file_fullpath_name))
                        {
                            string file_fullpath_def = file_fullpath_name;
                            string folder = "hazop";// sub_software;


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


                        if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                        {
                            FileInfo template_excel = new FileInfo(file_fullpath_name);

                            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                            using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                            {
                                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];  // Replace "SourceSheet" with the actual source sheet name
                                sourceWorksheet.Name = "HAZOP Attendee Sheet";
                                ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                                int i = 0;
                                int startRows = 4;
                                int icol_start = 4;
                                int icol_end = 0;// icol_start + (dtSession.Rows.Count > 6 ? dtSession.Rows.Count : 6);
                                if (dtSession?.Rows.Count > 0)
                                {
                                    int icount_session = dtSession?.Rows.Count ?? 0;

                                    icol_end = icol_start + (icount_session > 6 ? icount_session : 6);
                                }

                                for (int imember = 0; imember < dtMember?.Rows.Count; imember++)
                                {
                                    worksheet.InsertRow(startRows, 1);
                                    string user_name = (dtMember.Rows[imember]["user_name"] + "");
                                    //No.
                                    worksheet.Cells["A" + (i + startRows)].Value = (imember + 1);
                                    //Name
                                    worksheet.Cells["B" + (i + startRows)].Value = (dtMember.Rows[imember]["user_displayname"] + "");
                                    //Company
                                    worksheet.Cells["C" + (i + startRows)].Value = (dtMember.Rows[imember]["company_text"] + "");

                                    int irow_session = 0;
                                    if (imember == 0)
                                    {
                                        if (dtSession?.Rows.Count < 6)
                                        {
                                            //worksheet.Cells[2, icol_start, 2, icol_end].Merge = true; 
                                            for (int c = icol_end; c < 30; c++)
                                            {
                                                worksheet.DeleteColumn(icol_end);

                                            }
                                        }

                                        irow_session = 0;
                                        for (int c = icol_start; c < icol_end; c++)
                                        {
                                            try
                                            {
                                                //header 
                                                if ((dtSession.Rows[irow_session]["meeting_date"] + "") == "")
                                                {
                                                    worksheet.Cells[3, c].Value = "";
                                                }
                                                else
                                                {
                                                    worksheet.Cells[3, c].Value = (dtSession.Rows[irow_session]["meeting_date"] + "");
                                                }
                                            }
                                            catch { worksheet.Cells[3, c].Value = ""; }
                                            irow_session += 1;
                                        }
                                    }

                                    irow_session = 0;
                                    for (int c = icol_start; c < icol_end; c++)
                                    {
                                        try
                                        {
                                            string session_no = "";
                                            try { session_no = (dtSession.Rows[irow_session]["session_no"] + ""); } catch { }
                                            worksheet.Cells[startRows, c].Value = "";

                                            //DataRow[] dr = dtAll.Select("user_name = '" + user_name + "' and session_no = '" + session_no + "'");
                                            var filterParameters = new Dictionary<string, object>();
                                            filterParameters.Add("user_name", user_name);
                                            filterParameters.Add("session_no", session_no);
                                            var (dr, iMerge) = FilterDataTable(dtAll, filterParameters);
                                            if (dr != null)
                                            {
                                                if (dr?.Length > 0)
                                                {
                                                    worksheet.Cells[startRows, c].Value = "X";
                                                }
                                            }
                                        }
                                        catch { }
                                        irow_session++;

                                    }

                                    startRows++;
                                }
                                // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                                DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                                //TA2 
                                startRows += 13;
                                int startRows_ta2 = startRows;

                                for (int ita2 = 0; ita2 < dtTA2?.Rows.Count; ita2++)
                                {
                                    worksheet.InsertRow(startRows, 1);
                                    //No.
                                    worksheet.Cells["A" + (i + startRows)].Value = (ita2 + 1);
                                    //Name
                                    worksheet.Cells["B" + (i + startRows)].Value = (dtTA2.Rows[ita2]["user_displayname"] + "");
                                    //status
                                    worksheet.Cells["C" + (i + startRows)].Value = (dtTA2.Rows[ita2]["action_status"] + "");

                                    startRows++;
                                }
                                // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                                DrawTableBorders(worksheet, startRows_ta2, 1, startRows - 1, 3);

                                excelPackage.Save();

                            }
                        }
                    }
                }
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }
        public string excel_hazop_worksheet(string seq, string file_fullpath_name, Boolean report_all)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            try
            {
                List<SqlParameter> parameters = new List<SqlParameter>();

                sqlstr = @" select distinct nl.no, nl.id as id_node
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        where h.seq = @seq ";
                sqlstr += @" order by cast(nl.no as int)";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtNode = new DataTable();
                //dtNode = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtNode = new DataTable();
                        dtNode = _conn.ExecuteAdapter(command).Tables[0];
                        dtNode.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = @" select distinct
                        h.seq, nl.id as id_node, g.pha_request_name, convert(varchar,g.create_date,106) as create_date, nl.node, nl.design_intent, nl.descriptions as descriptions_worksheet, nl.design_conditions, nl.node_boundary, nl.operating_conditions
                        , d.document_no
                        , convert(varchar,mgw.no_guide_words) +'.'+ mgw.guide_words as guideword
                        , convert(varchar,mgw.no_deviations) +'.'+ mgw.deviations as deviation
                        , nw.causes, nw.consequences, nw.category_type, nw.ram_befor_security, nw.ram_befor_likelihood, nw.ram_befor_risk
                        , nw.major_accident_event, nw.safety_critical_equipment, nw.safety_critical_equipment_tag, nw.existing_safeguards, nw.ram_after_security, nw.ram_after_likelihood, nw.ram_after_risk
                        , nw.recommendations, nw.recommendations_no, nw.responder_user_displayname
                        , g.descriptions
                        , nl.no as node_no, nw.no, nw.causes_no, nw.consequences_no, nw.category_no
                        , mgw.no_guide_words, mgw.no_deviations 
                        , case when g.id_ram = 5 then 1 else 0 end show_cat
                        , h.safety_critical_equipment_show
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word    
                        where h.seq = @seq ";
                sqlstr += @" order by cast(nl.no as int),cast(nw.no as int), cast(nw.causes_no as int), cast(nw.consequences_no as int), cast(nw.category_no as int)";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtWorksheet = new DataTable();
                //dtWorksheet = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorksheet = new DataTable();
                        dtWorksheet = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorksheet.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = @"  select distinct
                         h.seq,nl.no as node_no,nl.node, 0 as no,  nw.safety_critical_equipment, nw.safety_critical_equipment_tag
                         , str(nw.consequences_no) + '.' + nw.consequences as consequences, isnull(nw.ram_befor_risk,'') as  ram_befor_risk
                          , h.safety_critical_equipment_show
                         from epha_t_header h
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_NODE nl on h.id = nl.id_pha  
                         left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                         left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word    
                         where h.seq = @seq ";
                sqlstr += @" order by cast(nl.no as int),nl.node, nw.safety_critical_equipment_tag  ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtSCE = new DataTable();
                //dtSCE = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtSCE = new DataTable();
                        dtSCE = _conn.ExecuteAdapter(command).Tables[0];
                        dtSCE.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "hazop";//sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);

                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                    {
                        Boolean sce_show = false;
                        if (report_all == false && dtSCE?.Rows.Count > 0)
                        {
                            if ((dtSCE.Rows[0]["safety_critical_equipment_show"]?.ToString() ?? "") != "0") { sce_show = true; }
                        }

                        string worksheet_name = "";
                        string worksheet_name_target = "";
                        for (int inode = 0; inode < dtNode?.Rows.Count; inode++)
                        {
                            if (worksheet_name_target == "") { worksheet_name_target = "WorksheetTemplate"; }
                            else { worksheet_name_target = "HAZOP Worksheet Node (" + (inode) + ")"; }
                            worksheet_name = "HAZOP Worksheet Node (" + (inode + 1) + ")";

                            ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets[(sce_show ? "WorksheetTemplateSCE" : "WorksheetTemplate")];  // Replace "SourceSheet" with the actual source sheet name
                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(worksheet_name, sourceWorksheet);


                            string id_node = (dtNode.Rows[inode]["id_node"] + "");

                            int i = 0;
                            int startRows = 3;

                            //DataRow[] dr = dtWorksheet.Select("id_node=" + id_node);
                            var filterParameters = new Dictionary<string, object>();
                            filterParameters.Add("id_node", id_node);
                            var (dr, iMerge) = FilterDataTable(dtWorksheet, filterParameters);
                            if (dr != null)
                            {
                                if (dr?.Length > 0)
                                {
                                    string show_cat = (dr[0]["show_cat"] + "");
                                    #region head text
                                    string cell_h_end = (sce_show == true ? "O" : "N");
                                    i = 0;
                                    //Project
                                    worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["pha_request_name"] + "");
                                    //NODE
                                    worksheet.Cells[cell_h_end + (i + startRows)].Value = (dr[0]["node"] + "");
                                    startRows++;

                                    //Design Intent :
                                    worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["design_intent"] + "");
                                    //System
                                    worksheet.Cells[cell_h_end + (i + startRows)].Value = (dr[0]["descriptions"] + "");
                                    startRows++;

                                    //"Design Conditions: -->design_conditions
                                    worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["design_conditions"] + "");
                                    //HAZOP Boundary
                                    worksheet.Cells[cell_h_end + (i + startRows)].Value = (dr[0]["node_boundary"] + "");
                                    startRows++;

                                    //"Operating Conditions: -->operating_conditions
                                    worksheet.Cells[cell_h_end + (i + startRows)].Value = (dr[0]["operating_conditions"] + "");
                                    startRows++;

                                    //PFD, PID No. : --> document_no
                                    worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["document_no"] + "");
                                    //Date
                                    worksheet.Cells[cell_h_end + (i + startRows)].Value = (dr[0]["create_date"] + "");
                                    startRows++;

                                    #endregion head text
                                    startRows = 14;
                                    for (i = 0; i < dr.Length; i++)
                                    {
                                        worksheet.InsertRow(startRows, 1);

                                        worksheet.Cells["A" + (startRows)].Value = dr[i]["guideword"].ToString();
                                        worksheet.Cells["B" + (startRows)].Value = dr[i]["deviation"].ToString();
                                        worksheet.Cells["C" + (startRows)].Value = dr[i]["causes"].ToString();
                                        worksheet.Cells["D" + (startRows)].Value = dr[i]["consequences"].ToString();
                                        worksheet.Cells["E" + (startRows)].Value = dr[i]["category_type"].ToString();

                                        worksheet.Cells["F" + (startRows)].Value = dr[i]["ram_befor_security"].ToString();
                                        worksheet.Cells["G" + (startRows)].Value = dr[i]["ram_befor_likelihood"].ToString();
                                        worksheet.Cells["H" + (startRows)].Value = dr[i]["ram_befor_risk"];
                                        worksheet.Cells["I" + (startRows)].Value = dr[i]["major_accident_event"].ToString();

                                        if (sce_show)
                                        {
                                            worksheet.Cells["J" + (startRows)].Value = dr[i]["safety_critical_equipment_tag"].ToString();
                                            worksheet.Cells["K" + (startRows)].Value = dr[i]["existing_safeguards"].ToString();

                                            worksheet.Cells["L" + (startRows)].Value = dr[i]["ram_after_security"].ToString();
                                            worksheet.Cells["M" + (startRows)].Value = dr[i]["ram_after_likelihood"].ToString();
                                            worksheet.Cells["N" + (startRows)].Value = dr[i]["ram_after_risk"].ToString();
                                            worksheet.Cells["O" + (startRows)].Value = dr[i]["recommendations_no"].ToString();
                                            worksheet.Cells["P" + (startRows)].Value = dr[i]["recommendations"].ToString();
                                            worksheet.Cells["Q" + (startRows)].Value = dr[i]["existing_safeguards"].ToString();
                                            worksheet.Cells["R" + (startRows)].Value = dr[i]["responder_user_displayname"].ToString();
                                        }
                                        else
                                        {
                                            worksheet.Cells["J" + (startRows)].Value = dr[i]["existing_safeguards"].ToString();

                                            worksheet.Cells["K" + (startRows)].Value = dr[i]["ram_after_security"].ToString();
                                            worksheet.Cells["L" + (startRows)].Value = dr[i]["ram_after_likelihood"].ToString();
                                            worksheet.Cells["M" + (startRows)].Value = dr[i]["ram_after_risk"].ToString();
                                            worksheet.Cells["N" + (startRows)].Value = dr[i]["recommendations_no"].ToString();
                                            worksheet.Cells["O" + (startRows)].Value = dr[i]["recommendations"].ToString();
                                            worksheet.Cells["P" + (startRows)].Value = dr[i]["responder_user_displayname"].ToString();
                                        }

                                        startRows++;
                                    }
                                    // วาดเส้นตาราง โดยใช้เซลล์ A3 ถึง P3 
                                    DrawTableBorders(worksheet, 14, 1, startRows - 1, (sce_show == true ? 18 : 16));

                                    worksheet.Cells["A" + (startRows)].Value = (dr[0]["descriptions_worksheet"] + "");

                                    if (show_cat == "0")
                                    {
                                        worksheet.DeleteColumn(5);
                                    }


                                    worksheet.Cells["A" + (14) + ":" + cell_h_end + (i + startRows)].Style.WrapText = true;

                                }
                            }

                            //new worksheet move after WorksheetTemplate 
                            excelPackage.Workbook.Worksheets.MoveBefore(worksheet_name, worksheet_name_target);
                        }
                        if (report_all == true)
                        {
                            if (dtSCE?.Rows.Count > 0)
                            {
                                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Safety Critical Equipment"];

                                int startRows = 3;
                                for (int s = 0; s < dtSCE?.Rows.Count; s++)
                                {
                                    worksheet.InsertRow(startRows, 1);

                                    worksheet.Cells["A" + (startRows)].Value = (s + 1);
                                    //if (s > 0)
                                    //{
                                    //    if (dtSCE.Rows[s - 1]["node"].ToString() != dtSCE.Rows[s]["node"].ToString())
                                    //    {
                                    //        worksheet.Cells["B" + (startRows)].Value = dtSCE.Rows[s]["node"].ToString();
                                    //    }
                                    //}
                                    //else
                                    //{
                                    worksheet.Cells["B" + (startRows)].Value = dtSCE.Rows[s]["node"].ToString();
                                    //}
                                    worksheet.Cells["C" + (startRows)].Value = dtSCE.Rows[s]["safety_critical_equipment_tag"].ToString();
                                    worksheet.Cells["D" + (startRows)].Value = dtSCE.Rows[s]["consequences"].ToString();
                                    worksheet.Cells["E" + (startRows)].Value = dtSCE.Rows[s]["ram_befor_risk"].ToString();
                                    startRows++;
                                }
                                // วาดเส้นตาราง โดยใช้เซลล์ A1 ถึง E3
                                DrawTableBorders(worksheet, 3, 1, startRows - 1, 5);

                            }
                        }
                        else
                        {
                            if (sce_show)
                            {
                                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Safety Critical Equipment"];

                                int startRows = 3;
                                for (int s = 0; s < dtSCE?.Rows.Count; s++)
                                {
                                    worksheet.InsertRow(startRows, 1);

                                    worksheet.Cells["A" + (startRows)].Value = (s + 1);
                                    //if (s > 0)
                                    //{
                                    //    if (dtSCE.Rows[s - 1]["node"].ToString() != dtSCE.Rows[s]["node"].ToString())
                                    //    {
                                    //        worksheet.Cells["B" + (startRows)].Value = dtSCE.Rows[s]["node"].ToString();
                                    //    }
                                    //}
                                    //else
                                    //{
                                    worksheet.Cells["B" + (startRows)].Value = dtSCE.Rows[s]["node"].ToString();
                                    //}
                                    worksheet.Cells["C" + (startRows)].Value = dtSCE.Rows[s]["safety_critical_equipment_tag"].ToString();
                                    worksheet.Cells["D" + (startRows)].Value = dtSCE.Rows[s]["consequences"].ToString();
                                    worksheet.Cells["E" + (startRows)].Value = dtSCE.Rows[s]["ram_befor_risk"].ToString();
                                    startRows++;
                                }
                                // วาดเส้นตาราง โดยใช้เซลล์ A1 ถึง E3
                                DrawTableBorders(worksheet, 3, 1, startRows - 1, 5);
                            }
                        }

                        if (report_all == true)
                        {
                            ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                            SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                            excelPackage.Save();
                        }
                        else
                        {
                            ExcelWorksheet SheetTemplateSCE = excelPackage.Workbook.Worksheets["WorksheetTemplateSCE"];
                            SheetTemplateSCE.Hidden = eWorkSheetHidden.Hidden;

                            ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                            SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                            excelPackage.Save();
                        }
                    }
                }

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }
        public string excel_hazop_recommendation(string seq, string file_fullpath_name, Boolean report_all, string seq_worksheet_def, string action_owner_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            //if (string.IsNullOrEmpty(seq_worksheet_def)) { return "Invalid Seq WorkSheet."; }

            string msg_error = "";

            try
            {
                #region Get Data
                List<SqlParameter> parameters = new List<SqlParameter>();
                sqlstr = @" select distinct
                        h.seq, h.pha_no, nl.id as id_node, g.pha_request_name
                        , nl.node, nl.node as node_check, nl.design_intent, nl.descriptions, nl.design_conditions, nl.node_boundary, nl.operating_conditions
                        , d.document_no, d.document_file_name
                        , mgw.guide_words as guideword, mgw.deviations as deviation, nw.causes, nw.consequences
                        , nw.category_type, nw.ram_befor_security, nw.ram_befor_likelihood, nw.ram_befor_risk
                        , nw.existing_safeguards, nw.recommendations, nw.recommendations_no, nw.responder_user_name, nw.responder_user_displayname
                        , nw.action_status
                        , nl.no as node_no, nw.no, nw.causes_no, nw.consequences_no
                        , nw.seq as seq_worksheet
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word
                        where h.seq = @seq and nw.responder_user_name is not null ";

                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-" });

                if (!string.IsNullOrEmpty(seq_worksheet_def))
                {
                    sqlstr += @" and nw.seq = @seq_worksheet_def ";
                    parameters.Add(new SqlParameter("@seq_worksheet_def", SqlDbType.VarChar, 50) { Value = seq_worksheet_def ?? "" });
                }
                if (!string.IsNullOrEmpty(action_owner_name))
                {
                    sqlstr += @" and lower(nw.responder_user_name) = lower(@action_owner_name) ";
                    parameters.Add(new SqlParameter("@action_owner_name", SqlDbType.VarChar, 400) { Value = action_owner_name ?? "" });
                }
                sqlstr += @" order by cast(nl.no as int),cast(nw.no as int), cast(nw.causes_no as int), cast(nw.consequences_no as int)";

                DataTable dtWorksheet = new DataTable();
                //dtWorksheet = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorksheet = new DataTable();
                        dtWorksheet = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorksheet.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                sqlstr = @" select distinct nl.no, nw.no, nw.seq, 0 as ref, nl.node, nl.node as node_check
                        , nw.ram_after_risk, nw.ram_after_risk_action, nw.recommendations, nw.recommendations_no, nw.action_status, nw.responder_user_name, nw.responder_user_displayname 
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word
                        where h.seq = @seq and nw.responder_user_name is not null ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

                if (!string.IsNullOrEmpty(action_owner_name))
                {
                    sqlstr += @" and lower(nw.responder_user_name) = lower(@action_owner_name) ";
                    parameters.Add(new SqlParameter("@action_owner_name", SqlDbType.VarChar, 50) { Value = action_owner_name ?? "" });
                }

                sqlstr += @" order by nl.no, nw.no ";

                DataTable dtTrack = new DataTable();
                //dtTrack = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtTrack = new DataTable();
                        dtTrack = _conn.ExecuteAdapter(command).Tables[0];
                        dtTrack.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (dtTrack!.Rows.Count > 0)
                {
                    for (int t = 0; t < dtTrack?.Rows.Count; t++)
                    {
                        dtTrack.Rows[t]["ref"] = (t + 1);
                        dtTrack.AcceptChanges();
                    }
                }
                #endregion Get Data

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "hazop";//sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);

                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                    {
                        if (!report_all)
                        {
                            var sheetsToDelete = excelPackage.Workbook.Worksheets
                               .Where(sheet => sheet.Name != "RecommTemplate" && sheet.Name != "TrackTemplate")
                               .ToList(); // ใช้ ToList เพื่อหลีกเลี่ยงปัญหา Collection Modified 
                            foreach (var sheet in sheetsToDelete)
                            {
                                excelPackage.Workbook.Worksheets.Delete(sheet);
                            }
                        }

                        DataTable dt = new DataTable();
                        dt = dtWorksheet.Copy(); dt.AcceptChanges();
                        for (int i = 0; i < dt?.Rows.Count; i++)
                        {
                            #region Sheet
                            if (true)
                            {
                                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["RecommTemplate"];
                                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("RecommTemplate" + i, sourceWorksheet);

                                string ref_no = (i + 1).ToString();
                                worksheet.Name = "Response Sheet(Ref." + ref_no + ")";

                                string responder_user_name = (dt.Rows[i]["responder_user_name"] + "");
                                string responder_user_displayname = (dt.Rows[i]["responder_user_displayname"] + "");
                                string pha_request_name = (dt.Rows[i]["pha_request_name"] + "");
                                string pha_no = (dt.Rows[i]["pha_no"] + "");
                                string seq_worksheet = (dt.Rows[i]["seq_worksheet"] + "");


                                int startRows = 2;
                                if (true)
                                {
                                    string node = "";
                                    string drawing_doc = "";
                                    string deviation = "";
                                    string causes = "";
                                    string consequences = "";
                                    string existing_safeguards = "";
                                    string recommendations = "";
                                    string recommendations_no = "";
                                    int action_no = 0;

                                    #region loop drawing_doc 
                                    drawing_doc = (dt.Rows[i]["document_no"] + "");
                                    if ((dt.Rows[i]["document_file_name"] + "") != "")
                                    {
                                        drawing_doc += " (" + dt.Rows[i]["document_file_name"] + ")";
                                    }
                                    #endregion loop drawing_doc 

                                    #region loop workksheet 
                                    //DataRow[] drWorksheet = dt.Select("seq_worksheet = '" + seq_worksheet + "'");
                                    var filterParameters = new Dictionary<string, object>();
                                    filterParameters.Add("seq_worksheet", seq_worksheet);
                                    var (drWorksheet, iMerge) = FilterDataTable(dtWorksheet, filterParameters);
                                    if (drWorksheet != null)
                                    {
                                        if (drWorksheet?.Length > 0)
                                        {
                                            for (int n = 0; n < drWorksheet.Length; n++)
                                            {
                                                if ((drWorksheet[n]["deviation"] + "") != "")
                                                {
                                                    if (deviation != "") { deviation += ","; }
                                                    deviation += (drWorksheet[n]["guideword"] + "") + "/" + (drWorksheet[n]["deviation"] + "");
                                                }
                                                if ((drWorksheet[n]["causes"] + "") != "")
                                                {
                                                    if (causes != "") { causes += ","; }
                                                    causes += (drWorksheet[n]["causes"] + "");
                                                }
                                                if ((drWorksheet[n]["consequences"] + "") != "")
                                                {
                                                    if (consequences != "") { consequences += ","; }
                                                    consequences += (drWorksheet[n]["consequences"] + "");
                                                }

                                                if ((drWorksheet[n]["existing_safeguards"] + "") != "")
                                                {
                                                    if (existing_safeguards.IndexOf((drWorksheet[n]["existing_safeguards"] + "")) > -1) { }
                                                    else
                                                    {
                                                        if (existing_safeguards != "") { existing_safeguards += ","; }
                                                        existing_safeguards += (drWorksheet[n]["existing_safeguards"] + "");
                                                    }
                                                }

                                                if ((drWorksheet[n]["recommendations"] + "") != "")
                                                {
                                                    if (recommendations != "") { recommendations += ","; }
                                                    recommendations += (drWorksheet[n]["recommendations"] + "");
                                                    action_no += 1;

                                                    if (recommendations_no != "") { recommendations_no += ","; }
                                                    recommendations_no += (drWorksheet[n]["recommendations_no"] + "");
                                                }

                                            }
                                        }
                                    }
                                    #endregion loop workksheet

                                    worksheet.Cells["A" + (startRows)].Value = "Project Title:" + pha_request_name;
                                    startRows += 1;
                                    worksheet.Cells["A" + (startRows)].Value = "Project No:" + pha_no;
                                    startRows += 1;
                                    worksheet.Cells["A" + (startRows)].Value = "Node:" + node;
                                    startRows += 1;

                                    worksheet.Cells["B" + (startRows)].Value = responder_user_displayname;
                                    worksheet.Cells["E" + (startRows)].Value = responder_user_displayname;
                                    startRows += 1;

                                    worksheet.Cells["B" + (startRows)].Value = action_no;
                                    startRows += 1;

                                    worksheet.Cells["B" + (startRows)].Value = drawing_doc;
                                    startRows += 1;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = deviation;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = causes;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = consequences;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = existing_safeguards;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = recommendations;
                                    startRows += 1;
                                }

                            }
                            #endregion Sheet

                        }

                        #region TrackTemplate
                        if (dtTrack?.Rows.Count > 0)
                        {
                            //ข้อมูลทั้งหมด
                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TrackTemplate"];
                            worksheet.Name = "Status Tracking Table";

                            int i = 0;
                            int startRows = 3;

                            dt = new DataTable(); dt = dtTrack.Copy(); dt.AcceptChanges();
                            if (dt?.Rows.Count > 0)
                            {
                                for (i = 0; i < dt?.Rows.Count; i++)
                                {
                                    worksheet.InsertRow(startRows, 1);
                                    worksheet.Cells["A" + (startRows)].Value = dt.Rows[i]["ref"].ToString();
                                    worksheet.Cells["B" + (startRows)].Value = dt.Rows[i]["node"].ToString();
                                    worksheet.Cells["C" + (startRows)].Value = dt.Rows[i]["ram_after_risk"].ToString();
                                    worksheet.Cells["D" + (startRows)].Value = dt.Rows[i]["recommendations"].ToString();
                                    worksheet.Cells["E" + (startRows)].Value = dt.Rows[i]["action_status"].ToString();
                                    worksheet.Cells["F" + (startRows)].Value = dt.Rows[i]["responder_user_displayname"].ToString();
                                    startRows++;
                                }

                                // วาดเส้นตาราง โดยใช้เซลล์ A1 ถึง C3
                                DrawTableBorders(worksheet, 3, 1, startRows - 1, 6);
                            }
                        }
                        #endregion Response Sheet

                        ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["RecommTemplate"];
                        SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                        excelPackage.Save();

                    }
                }

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }
        public string excel_hazop_guidewords(string seq, string file_fullpath_name, Boolean report_all)
        {

            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            try
            {
                List<SqlParameter> parameters = new List<SqlParameter>();
                sqlstr = @" select distinct parameter
                        from epha_m_guide_words where active_type = 1
                        order by parameter ";

                DataTable dtParam = new DataTable();
                //dtParam = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtParam = new DataTable();
                        dtParam = _conn.ExecuteAdapter(command).Tables[0];
                        dtParam.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable



                sqlstr = @" select '' as usef_selected, def_selected, parameter, deviations, guide_words, process_deviation, area_application
                        from epha_m_guide_words where active_type = 1
                        order by parameter, deviations, guide_words, process_deviation, area_application ";

                parameters = new List<SqlParameter>();
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

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "hazop";//sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                    {
                        if (!report_all)
                        {
                            var sheetsToDelete = excelPackage.Workbook.Worksheets
                              .Where(sheet => sheet.Name != "GuidewordsTemplate")
                              .ToList(); // ใช้ ToList เพื่อหลีกเลี่ยงปัญหา Collection Modified

                            foreach (var sheet in sheetsToDelete)
                            {
                                excelPackage.Workbook.Worksheets.Delete(sheet);
                            }
                        }


                        ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["GuidewordsTemplate"];
                        ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("Guidewords", sourceWorksheet);
                        worksheet.Name = "Guidewords";

                        int startRows = 3;
                        int i = 0;
                        for (int m = 0; m < dtParam?.Rows.Count; m++)
                        {
                            string parameter = (dtParam.Rows[m]["parameter"] + "");
                            worksheet.InsertRow(startRows, 1);
                            var startCell = worksheet.Cells["A" + startRows];
                            var endCell = worksheet.Cells["D" + startRows];
                            var mergeRange = worksheet.Cells[startCell.Address + ":" + endCell.Address];
                            // Merge the cells
                            mergeRange.Merge = true;
                            // Optionally set text alignment in the merged cell
                            mergeRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            mergeRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            worksheet.Cells["A" + (startRows)].Value = parameter;
                            startRows++;

                            //DataRow[] dr = dt.Select("parameter = '" + parameter + "'");
                            var filterParameters = new Dictionary<string, object>();
                            filterParameters.Add("parameter", parameter);
                            var (dr, iMerge) = FilterDataTable(dt, filterParameters);
                            if (dr != null)
                            {
                                if (dr?.Length > 0)
                                {
                                    for (i = 0; i < dr.Length; i++)
                                    {
                                        worksheet.InsertRow(startRows, 1);
                                        worksheet.Cells["A" + (startRows)].Value = dr[i]["deviations"].ToString();
                                        worksheet.Cells["B" + (startRows)].Value = dr[i]["guide_words"].ToString();
                                        worksheet.Cells["C" + (startRows)].Value = dr[i]["process_deviation"].ToString();
                                        worksheet.Cells["D" + (startRows)].Value = dr[i]["area_application"].ToString();
                                        startRows++;
                                    }
                                }
                            }
                        }

                        // วาดเส้นตาราง โดยใช้เซลล์ A1 ถึง D3
                        DrawTableBorders(worksheet, 3, 1, startRows - 1, 4);

                        excelPackage.Save();
                    }
                }

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }


        public string export_report_recommendation(ReportByWorksheetModel param, Boolean report_by_action_owner, Boolean all_items, Boolean res_fullpath = false)
        {
            string msg_error = "";
            DataSet _dsData = new DataSet();
            DataTable dtdef = (DataTable)ClassFile.DatatableFile();
            string _file_name = "";
            string _file_download_name = "";
            string _file_fullpath_name = "";

            try
            {
                if (param == null) { msg_error = $"The specified file does not param."; }
                else
                {

                    string seq = param?.seq ?? "";
                    string export_type = param?.export_type ?? "";
                    string sub_software = param.sub_software ?? "";
                    string seq_worksheet = (""); ;
                    string user_name = (param?.user_name ?? "");

                    if (!all_items) { seq_worksheet = (param?.seq_worksheet ?? ""); }

                    // ตรวจสอบ sub_software ว่ามีอยู่ใน whitelist
                    var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };
                    if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                    {
                        return cls_json.SetJSONresult(ClassFile.refMsg("Error", "Invalid sub_software."));
                    }

                    if (!string.IsNullOrEmpty(seq) && !string.IsNullOrEmpty(sub_software))
                    {
                        string folder = sub_software ?? "";
                        string file_part = (export_type == "template" ? "Template" : "Report");
                        if (export_type == "template") { export_type = "excel"; }

                        if (!res_fullpath) { _file_name = "Recommendation"; }

                        //copy template to new file report 
                        ClassFile.copy_file_excel_template(ref _file_name, ref _file_download_name, ref _file_fullpath_name, folder, file_part, "");
                        if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(_file_name) || string.IsNullOrEmpty(_file_download_name) || string.IsNullOrEmpty(_file_fullpath_name))
                        { msg_error = "Invalid folder."; }
                        else
                        {
                            if (!string.IsNullOrEmpty(_file_fullpath_name))
                            {
                                string file_fullpath_def = _file_fullpath_name;
                                //string folder = "hazop";//sub_software;


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
                                    _file_fullpath_name = fullPath;
                                }
                                else { _file_fullpath_name = ""; }
                            }
                            if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(_file_fullpath_name))
                            {
                                switch (sub_software?.ToLower())
                                {
                                    case "hazop":

                                        msg_error = excel_hazop_recommendation(seq, _file_fullpath_name, false, seq_worksheet, (report_by_action_owner == true ? user_name : ""));
                                        if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }

                                        break;
                                    case "whatif":

                                        msg_error = excel_whatif_recommendation(seq, _file_fullpath_name, false, seq_worksheet, (report_by_action_owner == true ? user_name : ""));
                                        if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }
                                        break;

                                    case "hra":

                                        msg_error = excel_hra_worksheet(seq, _file_fullpath_name, false, true);
                                        if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }
                                        break;
                                }

                                // Save the workbook as PDF

                                // ตรวจสอบการมีอยู่ของ LibreOffice ใน tools directory
                                string libreOfficePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "LibreOffice", "program", "soffice.exe");
                                if (!File.Exists(libreOfficePath))
                                {
                                    return "LibreOffice executable not found in project directory.";
                                }

                                if (export_type == "pdf")
                                {
                                    try
                                    {
                                        FileInfo template = new FileInfo(_file_fullpath_name);
                                        if (!template.Exists || template.IsReadOnly)
                                        {
                                            if (template.IsReadOnly)
                                            {
                                                template.IsReadOnly = false;
                                            }
                                            else
                                            {
                                                msg_error = "File permissions are not correctly set.";
                                                goto Next_Line;
                                            }
                                        }

                                        // ใช้ LibreOffice ในการแปลงไฟล์
                                        string _file_fullpath_name_pdf = _file_fullpath_name.Replace(".xlsx", ".pdf");

                                        var process = new System.Diagnostics.Process();
                                        process.StartInfo.FileName = libreOfficePath;
                                        //process.StartInfo.Arguments = $"--headless --convert-to pdf --outdir \"{Path.GetDirectoryName(_file_fullpath_name)}\" \"{_file_fullpath_name_pdf}\"";
                                        process.StartInfo.Arguments = $"--headless --convert-to pdf --outdir \"{Path.GetDirectoryName(_file_fullpath_name)}\" \"{_file_fullpath_name}\"";
                                        process.StartInfo.CreateNoWindow = true;
                                        process.StartInfo.UseShellExecute = false;
                                        process.StartInfo.RedirectStandardOutput = true;

                                        process.Start();

                                        bool exited = process.WaitForExit(30000); // รอให้กระบวนการทำงานเสร็จภายใน 30 วินาที (30000 มิลลิวินาที)
                                        if (!exited)
                                        {
                                            // ถ้ากระบวนการไม่เสร็จภายในเวลาที่กำหนด ให้บังคับปิด
                                            process.Kill();
                                            throw new Exception("Process timed out and was killed.");
                                        }

                                        if (process.ExitCode == 0)
                                        {
                                            // ตรวจสอบการสร้างไฟล์ PDF
                                            msg_error = ClassFile.check_file_other(_file_fullpath_name_pdf, ref _file_fullpath_name_pdf, ref _file_download_name, folder);

                                            if (string.IsNullOrEmpty(msg_error))
                                            {
                                                // เพิ่มไฟล์เข้าไปใน appendix
                                                msg_error = add_drawing_to_appendix(seq, _file_fullpath_name_pdf, folder);
                                            }
                                            else
                                            {
                                                msg_error = $"Failed to create PDF file: {_file_fullpath_name_pdf}";
                                            }
                                        }
                                        else
                                        {
                                            msg_error = "Failed to convert file using LibreOffice.";
                                        }

                                        // เปลี่ยนเส้นทาง fullpath filename
                                        _file_fullpath_name = _file_fullpath_name_pdf;
                                        if (!string.IsNullOrEmpty(_file_fullpath_name))
                                        {
                                            msg_error = ClassFile.check_format_file_name(_file_fullpath_name);
                                            if (!string.IsNullOrEmpty(msg_error))
                                            {
                                                msg_error = $"Failed to change path Excel to PDF file ";
                                            }
                                        }
                                        else
                                        {
                                            _file_name = (_file_name?.ToLower() ?? "").Replace(".xlsx", ".pdf");
                                            if (!string.IsNullOrEmpty(_file_name))
                                            {
                                                msg_error = ClassFile.check_format_file_name(_file_name);
                                                if (!string.IsNullOrEmpty(msg_error))
                                                {
                                                    msg_error = $"Failed to replace PDF file: {_file_name}";
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex_pdf)
                                    {
                                        msg_error = ex_pdf.Message.ToString();
                                    }
                                }
                            }

                        }
                    }



                Next_Line:;
                }
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }



            if (res_fullpath)
            {
                return _file_fullpath_name;
            }
            else
            {
                if (dtdef != null)
                {
                    ClassFile.AddRowToDataTable(ref dtdef, _file_name, _file_download_name, msg_error);
                    DataSet dsData = new DataSet();
                    _dsData.Tables.Add(dtdef.Copy());
                }
                return JsonConvert.SerializeObject(_dsData, Formatting.Indented);
            }

        }
        public string export_hazop_guidewords(ReportModel param)
        {
            string msg_error = "";
            DataSet _dsData = new DataSet();
            DataTable dtdef = (DataTable)ClassFile.DatatableFile();
            string _file_name = "";
            string _file_download_name = "";
            string _file_fullpath_name = "";
            try
            {
                if (param == null) { msg_error = $"The specified file does not param."; }
                else
                {
                    string seq = param?.seq ?? "";
                    string export_type = param?.export_type ?? "";
                    string sub_software = param.sub_software ?? "";
                    string user_name = (param?.user_name ?? "");

                    // ตรวจสอบ sub_software ว่ามีอยู่ใน whitelist
                    var allowedSubSoftware = new HashSet<string> { "hazop", "whatif", "jsea", "hra", "bowtie" };
                    if (!allowedSubSoftware.Contains(sub_software.ToLower()))
                    {
                        msg_error = "Invalid sub_software.";
                        return JsonConvert.SerializeObject(_dsData, Formatting.Indented);
                    }

                    if (!string.IsNullOrEmpty(seq) && !string.IsNullOrEmpty(sub_software))
                    {
                        string folder = sub_software ?? "";
                        string file_part = (export_type == "template" ? "Template" : "Report");
                        if (export_type == "template") { export_type = "excel"; }

                        //copy template to new file report 
                        ClassFile.copy_file_excel_template(ref _file_name, ref _file_download_name, ref _file_fullpath_name, folder, file_part, "");
                        if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(_file_name) || string.IsNullOrEmpty(_file_download_name) || string.IsNullOrEmpty(_file_fullpath_name))
                        { msg_error = "Invalid folder."; }
                        else
                        {
                            if (!string.IsNullOrEmpty(_file_fullpath_name))
                            {
                                string file_fullpath_def = _file_fullpath_name;
                                //string folder = "hazop";//sub_software;


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
                                    _file_fullpath_name = fullPath;
                                }
                                else { _file_fullpath_name = ""; }
                            }

                            if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(_file_fullpath_name))
                            {
                                msg_error = excel_hazop_guidewords(seq, _file_fullpath_name, false);
                            }
                        }
                    }


                }
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }


            if (dtdef != null)
            {
                ClassFile.AddRowToDataTable(ref dtdef, _file_name, _file_download_name, msg_error);
                DataSet dsData = new DataSet();
                _dsData.Tables.Add(dtdef.Copy());
            }
            return JsonConvert.SerializeObject(_dsData, Formatting.Indented);
        }

        #endregion export excel hazop

        #region export excel jsea
        public string excel_jsea_general(string seq, string file_fullpath_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";
            try
            {
                #region get data
                List<SqlParameter> parameters = new List<SqlParameter>();
                sqlstr = @" select g.work_scope, h.pha_no, h.pha_version_text as pha_version, g.descriptions, format(g.target_start_date, 'dd MMM yyyy ') as target_start_date 
                         ,  case when ums.user_name is null  then h.request_user_displayname else case when ums.departments is null  then  ums.user_displayname else  ums.user_displayname + ' (' + ums.departments +')' end end request_user_displayname
                         , g.mandatory_note 
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         left join VW_EPHA_PERSON_DETAILS ums on lower(h.request_user_name) = lower(ums.USER_NAME) 
                         where h.seq = @seq ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtWorkScope = new DataTable();
                //dtWorkScope = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorkScope = new DataTable();
                        dtWorkScope = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorkScope.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                sqlstr = @" select distinct d.no, d.document_name, d.document_no, d.document_file_name, d.descriptions 
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                        inner join EPHA_T_DRAWING d on h.id = d.id_pha    
                        where h.seq = @seq and d.document_name is not null order by convert(int,d.no) ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
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
                        dtDrawing = new DataTable();
                        dtDrawing = _conn.ExecuteAdapter(command).Tables[0];
                        dtDrawing.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                #endregion get data 

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "jsea";//sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                    {
                        ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];  // Replace "SourceSheet" with the actual source sheet name
                        ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                        //Whatif Cover Page
                        worksheet = excelPackage.Workbook.Worksheets["JSEA Cover Page"];

                        ClassHazop clshazop = new ClassHazop();
                        worksheet.Cells["A14"].Value = (dtWorkScope.Rows[0]["pha_version"] + "");
                        worksheet.Cells["B14"].Value = (dtWorkScope.Rows[0]["target_start_date"] + "");
                        worksheet.Cells["C14"].Value = (dtWorkScope.Rows[0]["request_user_displayname"] + "");
                        worksheet.Cells["D14"].Value = (dtWorkScope.Rows[0]["descriptions"] + "");



                        //Study Objective and Work Scope
                        worksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];
                        worksheet.Cells["A2"].Value = (dtWorkScope.Rows[0]["work_scope"] + "");

                        //Description/Remarks 
                        if ((dtWorkScope.Rows[0]["descriptions"] + "") == "")
                        {
                            //remove rows A4, A5
                            worksheet.DeleteRow(5); worksheet.DeleteRow(4);
                        }
                        else
                        {
                            worksheet.Cells["A5"].Value = (dtWorkScope.Rows[0]["descriptions"] + "");
                        }

                        //Drawing & Reference
                        #region Drawing & Reference
                        if (true)
                        {
                            worksheet = excelPackage.Workbook.Worksheets["Document & Reference"];

                            int startRows = 3;
                            int icol_end = 6;
                            int ino = 1;
                            for (int i = 0; i < dtDrawing?.Rows.Count; i++)
                            {
                                //No.	Document Name	Document File	Comment
                                worksheet.InsertRow(startRows, 1);
                                worksheet.Cells["A" + (startRows)].Value = (i + 1); ;
                                worksheet.Cells["B" + (startRows)].Value = (dtDrawing.Rows[i]["document_name"] + "");
                                worksheet.Cells["C" + (startRows)].Value = (dtDrawing.Rows[i]["document_file_name"] + "");
                                worksheet.Cells["D" + (startRows)].Value = (dtDrawing.Rows[i]["descriptions"] + "");
                                startRows++;
                            }
                            // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                            DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                        }
                        #endregion Drawing & Reference

                        //Study Objective and Work Scope
                        #region Study Objective and Work Scope
                        if (true)
                        {
                            worksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];
                            worksheet.Cells["A2"].Value = (dtWorkScope.Rows[0]["descriptions"] + "");
                        }
                        #endregion Study Objective and Work Scope

                        excelPackage.Save();
                    }
                }

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }
        public string excel_jsea_atendeesheet(string seq, string file_fullpath_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            string msg_error = "";

            #region Get Data
            List<SqlParameter> parameters = new List<SqlParameter>();
            sqlstr = @" select distinct a.id as id_pha, a.pha_status, ta2.no
                     , isnull(emp.user_name,'') as user_name, emp.user_displayname, emp.user_email
                     , case when ta2.action_review = 2 then (case when ta2.action_status = 'approve' then 'Approve' else
						(case when ta2.action_status = 'reject' and ta2.comment is not null  then 'Send Back' else 'Send Back with comment' end) 
					  end) else '' end action_status
					  
                     from epha_t_header a  
                     inner join EPHA_T_GENERAL g on a.id = g.id_pha   
                     inner join EPHA_T_SESSION s on a.id = s.id_pha 
                     inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s2 on a.id = s2.id_pha and s.id = s2.id_session 
                     inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha ) t2 on a.id = t2.id_pha and s2.id_session = t2.id_session 
                     inner join EPHA_T_APPROVER ta2 on a.id = ta2.id_pha and t2.id_pha = ta2.id_pha and t2.id_session = ta2.id_session 
                     left join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name) 
					 inner join  (select max(seq)as seq, pha_no from epha_t_header group by pha_no) hm on a.seq = hm.seq and a.pha_no = hm.pha_no
                     where a.request_approver = 1 and a.seq = @seq order by ta2.no";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            DataTable dtTA2 = new DataTable();
            //dtTA2 = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                    dtTA2 = new DataTable();
                    dtTA2 = _conn.ExecuteAdapter(command).Tables[0];
                    dtTA2.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable


            sqlstr = "select a.* from VW_EPHA_DATA_TEAMMEMBER_ALL a where seq is not null  where a.seq = @seq  ";
            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            DataTable dtAll = new DataTable();
            //dtAll = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                    dtAll = new DataTable();
                    dtAll = _conn.ExecuteAdapter(command).Tables[0];
                    dtAll.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            sqlstr = @" select distinct 0 as no, t.user_name, t.user_displayname, '' as company_text from VW_EPHA_DATA_TEAMMEMBER_ALL t where t.seq = @seq and  t.user_name <> '' order by t.user_name";
            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            DataTable dtMember = new DataTable();
            //dtMember = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                    dtMember = new DataTable();
                    dtMember = _conn.ExecuteAdapter(command).Tables[0];
                    dtMember.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            sqlstr = @" select distinct t.seq_session, t.session_no, t.meeting_date from VW_EPHA_DATA_TEAMMEMBER_ALL t where t.seq = @seq order by t.session_no ";
            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            DataTable dtSession = new DataTable();
            //dtSession = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                    dtSession = new DataTable();
                    dtSession = _conn.ExecuteAdapter(command).Tables[0];
                    dtSession.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            #endregion Get Data

            if (!string.IsNullOrEmpty(file_fullpath_name))
            {
                string file_fullpath_def = file_fullpath_name;
                string folder = "jsea";//sub_software;


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

            if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
            {
                FileInfo template_excel = new FileInfo(file_fullpath_name);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                {
                    ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];  // Replace "SourceSheet" with the actual source sheet name
                    sourceWorksheet.Name = "JSEA Attendee Sheet";
                    ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                    int i = 0;
                    int startRows = 4;
                    int icol_start = 4;
                    int icol_end = 0;// icol_start + (dtSession.Rows.Count > 6 ? dtSession.Rows.Count : 6);
                    if (dtSession?.Rows.Count > 0)
                    {
                        int icount_session = dtSession?.Rows.Count ?? 0;

                        icol_end = icol_start + (icount_session > 6 ? icount_session : 6);
                    }

                    for (int imember = 0; imember < dtMember?.Rows.Count; imember++)
                    {
                        worksheet.InsertRow(startRows, 1);
                        string user_name = (dtMember.Rows[imember]["user_name"] + "");
                        //No.
                        worksheet.Cells["A" + (i + startRows)].Value = (imember + 1);
                        //Name
                        worksheet.Cells["B" + (i + startRows)].Value = (dtMember.Rows[imember]["user_displayname"] + "");
                        //Company
                        worksheet.Cells["C" + (i + startRows)].Value = (dtMember.Rows[imember]["company_text"] + "");

                        int irow_session = 0;
                        if (imember == 0)
                        {
                            if (dtSession?.Rows.Count < 6)
                            {
                                //worksheet.Cells[2, icol_start, 2, icol_end].Merge = true; 
                                for (int c = icol_end; c < 30; c++)
                                {
                                    worksheet.DeleteColumn(icol_end);

                                }
                            }

                            irow_session = 0;
                            for (int c = icol_start; c < icol_end; c++)
                            {
                                try
                                {
                                    //header 
                                    if ((dtSession.Rows[irow_session]["meeting_date"] + "") == "")
                                    {
                                        worksheet.Cells[3, c].Value = "";
                                    }
                                    else
                                    {
                                        worksheet.Cells[3, c].Value = (dtSession.Rows[irow_session]["meeting_date"] + "");
                                    }
                                }
                                catch { worksheet.Cells[3, c].Value = ""; }
                                irow_session += 1;
                            }
                        }

                        irow_session = 0;
                        for (int c = icol_start; c < icol_end; c++)
                        {
                            try
                            {
                                string session_no = "";
                                try { session_no = (dtSession.Rows[irow_session]["session_no"] + ""); } catch { }
                                worksheet.Cells[startRows, c].Value = "";

                                //DataRow[] dr = dtAll.Select("user_name = '" + user_name + "' and session_no = '" + session_no + "'");
                                var filterParameters = new Dictionary<string, object>();
                                filterParameters.Add("user_name", user_name);
                                filterParameters.Add("session_no", session_no);
                                var (dr, iMerge) = FilterDataTable(dtAll, filterParameters);
                                if (dr != null)
                                {
                                    if (dr?.Length > 0) { worksheet.Cells[startRows, c].Value = "X"; }
                                }
                            }
                            catch { }
                            irow_session++;

                        }

                        startRows++;
                    }
                    // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                    DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                    //TA2 
                    startRows += 13;
                    int startRows_ta2 = startRows;

                    for (int ita2 = 0; ita2 < dtTA2?.Rows.Count; ita2++)
                    {
                        worksheet.InsertRow(startRows, 1);
                        //No.
                        worksheet.Cells["A" + (i + startRows)].Value = (ita2 + 1);
                        //Name
                        worksheet.Cells["B" + (i + startRows)].Value = (dtTA2.Rows[ita2]["user_displayname"] + "");
                        //status
                        worksheet.Cells["C" + (i + startRows)].Value = (dtTA2.Rows[ita2]["action_status"] + "");

                        startRows++;
                    }
                    // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                    DrawTableBorders(worksheet, startRows_ta2, 1, startRows - 1, 3);

                    excelPackage.Save();
                }
            }
        Next_Line:;

            return msg_error;
        }
        public string excel_jsea_worksheet(string seq, string file_fullpath_name, Boolean report_all)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            List<SqlParameter> parameters = new List<SqlParameter>();
            #region get data 
            Boolean bApproverSafetyReview = false;
            sqlstr = @"  select g.pha_request_name, format(g.target_start_date,'dd MMM yyyy') as target_start_date, g.mandatory_note, g.descriptions
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                         where h.seq = @seq ";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            DataTable dtHead = new DataTable();
            //dtHead = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                    dtHead = new DataTable();
                    dtHead = _conn.ExecuteAdapter(command).Tables[0];
                    dtHead.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            sqlstr = @"   select tw.row_type, tw.no, tw.workstep_no, tw.workstep, tw.taskdesc_no, tw.taskdesc, tw.potentailhazard_no, tw.potentailhazard, tw.possiblecase_no, tw.possiblecase  
                         , tw.category_no, tw.category_type, tw.ram_befor_security, tw.ram_befor_likelihood, tw.ram_befor_risk, tw.recommendations, tw.responder_action_by
                         , tw.ram_after_security, tw.ram_after_likelihood, tw.ram_after_risk
                         , g.id_ram
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                         inner join EPHA_T_TASKS_WORKSHEET tw on h.id  = tw.id_pha 
                         where h.seq = @seq order by tw.no  ";


            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            DataTable dtWorksheet = new DataTable();
            //dtWorksheet = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                    dtWorksheet = new DataTable();
                    dtWorksheet = _conn.ExecuteAdapter(command).Tables[0];
                    dtWorksheet.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            sqlstr = @"select t.* from VW_EPHA_DATA_RELATEDPEOPLE_ALL t where t.seq = @seq order by t.user_type, t.no ";
            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
            DataTable dtRelatedPeople = new DataTable();
            //dtRelatedPeople = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                    dtRelatedPeople = new DataTable();
                    dtRelatedPeople = _conn.ExecuteAdapter(command).Tables[0];
                    dtRelatedPeople.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            if (dtRelatedPeople != null)
            {
                if (dtRelatedPeople?.Rows.Count > 0)
                {
                    var filterParameters = new Dictionary<string, object>();
                    filterParameters.Add("user_type", "reviewer");
                    var (dr, iMerge) = FilterDataTable(dtRelatedPeople, filterParameters);
                    if (dr != null)
                    {
                        if (dr?.Length > 0)
                        {
                            bApproverSafetyReview = true;
                        }
                    }
                }
            }



            sqlstr = @" select s.id_pha, s.seq as seq_session, s.no as session_no
                         , convert(varchar,s.meeting_date,106) as meeting_date
                         , mt.no as member_no, isnull(mt.user_name,'') as user_name
                         ,'member' as user_type, mt.no, emp.user_displayname, emp.user_title
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_SESSION s on h.id = s.id_pha 
                         left join EPHA_T_MEMBER_TEAM mt on h.id = mt. id_pha and mt.id_session = s.id
                         left join VW_EPHA_PERSON_DETAILS emp on lower(emp.user_name) = lower(mt.user_name)
                         where h.seq = @seq and lower(mt.user_name) is not null order by mt.no";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

            DataTable dtMemberTeam = new DataTable();
            //dtMemberTeam = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                    dtMemberTeam = new DataTable();
                    dtMemberTeam = _conn.ExecuteAdapter(command).Tables[0];
                    dtMemberTeam.AcceptChanges();
                }
                catch { }
                finally { _conn.CloseConnection(); }
            }
            catch { }
            #endregion Execute to Datable

            #endregion get data

            ////fix RAM  => 5 
            //string ram_type = "5";

            if (!string.IsNullOrEmpty(file_fullpath_name))
            {
                string file_fullpath_def = file_fullpath_name;
                string folder = "jsea";//sub_software;


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

            if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
            {
                FileInfo template_excel = new FileInfo(file_fullpath_name);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                {
                    ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                    ExcelWorksheet worksheet = sourceWorksheet;

                    //Worksheet
                    #region Worksheet
                    if (true)
                    {
                        if (dtHead?.Rows.Count > 0)
                        {
                            //ram_type = (dtWorksheet.Rows[0]["id_ram"] + "");

                            //header
                            worksheet.Cells["C4"].Value = (dtHead.Rows[0]["pha_request_name"] + "");
                            //if (ram_type == "5")
                            //{
                            //    worksheet.Cells["K4"].Value = "" + (dtHead.Rows[0]["target_start_date"] + "");
                            //}
                            //else
                            //{
                            worksheet.Cells["J4"].Value = "" + (dtHead.Rows[0]["target_start_date"] + "");
                            //}
                            worksheet.Cells["D5"].Value = "" + (dtHead.Rows[0]["mandatory_note"] + "");
                            worksheet.Cells["D5"].Style.WrapText = true;
                        }

                        int startRows = 12;
                        int startRows_Def = 12;
                        int icol_start = 1;
                        int icol_end = 6;
                        if (true)
                        {
                            int iDefRow = 20;
                            for (int i = 0; i < dtWorksheet?.Rows.Count; i++)
                            {
                                if (i >= iDefRow)
                                {
                                    worksheet.InsertRow(startRows + i, 1, startRows_Def);
                                }
                            }

                            for (int i = 0; i < dtWorksheet?.Rows.Count; i++)
                            {
                                //if (i >= iDefRow)
                                //{
                                //    worksheet.InsertRow(startRows, 1, startRows_Def);
                                //} 

                                //Merge
                                if (true)
                                {
                                    string row_type = (dtWorksheet.Rows[i]["row_type"] + "");
                                    icol_start = 1;
                                    icol_start += 1;
                                    if (row_type == "workstep")
                                    {
                                        //int iMerge = (dtWorksheet.Select("workstep_no=" + dtWorksheet.Rows[i]["workstep_no"])).Length;
                                        // กำหนดพารามิเตอร์ที่ต้องการกรอง
                                        var filterParameters = new Dictionary<string, object>();
                                        filterParameters.Add("workstep_no", dtWorksheet.Rows[i]["workstep_no"]);
                                        // เรียกใช้ฟังก์ชัน FilterDataTable 
                                        var (filteredRows, iMerge) = FilterDataTable(dtWorksheet, filterParameters);
                                        if (iMerge > 1)
                                        {
                                            try
                                            {
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Merge = true;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            }
                                            catch (Exception) { }
                                        }
                                    }
                                    icol_start += 1;
                                    if (row_type == "workstep" || row_type == "taskdesc")
                                    {
                                        //int iMerge = (dtWorksheet.Select("workstep_no=" + dtWorksheet.Rows[i]["workstep_no"] + " and taskdesc_no=" + dtWorksheet.Rows[i]["taskdesc_no"])).Length;
                                        // กำหนดพารามิเตอร์ที่ต้องการกรอง
                                        var filterParameters = new Dictionary<string, object>();
                                        filterParameters.Add("workstep_no", dtWorksheet.Rows[i]["workstep_no"]);
                                        filterParameters.Add("taskdesc_no", dtWorksheet.Rows[i]["taskdesc_no"]);
                                        // เรียกใช้ฟังก์ชัน FilterDataTable 
                                        var (filteredRows, iMerge) = FilterDataTable(dtWorksheet, filterParameters);
                                        if (iMerge > 1)
                                        {
                                            try
                                            {
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Value = "";
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Merge = true;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            }
                                            catch (Exception) { }
                                        }
                                    }
                                    icol_start += 1;
                                    if (row_type == "workstep" || row_type == "taskdesc" || row_type == "potentailhazard")
                                    {
                                        //int iMerge = (dtWorksheet.Select("workstep_no=" + dtWorksheet.Rows[i]["workstep_no"] + " and taskdesc_no=" + dtWorksheet.Rows[i]["taskdesc_no"] + "  and potentailhazard_no=" + dtWorksheet.Rows[i]["potentailhazard_no"] + " )).Length;
                                        // กำหนดพารามิเตอร์ที่ต้องการกรอง
                                        var filterParameters = new Dictionary<string, object>();
                                        filterParameters.Add("workstep_no", dtWorksheet.Rows[i]["workstep_no"]);
                                        filterParameters.Add("taskdesc_no", dtWorksheet.Rows[i]["taskdesc_no"]);
                                        filterParameters.Add("potentailhazard_no", dtWorksheet.Rows[i]["potentailhazard_no"]);
                                        // เรียกใช้ฟังก์ชัน FilterDataTable 
                                        var (filteredRows, iMerge) = FilterDataTable(dtWorksheet, filterParameters);
                                        if (iMerge > 1)
                                        {
                                            try
                                            {
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Merge = true;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            }
                                            catch (Exception) { }
                                        }
                                    }
                                    icol_start += 1;
                                    if (row_type == "workstep" || row_type == "taskdesc" || row_type == "potentailhazard" || row_type == "possiblecase")
                                    {
                                        //int iMerge = (dtWorksheet.Select("workstep_no=" + dtWorksheet.Rows[i]["workstep_no"] + " and taskdesc_no=" + dtWorksheet.Rows[i]["taskdesc_no"] + " and potentailhazard_no=" + dtWorksheet.Rows[i]["potentailhazard_no"] + " and possiblecase_no=" + dtWorksheet.Rows[i]["possiblecase_no"])).Length;
                                        // กำหนดพารามิเตอร์ที่ต้องการกรอง
                                        var filterParameters = new Dictionary<string, object>();
                                        filterParameters.Add("workstep_no", dtWorksheet.Rows[i]["workstep_no"]);
                                        filterParameters.Add("taskdesc_no", dtWorksheet.Rows[i]["taskdesc_no"]);
                                        filterParameters.Add("potentailhazard_no", dtWorksheet.Rows[i]["potentailhazard_no"]);
                                        filterParameters.Add("possiblecase_no", dtWorksheet.Rows[i]["possiblecase_no"]);
                                        // เรียกใช้ฟังก์ชัน FilterDataTable 
                                        var (filteredRows, iMerge) = FilterDataTable(dtWorksheet, filterParameters);
                                        if (iMerge > 1)
                                        {
                                            try
                                            {
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Merge = true;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                                //icol_start += (ram_type == "5" ? 5 : 4);
                                                icol_start += 4;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Merge = true;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                                icol_start += 1;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Merge = true;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                                worksheet.Cells[startRows, icol_start, startRows + (iMerge - 1), icol_start].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            }
                                            catch (Exception) { }
                                        }
                                    }


                                }


                                icol_start = 1;
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = ((dtWorksheet.Rows[i]["workstep"] + "") == "" ? "" : (dtWorksheet.Rows[i]["workstep_no"] + "." + dtWorksheet.Rows[i]["workstep"]));
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = ((dtWorksheet.Rows[i]["taskdesc"] + "") == "" ? "" : (dtWorksheet.Rows[i]["taskdesc_no"] + "." + dtWorksheet.Rows[i]["taskdesc"]));
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["potentailhazard_no"] + "." + dtWorksheet.Rows[i]["potentailhazard"]);
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["possiblecase_no"] + "." + dtWorksheet.Rows[i]["possiblecase"]);
                                //if ((dtWorksheet.Rows[i]["id_ram"] + "") == "5")
                                //{
                                //icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["category_type"] + "");
                                //}
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["ram_befor_security"] + "");
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["ram_befor_likelihood"] + "");
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["ram_befor_risk"] + "");
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["recommendations"] + "");
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["responder_action_by"] + "");
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["ram_after_security"] + "");
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["ram_after_likelihood"] + "");
                                icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (dtWorksheet.Rows[i]["ram_after_risk"] + "");

                                icol_end = icol_start;

                                // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                                DrawTableBorders(worksheet, 12, 1, startRows - 1, icol_end - 1);


                                startRows++;
                            }


                            worksheet.Cells["A" + startRows_Def + ":O" + startRows].Style.WrapText = true;
                        }

                        //ให้แสดงเฉพาะกรณีที่ผ่านการ approve โดย safty review
                        if (bApproverSafetyReview)
                        {
                            startRows = (startRows > 36 ? startRows : 36);

                            int startRowsRP = startRows;
                            int startRowsRP_Def = startRows;
                            int endColRP = icol_end;
                            int iapprover_start_row = startRowsRP;

                            //DataRow[] drAttendees = dtMemberTeam.Select();
                            //DataRow[] drReviewer = dtRelatedPeople.Select("user_type = 'reviewer'");
                            //DataRow[] drApprover = dtRelatedPeople.Select("user_type = 'approver'");

                            var filterParameters = new Dictionary<string, object>();
                            var (drAttendees, iAttendees) = FilterDataTable(dtMemberTeam, filterParameters);

                            filterParameters = new Dictionary<string, object>();
                            filterParameters.Add("user_type", "reviewer");
                            var (drReviewer, iParameters) = FilterDataTable(dtRelatedPeople, filterParameters);

                            filterParameters = new Dictionary<string, object>();
                            filterParameters.Add("user_type", "approver");
                            var (drApprover, iApprover) = FilterDataTable(dtRelatedPeople, filterParameters);


                            //default row running  = 7 row
                            int iDefRow = 7;
                            if (drAttendees != null)
                            {
                                if (drAttendees?.Length > 0)
                                {
                                    for (int i = 0; i < drAttendees.Length; i++)
                                    {
                                        icol_start = 1;
                                        if (i >= iDefRow)
                                        {
                                            worksheet.InsertRow(startRows, 1, startRowsRP_Def);
                                        }
                                        icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (i + 1);
                                        icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (drAttendees[i]["user_displayname"] + "");
                                        icol_start += 1; worksheet.Cells[startRows, icol_start].Value = (drAttendees[i]["user_title"] + "");

                                        startRows++;
                                    }
                                }
                            }

                            //fix Safety  
                            if (drReviewer != null)
                            {
                                if (drReviewer?.Length > 0)
                                {
                                    iapprover_start_row += 1;
                                    worksheet.Cells[iapprover_start_row, endColRP - 4].Value = (drReviewer[0]["user_displayname"] + "");
                                    worksheet.Cells[iapprover_start_row, endColRP - 3].Value = (drReviewer[0]["reviewer_date"] + "");
                                    iapprover_start_row += 1;

                                }
                            }

                            //fix Other 
                            iapprover_start_row += 1;

                            if (drApprover != null)
                            {
                                if (drApprover?.Length > 0)
                                {
                                    startRowsRP_Def = iapprover_start_row + 3;
                                    iDefRow = 4;
                                    for (int i = 0; i < drApprover.Length; i++)
                                    {
                                        if (i > iDefRow && iapprover_start_row > startRowsRP_Def)
                                        {
                                            worksheet.InsertRow(iapprover_start_row, 1, startRowsRP_Def);
                                        }

                                        worksheet.Cells[iapprover_start_row, endColRP - 4].Value = (drApprover[i]["user_displayname"] + "");
                                        worksheet.Cells[iapprover_start_row, endColRP - 3].Value = (drApprover[i]["reviewer_date"] + "");

                                        iapprover_start_row++;
                                    }

                                    //DrawTableBorders(worksheet, iapprover_start_row, endColRP - 4, iapprover_start_row, endColRP - 3);
                                    //DrawTableBorders(worksheet, iapprover_start_row + 2, endColRP - 4, iapprover_start_row + 2, endColRP - 4);

                                }
                            }
                        }

                        //fix RAM = 5? แต่ไม่ต้้องแสดงใน Report
                        //if (ram_type != "5")
                        //{
                        //Delete Colums RAM
                        //worksheet.DeleteColumn(6);
                        //}
                    }
                    #endregion Worksheet

                    if (!report_all)
                    {
                        //Delete Sheet อื่นที่ไม่ใช่ WorksheetTemplate กับ Risk Assessment Matrix
                        string xWorksheetsNames = "";
                        for (int iw = 0; iw < excelPackage.Workbook.Worksheets.Count; iw++)
                        {
                            string WorksheetsNames = (excelPackage.Workbook.Worksheets[iw].Name.ToString() + "");
                            //if (!(WorksheetsNames == "WorksheetTemplate" || WorksheetsNames == "Risk Assessment Matrix"))
                            if (!(WorksheetsNames == "WorksheetTemplate"))
                            {
                                if (xWorksheetsNames != "") { xWorksheetsNames += ","; }
                                xWorksheetsNames += (excelPackage.Workbook.Worksheets[iw].Name.ToString() + "");
                            }
                        }
                        if (xWorksheetsNames != "")
                        {
                            string[] xSplitxWorksheetsNames = xWorksheetsNames.Split(",");
                            for (int isplit = 0; isplit < xSplitxWorksheetsNames.Length; isplit++)
                            {
                                string WorksheetsNames = (xSplitxWorksheetsNames[isplit] + "");
                                excelPackage.Workbook.Worksheets.Delete(WorksheetsNames);
                            }
                        }
                    }

                    excelPackage.Save();

                }
            }

            return msg_error;
        }


        public string excle_template_data_jsea(string seq, string file_fullpath_name, Boolean report_all)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            string msg_error = "";
            try
            {
                sqlstr = @" select distinct h.pha_no, g.pha_request_name, format(g.target_start_date,'dd MMMM yyyy') as target_start_date, g.id_ram
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        where h.seq = @seq ";
                sqlstr += @" order by g.pha_request_name";

                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

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

                if (dt != null)
                {
                    if (dt?.Rows.Count > 0)
                    {
                        //JSEA Study Worksheet Template.xlsx 
                        if (!string.IsNullOrEmpty(file_fullpath_name))
                        {
                            string file_fullpath_def = file_fullpath_name;
                            string folder = "jsea";//sub_software;


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

                        if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                        {
                            FileInfo template_excel = new FileInfo(file_fullpath_name);
                            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                            using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                            {
                                if (!report_all)
                                {
                                    var sheetsToDelete = excelPackage.Workbook.Worksheets
                                      .Where(sheet => sheet.Name != "WorksheetTemplate")
                                      .ToList(); // ใช้ ToList เพื่อหลีกเลี่ยงปัญหา Collection Modified

                                    foreach (var sheet in sheetsToDelete)
                                    {
                                        excelPackage.Workbook.Worksheets.Delete(sheet);
                                    }
                                }


                                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                                ExcelWorksheet worksheet = sourceWorksheet;
                                //var icol_end = 14;
                                var icol_end = 13;
                                try
                                {
                                    if (dt?.Rows.Count > 0)
                                    {
                                        // c4
                                        worksheet.Cells["C4"].Value = dt.Rows[0]["pha_request_name"].ToString();
                                        // j4 = วันที่ทำการประเมิน (Date): 23 January 2024
                                        worksheet.Cells["J4"].Value = "วันที่ทำการประเมิน (Date):" + dt.Rows[0]["target_start_date"].ToString();
                                    }
                                }
                                catch { }
                                try
                                {
                                    var startRows = 13;
                                    var endRows = 13;
                                    for (int i = 0; i < 10; i++)
                                    {
                                        worksheet.InsertRow(startRows, 1);

                                        endRows++;
                                    }
                                    DrawTableBorders(worksheet, startRows - 1, 2, endRows - 1, icol_end);
                                }
                                catch { }

                                //try
                                //{
                                //    //if (ram_type != "5")
                                //    //{
                                //    //Delete Colums RAM
                                //    //worksheet.DeleteColumn(6);
                                //    //}
                                //}
                                //catch { }


                                //Delete Sheet อื่นที่ไม่ใช่ WorksheetTemplate กับ Risk Assessment Matrix
                                string xWorksheetsNames = "";
                                for (int iw = 0; iw < excelPackage.Workbook.Worksheets.Count; iw++)
                                {
                                    string WorksheetsNames = (excelPackage.Workbook.Worksheets[iw].Name.ToString() + "");
                                    //if (!(WorksheetsNames == "WorksheetTemplate" || WorksheetsNames == "Risk Assessment Matrix"))
                                    if (!(WorksheetsNames == "WorksheetTemplate"))
                                    {
                                        if (xWorksheetsNames != "") { xWorksheetsNames += ","; }
                                        xWorksheetsNames += (excelPackage.Workbook.Worksheets[iw].Name.ToString() + "");
                                    }
                                }
                                if (xWorksheetsNames != "")
                                {
                                    string[] xSplitxWorksheetsNames = xWorksheetsNames.Split(",");
                                    for (int isplit = 0; isplit < xSplitxWorksheetsNames.Length; isplit++)
                                    {
                                        string WorksheetsNames = (xSplitxWorksheetsNames[isplit] + "");
                                        excelPackage.Workbook.Worksheets.Delete(WorksheetsNames);
                                    }
                                }

                                excelPackage.Save();
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }
            return msg_error;
        }

        public string import_excel_jsea_worksheet(string user_name, string seq, string file_fullpath_name, ref DataSet _dsData)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            List<SqlParameter> parameters = new List<SqlParameter>();
            string msg_error = "";
            try
            {
                // string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all
                #region get data
                // j4 = วันที่ทำการประเมิน (Date): 23 January 2024
                sqlstr = @" select g.pha_request_name, format(g.target_start_date,'dd MMMM yyyy') as target_start_date, g.descriptions
                         , emp.user_name, emp.user_displayname, emp.user_title, emp.user_email, lower(emp.user_name) as user_name_check
                         , g.mandatory_note
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                         left join VW_EPHA_PERSON_DETAILS emp on  lower(h.pha_request_by) = lower(emp.user_name)
                         where h.seq = @seq ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

                DataTable dtHead = new DataTable();
                //dtHead = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtHead = new DataTable();
                        dtHead = _conn.ExecuteAdapter(command).Tables[0];
                        dtHead.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                        , 0 as index_rows
                        , g.id_ram
                        , '' as workstep_no_def,'' as taskdesc_no_def,'' as potentailhazard_no_def,'' as possiblecase_no_def
                        , '' as row_type_group
                        from epha_t_header a 
                        inner join EPHA_T_GENERAL g on a.id = g.id_pha  
                        inner join EPHA_T_TASKS_WORKSHEET b on a.id  = b.id_pha
                        where 1=1 ";
                sqlstr += " and a.seq = @seq  ";
                sqlstr += " order by b.workstep_no, b.taskdesc_no, b.potentailhazard_no, b.possiblecase_no, b.category_no";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtWorksheet = new DataTable();
                //dtWorksheet = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorksheet = new DataTable();
                        dtWorksheet = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorksheet.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable
                dtWorksheet.Rows.Clear(); dtWorksheet.AcceptChanges();


                sqlstr = @" select seq,seq as id,user_id as employee_id, user_name as employee_name, user_displayname as employee_displayname, user_email as employee_email
                        , t.user_title as employee_position
                        , 'assets/img/team/avatar.webp' as employee_img, user_type as employee_type
                        , 0 as selected_type
                        , lower(user_name) as user_name_check
                        , trim(lower(replace(user_displayname,' ',''))) as employee_displayname_check
                         from VW_EPHA_PERSON_DETAILS t 
                         order by user_name";
                parameters = new List<SqlParameter>();
                DataTable dtEmployee = new DataTable();
                //dtEmployee = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtEmployee = new DataTable();
                        dtEmployee = _conn.ExecuteAdapter(command).Tables[0];
                        dtEmployee.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                #region memberteam,approver,specialist,user_displayname
                sqlstr = @"select 0 as no, 'member' as user_type, 'member' as approver_type, '' as user_name,  '' as user_displayname,  '' as user_title ,  '' as date_review ";
                parameters = new List<SqlParameter>();
                //DataTable _dtNewTable = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                DataTable _dtNewTable = new DataTable();
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
                        _dtNewTable = new DataTable();
                        _dtNewTable = _conn.ExecuteAdapter(command).Tables[0];
                        _dtNewTable.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                DataTable dtMemberteam = new DataTable();
                sqlstr = @" select c.* , 'assets/img/team/avatar.webp' as user_img, 'update' as action_type, 0 as action_change
                        ,  '' as user_title
                        from epha_t_header a 
                        inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                        inner join EPHA_T_MEMBER_TEAM c on a.id  = c.id_pha and b.id  = c.id_session";
                sqlstr += " and a.seq = @seq  ";
                sqlstr += " order by a.seq,b.seq,c.seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                dtMemberteam = new DataTable();
                //dtMemberteam = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtMemberteam = new DataTable();
                        dtMemberteam = _conn.ExecuteAdapter(command).Tables[0];
                        dtMemberteam.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                dtMemberteam.Rows.Clear();
                dtMemberteam.Columns.Add("user_type");
                dtMemberteam.Columns.Add("approver_type");


                DataTable dtApprover = new DataTable();

                sqlstr = @" select c.* , 'assets/img/team/avatar.webp' as user_img, 'update' as action_type, 0 as action_change
                        ,  '' as user_title
                        , format(c.date_review,'dd MMM yyyy') as date_review_show
                        from epha_t_header a 
                        inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                        inner join EPHA_T_APPROVER c on a.id  = c.id_pha and b.id  = c.id_session";
                sqlstr += " and a.seq = @seq  ";
                sqlstr += " order by a.seq,b.seq,c.seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                dtApprover = new DataTable();
                //dtApprover = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtApprover = new DataTable();
                        dtApprover = _conn.ExecuteAdapter(command).Tables[0];
                        dtApprover.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                dtApprover.Rows.Clear();

                DataTable dtRelatedpeople = new DataTable();
                sqlstr = @" select c.* , 'assets/img/team/avatar.webp' as user_img, 'update' as action_type, 0 as action_change
                         , format(c.date_review,'dd MMM yyyy') as date_review_show
                        from epha_t_header a 
                        inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                        inner join EPHA_T_RELATEDPEOPLE c on a.id  = c.id_pha and b.id  = c.id_session";
                sqlstr += " and a.seq = @seq  ";
                sqlstr += " order by a.seq,b.seq,c.seq";


                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                dtRelatedpeople = new DataTable();
                //dtRelatedpeople = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtRelatedpeople = new DataTable();
                        dtRelatedpeople = _conn.ExecuteAdapter(command).Tables[0];
                        dtRelatedpeople.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable
                dtRelatedpeople.Rows.Clear();

                DataTable dtRelatedpeopleOutsider = new DataTable();
                sqlstr = @" select c.* , 'assets/img/team/avatar.webp' as user_img, 'update' as action_type, 0 as action_change
                         , format(c.date_review,'dd MMM yyyy') as date_review_show
                        from epha_t_header a 
                        inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                        inner join EPHA_T_RELATEDPEOPLE_OUTSIDER c on a.id  = c.id_pha and b.id  = c.id_session";
                sqlstr += " and a.seq = @seq  ";
                sqlstr += " order by a.seq,b.seq,c.seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                dtRelatedpeopleOutsider = new DataTable();
                //dtRelatedpeopleOutsider = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtRelatedpeopleOutsider = new DataTable();
                        dtRelatedpeopleOutsider = _conn.ExecuteAdapter(command).Tables[0];
                        dtRelatedpeopleOutsider.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable
                dtRelatedpeopleOutsider.Rows.Clear();


                #endregion memberteam,approver,specialist,reviewer


                #endregion get data

                //เนื่องจาก jsea ใช้ RAM = 5 แต่ไม่ได้แสดงข้อมูล Category
                Boolean bRAM5 = true;
                int iRowIsNull = 0;

                string pha_request_name = "";
                string target_start_date = "";
                string mandatory_note = "";
                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "jsea";//sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                    {
                        ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                        ExcelWorksheet worksheet = sourceWorksheet;

                        //Worksheet
                        #region Worksheet
                        if (true)
                        {
                            //header 
                            pha_request_name = worksheet.Cells["C4"].Value?.ToString() + "";
                            target_start_date = worksheet.Cells["J4"].Value?.ToString() + "";
                            mandatory_note = worksheet.Cells["C5"].Value?.ToString() + "";

                            if (dtHead?.Rows.Count == 0) { dtHead.Rows.Add(dtHead.NewRow()); }

                            //check format date -> excel text :  20 January 2024   
                            DateTime dateTarget;
                            if (!(DateTime.TryParseExact(target_start_date, "d MMMM yyyy", System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.None, out dateTarget)))
                            {
                                //กรณีที่ไม่ใช่ format date ไม่ต้อง update ค่าจาก excel ใน table ?จะให้แสดงข้อมูลที่เคยบันทึกใน db
                                target_start_date = dtHead.Rows[0]["target_start_date"]?.ToString() ?? "";
                            }
                            dtHead.Rows[0]["pha_request_name"] = pha_request_name;
                            dtHead.Rows[0]["target_start_date"] = target_start_date;
                            dtHead.Rows[0]["mandatory_note"] = mandatory_note;
                            dtHead.AcceptChanges();

                            int irows = 0;
                            int startRows = 12;
                            int endRowsWorksheet = 1;
                            int icol_start = 1;
                            string no_ref = ""; string desc_ref = "";

                            Boolean bStart = false;

                            string row_type_group = "worksheet";

                            int iWorksheetRows = worksheet.Dimension.Rows;
                            for (int i = 1; i <= iWorksheetRows; i++) // ใช้ Dimension.Rows เพื่อหาขนาดของชีทแน่นอน
                            {
                                icol_start = 1;//เริ่มจาก row 1 = A
                                #region check เงื่อนไข Start, End Loop
                                if (true)
                                {
                                    startRows = i;
                                    if (!bStart)
                                    {
                                        try
                                        {
                                            if (worksheet.Cells[(startRows), 2].Value?.ToString().Substring(0, 5) == ("ลำดับขั้นตอนหลัก").Substring(0, 5)
                                                || worksheet.Cells[(startRows), 2].Value?.ToString().Substring(0, 5) == ("(Work step)").Substring(0, 5))
                                            {
                                                i += 1;
                                                bStart = true;
                                                continue;
                                            }
                                        }
                                        catch { }
                                    }
                                    if (!bStart) { continue; }

                                    try
                                    {
                                        if ((worksheet.Cells[(startRows), 2].Value?.ToString() ?? "").Substring(0, 5) == ("Remark").Substring(0, 5))
                                        {
                                            startRows += 2;
                                        }
                                        if ((worksheet.Cells[(startRows), 2].Value?.ToString() ?? "").Substring(0, 5) == ("ผู้จัดทำ (Attendees)").Substring(0, 5))
                                        {
                                            endRowsWorksheet = startRows + 1;
                                            row_type_group = "employee";
                                            break;
                                        }
                                    }
                                    catch { }
                                }
                                #endregion check เงื่อนไข Start, End Loop

                                //check data in row all columns
                                Boolean bCheckDataInRow = false;
                                for (int icol = 1; icol < 16; icol++)
                                {
                                    if (!string.IsNullOrEmpty(worksheet.Cells[startRows, icol].Value?.ToString()))
                                    {
                                        bCheckDataInRow = true; break;
                                    }
                                }
                                if (!bCheckDataInRow)
                                {
                                    //กรณีที่มี row ว่างมากกว่า 1 รายการให้ break
                                    if (iRowIsNull > 1) { break; }
                                    iRowIsNull += 1;
                                }
                                else if (bCheckDataInRow)
                                {
                                    //if ((worksheet.Cells[startRows, icol_start + 1].Value?.ToString() + "") == "") { continue; } 
                                    iRowIsNull = 0;

                                    no_ref = ""; desc_ref = "";
                                    dtWorksheet.Rows.Add(dtWorksheet.NewRow()); dtWorksheet.AcceptChanges();

                                    //replace no format 1. xxxx, 1.1. xxxx, 1.1.1. xxxx, 1.1.1.1. xxxx  
                                    dtWorksheet.Rows[irows]["row_type_group"] = row_type_group;

                                    icol_start += 1;
                                    fnSplit_Data_No((worksheet.Cells[startRows, icol_start].Value?.ToString() + ""), ref no_ref, ref desc_ref);

                                    dtWorksheet.Rows[irows]["workstep_no_def"] = no_ref;
                                    dtWorksheet.Rows[irows]["workstep"] = desc_ref;
                                    no_ref = ""; desc_ref = "";

                                    icol_start += 1;
                                    //fnSplit_Data_No((worksheet.Cells[startRows, icol_start].Value?.ToString() + ""), ref no_ref, ref desc_ref);
                                    desc_ref = (worksheet.Cells[startRows, icol_start].Value?.ToString() + "");
                                    dtWorksheet.Rows[irows]["taskdesc_no_def"] = no_ref;
                                    dtWorksheet.Rows[irows]["taskdesc"] = desc_ref;

                                    icol_start += 1;
                                    //fnSplit_Data_No((worksheet.Cells[startRows, icol_start].Value?.ToString() + ""), ref no_ref, ref desc_ref);
                                    desc_ref = (worksheet.Cells[startRows, icol_start].Value?.ToString() + "");
                                    dtWorksheet.Rows[irows]["potentailhazard_no_def"] = no_ref;
                                    dtWorksheet.Rows[irows]["potentailhazard"] = desc_ref;

                                    icol_start += 1;
                                    //fnSplit_Data_No((worksheet.Cells[startRows, icol_start].Value?.ToString() + ""), ref no_ref, ref desc_ref);
                                    desc_ref = (worksheet.Cells[startRows, icol_start].Value?.ToString() + "");
                                    dtWorksheet.Rows[irows]["possiblecase_no_def"] = no_ref;
                                    dtWorksheet.Rows[irows]["possiblecase"] = desc_ref;

                                    //fix
                                    dtWorksheet.Rows[irows]["category_type"] = 5;

                                    icol_start += 1; dtWorksheet.Rows[irows]["ram_befor_security"] = worksheet.Cells[startRows, icol_start].Value?.ToString();
                                    icol_start += 1; dtWorksheet.Rows[irows]["ram_befor_likelihood"] = worksheet.Cells[startRows, icol_start].Value?.ToString();
                                    icol_start += 1; dtWorksheet.Rows[irows]["ram_befor_risk"] = worksheet.Cells[startRows, icol_start].Value?.ToString();
                                    icol_start += 1; dtWorksheet.Rows[irows]["recommendations"] = worksheet.Cells[startRows, icol_start].Value?.ToString();
                                    icol_start += 1; dtWorksheet.Rows[irows]["responder_action_by"] = worksheet.Cells[startRows, icol_start].Value?.ToString();
                                    icol_start += 1; dtWorksheet.Rows[irows]["ram_after_security"] = worksheet.Cells[startRows, icol_start].Value?.ToString();
                                    icol_start += 1; dtWorksheet.Rows[irows]["ram_after_likelihood"] = worksheet.Cells[startRows, icol_start].Value?.ToString();
                                    icol_start += 1; dtWorksheet.Rows[irows]["ram_after_risk"] = worksheet.Cells[startRows, icol_start].Value?.ToString();

                                    dtWorksheet.Rows[irows]["no"] = (irows);
                                    irows += 1;
                                }

                            }

                            if (true)
                            {
                                if (endRowsWorksheet == 1)
                                {
                                    endRowsWorksheet = FindRowWithAttendees(worksheet, "B", "ผู้จัดทำ (Attendees)");
                                    endRowsWorksheet += 1;
                                }

                                //get memberteam / attendees 
                                Boolean bSpecialist = false;

                                string user_displayname_ref = "";
                                string user_title_ref = "";
                                irows = 0;
                                //bRelatedPeople = false;
                                for (int i = endRowsWorksheet; i <= iWorksheetRows; i++) // ใช้ Dimension.Rows เพื่อหาขนาดของชีทแน่นอน
                                {
                                    startRows = i;
                                    if (true)
                                    {
                                        if (!bSpecialist)
                                        {
                                            var cellValue = worksheet.Cells[startRows, 2].Value?.ToString();
                                            if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith("ผู้เชี่ยวชาญเฉพาะด้าน (ถ้ามี)"))
                                            {
                                                bSpecialist = true;
                                            }
                                        }


                                        icol_start = 3;//C = Name (ชื่อ-นามสกุล)
                                        user_displayname_ref = worksheet.Cells[(startRows), icol_start].Value?.ToString(); icol_start++;
                                        user_title_ref = worksheet.Cells[(startRows), icol_start].Value?.ToString();
                                        if (user_displayname_ref != "")
                                        {
                                            dtMemberteam.Rows.Add(dtMemberteam.NewRow());
                                            dtMemberteam.Rows[irows]["no"] = 1;
                                            dtMemberteam.Rows[irows]["user_type"] = (bSpecialist ? "specialist" : "member");
                                            dtMemberteam.Rows[irows]["user_name"] = null;
                                            dtMemberteam.Rows[irows]["user_displayname"] = user_displayname_ref;
                                            dtMemberteam.Rows[irows]["user_title"] = user_title_ref;
                                            dtMemberteam.AcceptChanges();
                                            irows += 1;
                                        }
                                        else { break; }
                                    }
                                }

                                //เนื่องจาก jsea ใช้ RAM = 5 แต่ไม่ได้แสดงข้อมูล Category
                                int icolReviewer = 8;

                                endRowsWorksheet = endRowsWorksheet + 1;
                                irows = 0;

                                for (int i = endRowsWorksheet; i <= iWorksheetRows; i++) // ใช้ Dimension.Rows เพื่อหาขนาดของชีทแน่นอน
                                {
                                    startRows = i;
                                    string approver_type = "safety";//approver,safety
                                    try
                                    {
                                        //cell 7, 6, 5
                                        for (int icell = 0; icell < 3; icell++)
                                        {
                                            var cellValue = worksheet.Cells[startRows, icolReviewer - icell].Value?.ToString();
                                            if (!string.IsNullOrEmpty(cellValue) && cellValue.Substring(0, 5) == "Safety".Substring(0, 5))
                                            {
                                                approver_type = "approver";
                                                //bRelatedPeople = true;
                                                break;
                                            }
                                        }
                                    }
                                    catch { }
                                    try
                                    {
                                        if ((worksheet.Cells[(startRows), icolReviewer].Value?.ToString() ?? "").Substring(0, 5) == ("Note:").Substring(0, 5)) { break; }
                                    }
                                    catch { }
                                    if (true)
                                    {
                                        icol_start = (icolReviewer + 1);
                                        try
                                        {
                                            if (worksheet.Cells[(startRows), icolReviewer].Value?.ToString() == "") { continue; }
                                        }
                                        catch { }

                                        user_displayname_ref = worksheet.Cells[(startRows), icol_start].Value?.ToString() ?? ""; icol_start++;
                                        user_title_ref = worksheet.Cells[(startRows), icol_start].Value?.ToString() ?? "";
                                        if (user_displayname_ref != "")
                                        {
                                            dtApprover.Rows.Add(dtApprover.NewRow());
                                            dtApprover.Rows[irows]["no"] = 1;
                                            dtApprover.Rows[irows]["approver_type"] = approver_type; //safety, member, free_text 
                                            dtApprover.Rows[irows]["user_name"] = null;
                                            dtApprover.Rows[irows]["user_displayname"] = user_displayname_ref;
                                            dtApprover.Rows[irows]["user_title"] = user_title_ref;
                                            dtApprover.AcceptChanges();
                                            irows += 1;
                                        }
                                    }
                                }

                            }

                        }
                        #endregion Worksheet

                    }

                    //set datatable return
                    DataTable dtWorksheetCopy = dtWorksheet.Clone(); dtWorksheetCopy.AcceptChanges();
                    DataTable dtMemberteamCopy = dtMemberteam.Clone(); dtMemberteamCopy.AcceptChanges();
                    DataTable dtApproverCopy = dtApprover.Clone(); dtApproverCopy.AcceptChanges();
                    DataTable dtRelatedpeopleCopy = dtRelatedpeople.Clone(); dtRelatedpeopleCopy.AcceptChanges();
                    DataTable dtRelatedpeopleOutsiderCopy = dtRelatedpeopleOutsider.Clone(); dtRelatedpeopleOutsiderCopy.AcceptChanges();

                    ClassHazop clshazop = new ClassHazop();
                    int id_tasks = (clshazop.get_max("EPHA_T_TASKS_WORKSHEET", seq));
                    int id_memberteam = (clshazop.get_max("EPHA_T_MEMBER_TEAM", seq));
                    int id_approver = (clshazop.get_max("EPHA_T_APPROVER", seq));
                    int id_relatedpeople = (clshazop.get_max("EPHA_T_RELATEDPEOPLE", seq));
                    int id_relatedpeople_outsider = (clshazop.get_max("EPHA_T_RELATEDPEOPLE_OUTSIDER", seq));

                    int id_pha = Convert.ToInt32(seq);

                    Boolean bEmployeeinList = false;

                    if (true)
                    {
                        if (true)
                        {
                            //dtMemberteam, dtSpecialist, dtReviewer 
                            string user_displayname_ref = "";
                            string user_type_ref = "";
                            int iNoMem = 1;
                            int id_session = 0; //ดึงค่า id_session จากหน้าจอ


                            // add def user request dtMemberteamCopy 
                            if (true)
                            {
                                string user_name_check = (dtHead.Rows[0]["user_name_check"] + "");
                                //DataRow[] drEmp = dtEmployee.Select("user_name_check='" + user_name_check + "'");
                                var filterParameters = new Dictionary<string, object>();
                                filterParameters.Add("user_name_check", user_name_check);
                                var (drEmp, iMerge) = FilterDataTable(dtEmployee, filterParameters);
                                if (drEmp != null)
                                {
                                    if (drEmp?.Length > 0)
                                    {
                                        dtMemberteamCopy.Rows.Add(dtMemberteamCopy.NewRow());

                                        dtMemberteamCopy.Rows[0]["id_pha"] = id_pha;
                                        dtMemberteamCopy.Rows[0]["id_session"] = id_session;
                                        dtMemberteamCopy.Rows[0]["no"] = iNoMem;

                                        dtMemberteamCopy.Rows[0]["user_img"] = "assets/img/team/avatar.webp";

                                        //dtMemberteamCopy.Rows[0]["create_by"] = user_name;
                                        dtMemberteamCopy.Rows[0]["action_type"] = "insert";
                                        dtMemberteamCopy.Rows[0]["action_change"] = 1;
                                        dtMemberteamCopy.AcceptChanges();

                                        dtMemberteamCopy.Rows[0]["seq"] = id_memberteam;
                                        dtMemberteamCopy.Rows[0]["id"] = id_memberteam;

                                        dtMemberteamCopy.Rows[0]["user_name"] = (drEmp[0]["employee_name"] + "");
                                        dtMemberteamCopy.Rows[0]["user_displayname"] = (drEmp[0]["employee_displayname"] + "");
                                        dtMemberteamCopy.Rows[0]["user_title"] = (drEmp[0]["employee_position"] + "");
                                        iNoMem += 1;
                                        dtMemberteamCopy.AcceptChanges();
                                    }
                                }
                            }

                            for (int i = 0; i < dtMemberteam?.Rows.Count; i++)
                            {
                                bEmployeeinList = false;
                                //ตรวจสอบโดยใช้ format --> Chaiyut Nontakote
                                user_displayname_ref = (dtMemberteam.Rows[i]["user_displayname"] + "").ToLower().Trim().Replace(" ", "");
                                user_type_ref = (dtMemberteam.Rows[i]["user_type"] + "").ToLower().Trim().Replace(" ", "");
                                if (user_displayname_ref != "")
                                {
                                    dtMemberteam.Rows[i]["id_pha"] = id_pha;
                                    dtMemberteam.Rows[i]["id_session"] = id_session;
                                    dtMemberteam.Rows[i]["no"] = iNoMem;
                                    dtMemberteam.Rows[i]["user_type"] = user_type_ref;
                                    dtMemberteam.Rows[i]["approver_type"] = "member";
                                    dtMemberteam.Rows[i]["user_img"] = "assets/img/team/avatar.webp";
                                    dtMemberteam.Rows[i]["action_type"] = "insert";
                                    dtMemberteam.Rows[i]["action_change"] = 1;
                                    dtMemberteam.AcceptChanges();

                                    iNoMem += 1;

                                    //dtMemberteam.Rows[irows]["user_type"] = (bSpecialist? "specialist" : "member") ; 
                                    //DataRow[] drEmp = dtEmployee.Select("employee_displayname_check='" + user_displayname_ref + "'");
                                    var filterParameters = new Dictionary<string, object>();
                                    filterParameters.Add("employee_displayname_check", user_displayname_ref);
                                    var (drEmp, iMerge) = FilterDataTable(dtEmployee, filterParameters);
                                    if (drEmp != null)
                                    {
                                        if (drEmp?.Length > 0)
                                        {
                                            bEmployeeinList = true;

                                            if (user_type_ref == "member" && bEmployeeinList)
                                            {
                                                dtMemberteam.Rows[i]["seq"] = id_memberteam;
                                                dtMemberteam.Rows[i]["id"] = id_memberteam;
                                                id_memberteam += 1;

                                                dtMemberteam.Rows[i]["user_name"] = (drEmp[0]["employee_name"] + "");
                                                dtMemberteam.Rows[i]["user_displayname"] = (drEmp[0]["employee_displayname"] + "");
                                                dtMemberteam.Rows[i]["user_title"] = (drEmp[0]["employee_position"] + "");

                                                dtMemberteamCopy.ImportRow(dtMemberteam.Rows[i]); dtMemberteamCopy.AcceptChanges();
                                                continue;
                                            }
                                            if (user_type_ref == "specialist" && bEmployeeinList)
                                            {
                                                dtMemberteam.Rows[i]["seq"] = id_relatedpeople;
                                                dtMemberteam.Rows[i]["id"] = id_relatedpeople;
                                                id_relatedpeople += 1;

                                                dtMemberteam.Rows[i]["user_name"] = (drEmp[0]["employee_name"] + "");
                                                dtMemberteam.Rows[i]["user_displayname"] = (drEmp[0]["employee_displayname"] + "");
                                                dtMemberteam.Rows[i]["user_title"] = (drEmp[0]["employee_position"] + "");

                                                dtRelatedpeopleCopy.ImportRow(dtMemberteam.Rows[i]); dtRelatedpeopleCopy.AcceptChanges();
                                                continue;
                                            }
                                        }
                                    }
                                    //Case outsider
                                    if (true)
                                    {
                                        dtMemberteam.Rows[i]["seq"] = id_relatedpeople_outsider;
                                        dtMemberteam.Rows[i]["id"] = id_relatedpeople_outsider;
                                        dtMemberteam.Rows[i]["approver_type"] = "free_text";
                                        id_relatedpeople_outsider += 1;

                                        dtRelatedpeopleOutsiderCopy.ImportRow(dtMemberteam.Rows[i]); dtRelatedpeopleOutsiderCopy.AcceptChanges();
                                    }
                                }
                            }


                            iNoMem = 1;
                            for (int i = 0; i < dtApprover?.Rows.Count; i++)
                            {
                                bEmployeeinList = false;
                                //ตรวจสอบโดยใช้ format --> Chaiyut Nontakote
                                user_displayname_ref = (dtApprover.Rows[i]["user_displayname"] + "").ToLower().Trim().Replace(" ", "");
                                if (user_displayname_ref != "")
                                {
                                    //DataRow[] drEmp = dtEmployee.Select("employee_displayname_check='" + user_displayname_ref + "'");
                                    //if (drEmp.Length > 0)
                                    //{
                                    //    bEmployeeinList = true;
                                    //}
                                    dtApprover.Rows[i]["id_pha"] = id_pha;
                                    dtApprover.Rows[i]["id_session"] = id_session;
                                    dtApprover.Rows[i]["no"] = iNoMem;
                                    dtApprover.Rows[i]["user_img"] = "assets/img/team/avatar.webp";

                                    //dtApprover.Rows[i]["create_by"] = user_name;
                                    dtApprover.Rows[i]["action_type"] = "insert";
                                    dtApprover.Rows[i]["action_change"] = 1;

                                    dtApprover.AcceptChanges();
                                    iNoMem += 1;

                                    var filterParameters = new Dictionary<string, object>();
                                    filterParameters.Add("employee_displayname_check", user_displayname_ref);
                                    var (drEmp, iMerge) = FilterDataTable(dtEmployee, filterParameters);
                                    if (drEmp != null)
                                    {
                                        if (drEmp?.Length > 0)
                                        {
                                            dtApprover.Rows[i]["seq"] = id_approver;
                                            dtApprover.Rows[i]["id"] = id_approver;
                                            id_approver += 1;

                                            dtApprover.Rows[i]["user_name"] = (drEmp[0]["employee_name"] + "");
                                            dtApprover.Rows[i]["user_displayname"] = (drEmp[0]["employee_displayname"] + "");
                                            dtApprover.Rows[i]["user_title"] = (drEmp[0]["employee_position"] + "");
                                            dtApproverCopy.ImportRow(dtApprover.Rows[i]); dtApproverCopy.AcceptChanges();
                                        }
                                    }

                                    if (!bEmployeeinList)
                                    {
                                        dtApprover.Rows[i]["seq"] = id_relatedpeople_outsider;
                                        dtApprover.Rows[i]["id"] = id_relatedpeople_outsider;
                                        id_relatedpeople_outsider += 1;

                                        dtRelatedpeopleCopy.ImportRow(dtApprover.Rows[i]); dtRelatedpeopleCopy.AcceptChanges();
                                    }

                                }
                            }


                            //Set Data 
                            for (int i = 0; i < dtMemberteamCopy?.Rows.Count; i++)
                            {
                                dtMemberteamCopy.Rows[i]["no"] = (i + 1);
                                dtMemberteamCopy.AcceptChanges();
                            }
                            for (int i = 0; i < dtApproverCopy?.Rows.Count; i++)
                            {
                                dtApproverCopy.Rows[i]["no"] = (i + 1);
                                dtApproverCopy.AcceptChanges();
                            }
                            for (int i = 0; i < dtRelatedpeopleCopy?.Rows.Count; i++)
                            {
                                dtRelatedpeopleCopy.Rows[i]["no"] = (i + 1);
                                dtRelatedpeopleCopy.AcceptChanges();
                            }
                            for (int i = 0; i < dtRelatedpeopleOutsiderCopy?.Rows.Count; i++)
                            {
                                dtRelatedpeopleOutsiderCopy.Rows[i]["no"] = (i + 1);
                                dtRelatedpeopleOutsiderCopy.AcceptChanges();
                            }

                        }

                        if (dtWorksheet?.Rows.Count > 0)
                        {
                            string row_type = "";

                            Decimal workstep_no = 0;
                            Decimal taskdesc_no = 0;
                            Decimal potentailhazard_no = 0;
                            Decimal possiblecase_no = 0;
                            Decimal category_no = 0;

                            string workstep = "";
                            string taskdesc = "";
                            string potentailhazard = "";
                            string possiblecase = "";

                            int seq_workstep = 1;
                            int seq_taskdesc = 1;
                            int seq_potentailhazard = 1;
                            int seq_possiblecase = 1;
                            int seq_category = 1;

                            //runngin no auto
                            //1. xx
                            //1.1. xx
                            //1.1.1. xx
                            //1.1.1.1. xx 

                            //Stamp value, no
                            for (int i = 0; i < dtWorksheet?.Rows.Count; i++)
                            {
                                //if (!((dtWorksheet.Rows[i]["row_type_group"] + "") == "worksheet")) { break; } 

                                string row_type_def = (dtWorksheet.Rows[i]["row_type"] + "");

                                string workstep_def = (dtWorksheet.Rows[i]["workstep"] + "");
                                string taskdesc_def = (dtWorksheet.Rows[i]["taskdesc"] + "");
                                string potentailhazard_def = (dtWorksheet.Rows[i]["potentailhazard"] + "");
                                string possiblecase_def = (dtWorksheet.Rows[i]["possiblecase"] + "");

                                string category_type_def = (dtWorksheet.Rows[i]["category_type"] + "");
                                if (!bRAM5) { category_type_def = ""; }

                                if (workstep_def == "" && taskdesc_def == "" && potentailhazard_def == "" && possiblecase_def == "")
                                {
                                    category_no += 1;
                                    seq_category += 1;
                                    row_type = "category";
                                }
                                else if (workstep_def == "" && taskdesc_def == "" && potentailhazard_def == "" && possiblecase_def != "")
                                {
                                    possiblecase_no += 1; category_no = 1;
                                    seq_possiblecase += 1;
                                    row_type = "possiblecase";
                                }
                                else if (workstep_def == "" && taskdesc_def == "" && potentailhazard_def != "")
                                {
                                    potentailhazard_no += 1; possiblecase_no = 1; category_no = 1;
                                    seq_potentailhazard += 1;
                                    row_type = "potentailhazard";
                                }
                                else if (workstep_def == "" && taskdesc_def != "")
                                {
                                    taskdesc_no += 1; potentailhazard_no = 1; possiblecase_no = 1; category_no = 1;
                                    seq_taskdesc += 1;
                                    row_type = "taskdesc";
                                }
                                else if (workstep_def != "")
                                {
                                    workstep_no += 1; taskdesc_no = 1; potentailhazard_no = 1; possiblecase_no = 1; category_no = 1;
                                    seq_workstep += 1;
                                    row_type = "workstep";
                                }

                                workstep = (workstep_def == "" ? workstep : dtWorksheet.Rows[i]["workstep"] + "");
                                taskdesc = (taskdesc_def == "" ? taskdesc : dtWorksheet.Rows[i]["taskdesc"] + "");
                                potentailhazard = (potentailhazard_def == "" ? potentailhazard : dtWorksheet.Rows[i]["potentailhazard"] + "");
                                possiblecase = (possiblecase_def == "" ? possiblecase : dtWorksheet.Rows[i]["possiblecase"] + "");

                                dtWorksheet.Rows[i]["seq_workstep"] = seq_workstep;
                                dtWorksheet.Rows[i]["seq_taskdesc"] = seq_taskdesc;
                                dtWorksheet.Rows[i]["seq_potentailhazard"] = seq_potentailhazard;
                                dtWorksheet.Rows[i]["seq_possiblecase"] = seq_possiblecase;
                                dtWorksheet.Rows[i]["seq_category"] = seq_category;

                                dtWorksheet.Rows[i]["workstep_no"] = workstep_no;
                                dtWorksheet.Rows[i]["taskdesc_no"] = taskdesc_no;
                                dtWorksheet.Rows[i]["potentailhazard_no"] = potentailhazard_no;
                                dtWorksheet.Rows[i]["possiblecase_no"] = possiblecase_no;

                                dtWorksheet.Rows[i]["workstep"] = workstep;
                                dtWorksheet.Rows[i]["taskdesc"] = taskdesc;
                                dtWorksheet.Rows[i]["potentailhazard"] = potentailhazard;
                                dtWorksheet.Rows[i]["possiblecase"] = possiblecase;

                                //Details
                                dtWorksheet.Rows[i]["seq"] = id_tasks;
                                dtWorksheet.Rows[i]["id"] = id_tasks;
                                dtWorksheet.Rows[i]["id_pha"] = id_pha;

                                dtWorksheet.Rows[i]["index_rows"] = i;//ใช้ในการค้นหาลำดับ

                                dtWorksheet.Rows[i]["no"] = (i + 1);
                                dtWorksheet.Rows[i]["row_type"] = row_type;//workstep,taskdesc,potentailhazard,possiblecase,category

                                dtWorksheet.Rows[i]["action_status"] = "open";
                                dtWorksheet.Rows[i]["action_type"] = "insert";
                                dtWorksheet.Rows[i]["action_change"] = 1;
                                dtWorksheet.AcceptChanges();

                                dtWorksheetCopy.ImportRow(dtWorksheet.Rows[i]); dtWorksheetCopy.AcceptChanges();

                                id_tasks += 1;

                            }
                        }

                    }
                    if (true)
                    {
                        _dsData = new DataSet();
                        clshazop = new ClassHazop();
                        DataTable dtma = new DataTable();

                        clshazop.set_max_id(ref dtma, "tasks_worksheet", (id_tasks + 1).ToString());
                        clshazop.set_max_id(ref dtma, "memberteam", (id_memberteam + 1).ToString());
                        clshazop.set_max_id(ref dtma, "approver", (id_approver + 1).ToString());
                        clshazop.set_max_id(ref dtma, "relatedpeople", (id_relatedpeople + 1).ToString());
                        clshazop.set_max_id(ref dtma, "relatedpeople_outsider", (id_relatedpeople_outsider + 1).ToString());

                        dtma.TableName = "max";
                        _dsData.Tables.Add(dtma.Copy()); _dsData.AcceptChanges();

                        dtWorksheetCopy.TableName = "tasks_worksheet";
                        _dsData.Tables.Add(dtWorksheetCopy.Copy()); _dsData.AcceptChanges();

                        dtMemberteamCopy.TableName = "memberteam";
                        _dsData.Tables.Add(dtMemberteamCopy.Copy()); _dsData.AcceptChanges();

                        dtApproverCopy.TableName = "approver";
                        _dsData.Tables.Add(dtApproverCopy.Copy()); _dsData.AcceptChanges();

                        dtRelatedpeopleCopy.TableName = "relatedpeople";
                        _dsData.Tables.Add(dtRelatedpeopleCopy.Copy()); _dsData.AcceptChanges();

                        dtRelatedpeopleOutsiderCopy.TableName = "relatedpeople_outsider";
                        _dsData.Tables.Add(dtRelatedpeopleOutsiderCopy.Copy()); _dsData.AcceptChanges();

                        DataTable dtgeneral = new DataTable();
                        dtgeneral.Columns.Add("pha_request_name");
                        dtgeneral.Columns.Add("target_start_date", typeof(DateTime));
                        dtgeneral.Columns.Add("mandatory_note");

                        try
                        {
                            // target_start_date ตัวอย่างวันที่ในรูปแบบ "dd/MM/yyyy"
                            //DateTime dNow = DateTime.ParseExact(target_start_date, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                            DateTime dNow = DateTime.ParseExact(target_start_date, "d MMMM yyyy", CultureInfo.InvariantCulture);
                            dtgeneral.Rows.Add(pha_request_name, dNow, mandatory_note);
                        }
                        catch { }
                        dtgeneral.TableName = "general";
                        _dsData.Tables.Add(dtgeneral.Copy()); _dsData.AcceptChanges();
                    }
                } 

            }
            catch (Exception e) { msg_error = e.Message.ToString(); }

            return msg_error;

        }

        #endregion export excel jsea

        #region export excel what'if
        public string excel_whatif_general(string seq, string file_fullpath_name)
        {

            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            try
            {
                List<SqlParameter> parameters = new List<SqlParameter>();
                #region get data
                sqlstr = @" select g.work_scope, h.pha_no, h.pha_version_text as pha_version, g.descriptions, format(g.target_start_date, 'dd MMM yyyy ') as target_start_date 
                         ,  case when ums.user_name is null  then h.request_user_displayname else case when ums.departments is null  then  ums.user_displayname else  ums.user_displayname + ' (' + ums.departments +')' end end request_user_displayname
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         left join VW_EPHA_PERSON_DETAILS ums on lower(h.request_user_name) = lower(ums.USER_NAME) 
                         where h.seq = @seq ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtWorkScope = new DataTable();
                //dtWorkScope = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorkScope = new DataTable();
                        dtWorkScope = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorkScope.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                sqlstr = @" select distinct d.no, d.document_name, d.document_no, d.document_file_name, d.descriptions 
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                        inner join EPHA_T_DRAWING d on h.id = d.id_pha    
                        where h.seq = @seq and d.document_name is not null order by convert(int,d.no) ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
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
                        dtDrawing = new DataTable();
                        dtDrawing = _conn.ExecuteAdapter(command).Tables[0];
                        dtDrawing.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                sqlstr = @" select nl.no, nl.list , nl.design_intent, nl.design_conditions, nl.operating_conditions, nl.list_boundary 
                         , d.document_no
                         , isnull(replace(replace( convert(char,nd.page_start_first) + (case when isnull(nd.page_start_first,'') ='' then '' else
                         (case when isnull(nd.page_end_first,'') ='' then '' else 'to'end)  end) 
                         + convert(char,nd.page_end_first)  ,' ',''),'to',' to '),'All') as  document_page
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_LIST nl on h.id = nl.id_pha 
                         left join EPHA_T_LIST_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_list 
                         left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                         where h.seq = @seq  and nl.list is not null order by convert(int,nl.no), convert(int,nd.no) ";


                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtNode = new DataTable();
                //dtNode = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtNode = new DataTable();
                        dtNode = _conn.ExecuteAdapter(command).Tables[0];
                        dtNode.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                #endregion get data

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "whatif";//sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))


                    {
                        ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];  // Replace "SourceSheet" with the actual source sheet name
                        ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                        //Whatif Cover Page
                        worksheet = excelPackage.Workbook.Worksheets["What'if Cover Page"];

                        ClassHazop clshazop = new ClassHazop();
                        worksheet.Cells["A14"].Value = clshazop.convert_revision_text(dtWorkScope.Rows[0]["pha_version"] + "");
                        worksheet.Cells["B14"].Value = (dtWorkScope.Rows[0]["target_start_date"] + "");
                        worksheet.Cells["C14"].Value = (dtWorkScope.Rows[0]["request_user_displayname"] + "");
                        worksheet.Cells["D14"].Value = (dtWorkScope.Rows[0]["descriptions"] + "");


                        //Study Objective and Work Scope
                        worksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];
                        worksheet.Cells["A2"].Value = (dtWorkScope.Rows[0]["work_scope"] + "");

                        //Description/Remarks 
                        if ((dtWorkScope.Rows[0]["descriptions"] + "") == "")
                        {
                            //remove rows A4, A5
                            worksheet.DeleteRow(5); worksheet.DeleteRow(4);
                        }
                        else
                        {
                            worksheet.Cells["A5"].Value = (dtWorkScope.Rows[0]["descriptions"] + "");
                        }

                        //Drawing & Reference
                        #region Drawing & Reference
                        if (true)
                        {
                            worksheet = excelPackage.Workbook.Worksheets["Drawing & Reference"];

                            int startRows = 3;
                            int icol_end = 6;
                            int ino = 1;
                            for (int i = 0; i < dtDrawing?.Rows.Count; i++)
                            {
                                //No.	Document Name	Drawing No	Document File	Comment
                                worksheet.InsertRow(startRows, 1);
                                worksheet.Cells["A" + (startRows)].Value = (i + 1); ;
                                worksheet.Cells["B" + (startRows)].Value = (dtDrawing.Rows[i]["document_name"] + "");
                                worksheet.Cells["C" + (startRows)].Value = (dtDrawing.Rows[i]["document_no"] + "");
                                worksheet.Cells["D" + (startRows)].Value = (dtDrawing.Rows[i]["document_file_name"] + "");
                                worksheet.Cells["E" + (startRows)].Value = (dtDrawing.Rows[i]["descriptions"] + "");
                                startRows++;
                            }
                            // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                            DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                            //var eRange = worksheet.Cells[worksheet.Cells["A3"].Address + ":" + worksheet.Cells["D" + startRows].Address];
                            //eRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            //eRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        #endregion Drawing & Reference

                        //Task List
                        #region Task List
                        if (true)
                        {
                            worksheet = excelPackage.Workbook.Worksheets["Task List"];

                            int startRows = 3;
                            int icol_end = 9;
                            for (int i = 0; i < dtNode?.Rows.Count; i++)
                            {
                                //No.	Node	Design Intent	Design Conditions	Operating Conditions	Node Boundary	Drawing	Drawing Page (From-To)
                                worksheet.InsertRow(startRows, 1);
                                worksheet.Cells["A" + (startRows)].Value = (i + 1);
                                worksheet.Cells["B" + (startRows)].Value = (dtNode.Rows[i]["list"] + "");
                                worksheet.Cells["C" + (startRows)].Value = (dtNode.Rows[i]["design_intent"] + "");
                                worksheet.Cells["D" + (startRows)].Value = (dtNode.Rows[i]["design_conditions"] + "");
                                worksheet.Cells["E" + (startRows)].Value = (dtNode.Rows[i]["operating_conditions"] + "");
                                worksheet.Cells["F" + (startRows)].Value = (dtNode.Rows[i]["list_boundary"] + "");
                                worksheet.Cells["G" + (startRows)].Value = (dtNode.Rows[i]["document_no"] + "");
                                worksheet.Cells["H" + (startRows)].Value = (dtNode.Rows[i]["document_page"] + "");

                                startRows++;
                            }
                            // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                            DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);
                        }
                        #endregion Task List

                        // Major Accident Event (MAE),
                        #region Major Accident Event (MAE)
                        //if (true)
                        //{
                        //    worksheet = excelPackage.Workbook.Worksheets["Major Accident Event (MAE)"];

                        //    int startRows = 3;
                        //    int icol_end = 6;
                        //    for (int i = 0; i < dtMajor?.Rows.Count; i++)
                        //    {
                        //        //No.	 nl.node, nw.causes, nw.causes_no, nw.ram_befor_risk
                        //        worksheet.InsertRow(startRows, 1);
                        //        worksheet.Cells["A" + (i + startRows)].Value = (i + 1);
                        //        worksheet.Cells["B" + (i + startRows)].Value = (dtMajor.Rows[i]["list"] + "");
                        //        worksheet.Cells["C" + (i + startRows)].Value = (dtMajor.Rows[i]["causes"] + "");
                        //        worksheet.Cells["D" + (i + startRows)].Value = (dtMajor.Rows[i]["ram_befor_risk"] + "");

                        //        startRows++;
                        //    }
                        //    // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                        //    DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);
                        //}
                        #endregion Node List



                        //Study Objective and Work Scope
                        #region Study Objective and Work Scope
                        if (true)
                        {
                            worksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];
                            worksheet.Cells["A2"].Value = (dtWorkScope.Rows[0]["work_scope"] + "");

                            //Description/Remarks 
                            if ((dtWorkScope.Rows[0]["descriptions"] + "") == "")
                            {
                                //remove rows A4, A5
                                worksheet.DeleteRow(5); worksheet.DeleteRow(4);
                            }
                            else
                            {
                                worksheet.Cells["A5"].Value = (dtWorkScope.Rows[0]["descriptions"] + "");
                            }
                        }
                        #endregion Study Objective and Work Scope


                        excelPackage.Save();
                    }
                }

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }
        public string excel_whatif_atendeesheet(string seq, string file_fullpath_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            try
            {

                List<SqlParameter> parameters = new List<SqlParameter>();
                sqlstr = @" select distinct a.id as id_pha, a.pha_status, ta2.no
                     , isnull(emp.user_name,'') as user_name, emp.user_displayname, emp.user_email
                     , case when ta2.action_review = 2 then (case when ta2.action_status = 'approve' then 'Approve' else
						(case when ta2.action_status = 'reject' and ta2.comment is not null  then 'Send Back' else 'Send Back with comment' end) 
					  end) else '' end action_status
					  
                     from epha_t_header a  
                     inner join EPHA_T_GENERAL g on a.id = g.id_pha   
                     inner join EPHA_T_SESSION s on a.id = s.id_pha 
                     inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s2 on a.id = s2.id_pha and s.id = s2.id_session 
                     inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha ) t2 on a.id = t2.id_pha and s2.id_session = t2.id_session 
                     inner join EPHA_T_APPROVER ta2 on a.id = ta2.id_pha and t2.id_pha = ta2.id_pha and t2.id_session = ta2.id_session 
                     left join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name) 
					 inner join  (select max(seq)as seq, pha_no from epha_t_header group by pha_no) hm on a.seq = hm.seq and a.pha_no = hm.pha_no
                     where a.request_approver = 1 and a.seq = @seq order by ta2.no";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtTA2 = new DataTable();
                //dtTA2 = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtTA2 = new DataTable();
                        dtTA2 = _conn.ExecuteAdapter(command).Tables[0];
                        dtTA2.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                sqlstr = $@"select * from (
                         select h.seq, s.id_pha, s.seq as seq_session, s.no as session_no
                         , convert(varchar,s.meeting_date,106) as meeting_date
                         , mt.no as member_no, isnull(mt.user_name,'') as user_name, emp.user_displayname
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_SESSION s on h.id = s.id_pha 
                         left join EPHA_T_MEMBER_TEAM mt on h.id = mt. id_pha and mt.id_session = s.id
                         left join VW_EPHA_PERSON_DETAILS emp on lower(emp.user_name) = lower(mt.user_name)
                         where lower(mt.user_name) is not null 
                        )t where t.seq = @seq ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

                DataTable dtAll = new DataTable();
                //dtAll = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtAll = new DataTable();
                        dtAll = _conn.ExecuteAdapter(command).Tables[0];
                        dtAll.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = $@" select distinct 0 as no, t.user_name, t.user_displayname, '' as company_text from (
                         select h.seq, s.id_pha, s.seq as seq_session, s.no as session_no
                         , convert(varchar,s.meeting_date,106) as meeting_date
                         , mt.no as member_no, isnull(mt.user_name,'') as user_name, emp.user_displayname
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_SESSION s on h.id = s.id_pha 
                         left join EPHA_T_MEMBER_TEAM mt on h.id = mt. id_pha and mt.id_session = s.id
                         left join VW_EPHA_PERSON_DETAILS emp on lower(emp.user_name) = lower(mt.user_name)
                         where lower(mt.user_name) is not null 
                        )t where t.seq = @seq and t.user_name <> '' order by t.user_name";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtMember = new DataTable();
                //dtMember = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtMember = new DataTable();
                        dtMember = _conn.ExecuteAdapter(command).Tables[0];
                        dtMember.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = $@" select distinct t.seq_session, t.session_no, t.meeting_date from (
                         select h.seq, s.id_pha, s.seq as seq_session, s.no as session_no
                         , convert(varchar,s.meeting_date,106) as meeting_date
                         , mt.no as member_no, isnull(mt.user_name,'') as user_name, emp.user_displayname
                         from epha_t_header h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_SESSION s on h.id = s.id_pha 
                         left join EPHA_T_MEMBER_TEAM mt on h.id = mt. id_pha and mt.id_session = s.id
                         left join VW_EPHA_PERSON_DETAILS emp on lower(emp.user_name) = lower(mt.user_name)
                         where lower(mt.user_name) is not null 
                        )t where t.seq = @seq order by t.session_no ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtSession = new DataTable();
                //dtSession = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                       dtSession = new DataTable();
                       dtSession = _conn.ExecuteAdapter(command).Tables[0];
                        dtSession.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "whatif";//sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))


                    {
                        ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];  // Replace "SourceSheet" with the actual source sheet name
                        sourceWorksheet.Name = "Whatif Attendee Sheet";
                        ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                        int i = 0;
                        int startRows = 4;
                        int icol_start = 4;
                        int icol_end = 0;// icol_start + (dtSession.Rows.Count > 6 ? dtSession.Rows.Count : 6);
                        if (dtSession?.Rows.Count > 0)
                        {
                            int icount_session = dtSession?.Rows.Count ?? 0;

                            icol_end = icol_start + (icount_session > 6 ? icount_session : 6);
                        }

                        for (int imember = 0; imember < dtMember?.Rows.Count; imember++)
                        {
                            worksheet.InsertRow(startRows, 1);
                            string user_name = (dtMember.Rows[imember]["user_name"] + "");
                            //No.
                            worksheet.Cells["A" + (i + startRows)].Value = (imember + 1);
                            //Name
                            worksheet.Cells["B" + (i + startRows)].Value = (dtMember.Rows[imember]["user_displayname"] + "");
                            //Company
                            worksheet.Cells["C" + (i + startRows)].Value = (dtMember.Rows[imember]["company_text"] + "");

                            int irow_session = 0;
                            if (imember == 0)
                            {
                                if (dtSession?.Rows.Count < 6)
                                {
                                    //worksheet.Cells[2, icol_start, 2, icol_end].Merge = true; 
                                    for (int c = icol_end; c < 30; c++)
                                    {
                                        worksheet.DeleteColumn(icol_end);

                                    }
                                }

                                irow_session = 0;
                                for (int c = icol_start; c < icol_end; c++)
                                {
                                    try
                                    {
                                        //header 
                                        if ((dtSession.Rows[irow_session]["meeting_date"] + "") == "")
                                        {
                                            worksheet.Cells[3, c].Value = "";
                                        }
                                        else
                                        {
                                            worksheet.Cells[3, c].Value = (dtSession.Rows[irow_session]["meeting_date"] + "");
                                        }
                                    }
                                    catch { worksheet.Cells[3, c].Value = ""; }
                                    irow_session += 1;
                                }
                            }

                            irow_session = 0;
                            for (int c = icol_start; c < icol_end; c++)
                            {
                                try
                                {
                                    string session_no = "";
                                    try { session_no = (dtSession.Rows[irow_session]["session_no"] + ""); } catch { }

                                    //DataRow[] dr = dtAll.Select("user_name = '" + user_name + "' and session_no = '" + session_no + "'");
                                    //if (dr.Length > 0)
                                    //{
                                    //    worksheet.Cells[startRows, c].Value = "X";
                                    //}
                                    //else { worksheet.Cells[startRows, c].Value = ""; }
                                    worksheet.Cells[startRows, c].Value = "";

                                    //DataRow[] dr = dtAll.Select("user_name = '" + user_name + "' and session_no = '" + session_no + "'");
                                    var filterParameters = new Dictionary<string, object>();
                                    filterParameters.Add("user_name", user_name);
                                    filterParameters.Add("session_no", session_no);
                                    var (dr, iMerge) = FilterDataTable(dtAll, filterParameters);
                                    if (dr != null)
                                    {
                                        if (dr?.Length > 0)
                                        {
                                            worksheet.Cells[startRows, c].Value = "X";
                                        }
                                    }
                                }
                                catch { }
                                irow_session++;

                            }

                            startRows++;
                        }
                        // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                        DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                        //TA2 
                        startRows += 13;
                        int startRows_ta2 = startRows;

                        for (int ita2 = 0; ita2 < dtTA2?.Rows.Count; ita2++)
                        {
                            worksheet.InsertRow(startRows, 1);
                            //No.
                            worksheet.Cells["A" + (i + startRows)].Value = (ita2 + 1);
                            //Name
                            worksheet.Cells["B" + (i + startRows)].Value = (dtTA2.Rows[ita2]["user_displayname"] + "");
                            //status
                            worksheet.Cells["C" + (i + startRows)].Value = (dtTA2.Rows[ita2]["action_status"] + "");

                            startRows++;
                        }
                        // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                        DrawTableBorders(worksheet, startRows_ta2, 1, startRows - 1, 3);

                        excelPackage.Save();
                    }
                }
       

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }
        public string excel_whatif_worksheet(string seq, string file_fullpath_name, Boolean report_all)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            try
            {

                #region Get Data
                List<SqlParameter> parameters = new List<SqlParameter>();

                sqlstr = @" select distinct nl.no, nl.id as id_list
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_LIST nl on h.id = nl.id_pha 
                        left join EPHA_T_LIST_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_list 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_LIST_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_list   
                        where h.seq = @seq ";
                sqlstr += @" order by cast(nl.no as int)";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtList = new DataTable();
                //dtList = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtList = new DataTable();
                        dtList = _conn.ExecuteAdapter(command).Tables[0];
                        dtList.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = @" select distinct
                        h.seq, nl.id as id_list, g.pha_request_name, convert(varchar,g.create_date,106) as create_date
						, nl.list, nl.design_intent, nl.descriptions as descriptions_worksheet, nl.design_conditions, nl.list_boundary, nl.operating_conditions
                        , d.document_no
                        , nw.list_system, nw.list_sub_system
                        , nw.causes, nw.consequences, nw.category_type, nw.ram_befor_security, nw.ram_befor_likelihood, nw.ram_befor_risk
                        , nw.major_accident_event, nw.safety_critical_equipment, nw.safety_critical_equipment_tag, nw.existing_safeguards, nw.ram_after_security, nw.ram_after_likelihood, nw.ram_after_risk
                        , nw.recommendations, nw.recommendations_no, nw.responder_user_displayname
                        , g.descriptions
                        , nl.no as list_no, nw.no, nw.list_system_no, nw.list_sub_system_no, nw.causes_no, nw.consequences_no, nw.category_no 
                        , case when g.id_ram = 5 then 1 else 0 end show_cat
                        , h.safety_critical_equipment_show
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_LIST nl on h.id = nl.id_pha 
                        left join EPHA_T_LIST_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_list
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_LIST_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_list   
                        where h.seq = @seq ";
                sqlstr += @" order by cast(nl.no as int),cast(nw.list_system_no as int)
                        , cast(nw.list_sub_system_no as int), cast(nw.causes_no as int), cast(nw.consequences_no as int), cast(nw.category_no as int) ";


                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtWorksheet = new DataTable();
                //dtWorksheet = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorksheet = new DataTable();
                        dtWorksheet = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorksheet.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable



                #endregion Get Data

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "whatif";//sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                    {
                        if (!report_all)
                        {
                            var sheetsToDelete = excelPackage.Workbook.Worksheets
                              .Where(sheet => sheet.Name != "WorksheetTemplate")
                              .ToList(); // ใช้ ToList เพื่อหลีกเลี่ยงปัญหา Collection Modified

                            foreach (var sheet in sheetsToDelete)
                            {
                                excelPackage.Workbook.Worksheets.Delete(sheet);
                            }
                        }

                        Boolean sce_show = false;
                        string worksheet_name = "";
                        string worksheet_name_target = "";
                        for (int ilist = 0; ilist < dtList?.Rows.Count; ilist++)
                        {
                            if (worksheet_name_target == "") { worksheet_name_target = "WorksheetTemplate"; }
                            else { worksheet_name_target = "What's If Worksheet List (" + (ilist) + ")"; }
                            worksheet_name = "What's If Worksheet Node (" + (ilist + 1) + ")";

                            ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets[(sce_show ? "WorksheetTemplateSCE" : "WorksheetTemplate")];  // Replace "SourceSheet" with the actual source sheet name
                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(worksheet_name, sourceWorksheet);


                            string id_list = (dtList.Rows[ilist]["id_list"] + "");

                            int i = 0;
                            int startRows = 3;
                            int startRows_Def = startRows;

                            //DataRow[] dr = dtWorksheet.Select("id_list=" + id_list);
                            var filterParameters = new Dictionary<string, object>();
                            filterParameters.Add("id_list", id_list);
                            var (dr, iMerge) = FilterDataTable(dtWorksheet, filterParameters);
                            if (dr != null)
                            {
                                if (dr?.Length > 0)
                                {
                                    string show_cat = (dr[0]["show_cat"] + "");
                                    if (dr.Length > 0)
                                    {
                                        #region head text
                                        string cell_h_end = (sce_show == true ? "O" : "N");
                                        i = 0;
                                        //Project
                                        worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["pha_request_name"] + "");
                                        //List
                                        worksheet.Cells[cell_h_end + (i + startRows)].Value = (dr[0]["list"] + "");
                                        startRows++;

                                        //Design Intent :
                                        worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["design_intent"] + "");
                                        //System
                                        worksheet.Cells[cell_h_end + (i + startRows)].Value = (dr[0]["descriptions"] + "");
                                        startRows++;

                                        //"Design Conditions: -->design_conditions
                                        worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["design_conditions"] + "");

                                        //List Boundary
                                        worksheet.Cells[cell_h_end + (i + startRows)].Value = (dr[0]["list_boundary"] + "");
                                        startRows++;

                                        //"Operating Conditions: -->operating_conditions
                                        worksheet.Cells[cell_h_end + (i + startRows)].Value = (dr[0]["operating_conditions"] + "");
                                        startRows++;

                                        //PFD, PID No. : --> document_no
                                        worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["document_no"] + "");
                                        //Date
                                        worksheet.Cells[cell_h_end + (i + startRows)].Value = (dr[0]["create_date"] + "");
                                        startRows++;

                                        #endregion head text
                                        startRows = 14;
                                        for (i = 0; i < dr.Length; i++)
                                        {
                                            worksheet.InsertRow(startRows, 1);

                                            worksheet.Cells["A" + (startRows)].Value = dr[i]["list_system"].ToString();
                                            worksheet.Cells["B" + (startRows)].Value = dr[i]["list_sub_system"].ToString();
                                            worksheet.Cells["C" + (startRows)].Value = dr[i]["causes"].ToString();
                                            worksheet.Cells["D" + (startRows)].Value = dr[i]["consequences"].ToString();

                                            worksheet.Cells["E" + (startRows)].Value = dr[i]["category_type"].ToString();

                                            worksheet.Cells["F" + (startRows)].Value = dr[i]["ram_befor_security"].ToString();
                                            worksheet.Cells["G" + (startRows)].Value = dr[i]["ram_befor_likelihood"].ToString();
                                            worksheet.Cells["H" + (startRows)].Value = dr[i]["ram_befor_risk"];
                                            worksheet.Cells["I" + (startRows)].Value = dr[i]["existing_safeguards"].ToString();

                                            worksheet.Cells["J" + (startRows)].Value = dr[i]["ram_after_security"].ToString();

                                            worksheet.Cells["K" + (startRows)].Value = dr[i]["ram_after_likelihood"].ToString();
                                            worksheet.Cells["L" + (startRows)].Value = dr[i]["ram_after_risk"].ToString();
                                            worksheet.Cells["M" + (startRows)].Value = dr[i]["recommendations_no"].ToString();
                                            worksheet.Cells["N" + (startRows)].Value = dr[i]["recommendations"].ToString();
                                            worksheet.Cells["O" + (startRows)].Value = dr[i]["responder_user_displayname"].ToString();


                                            startRows++;
                                        }
                                        // วาดเส้นตาราง โดยใช้เซลล์ A3 ถึง P3 
                                        DrawTableBorders(worksheet, 14, 1, startRows - 1, (sce_show == true ? 18 : 16));

                                        worksheet.Cells["A" + (startRows)].Value = (dr[0]["descriptions_worksheet"] + "");

                                        //Delete Column major_accident_event
                                        worksheet.DeleteColumn(9);

                                        if (show_cat == "0")
                                        {
                                            worksheet.DeleteColumn(5);
                                        }

                                        worksheet.Cells["A" + startRows_Def + ":O" + startRows].Style.WrapText = true;
                                    }
                                    try
                                    {
                                        //new worksheet move after WorksheetTemplate 
                                        excelPackage.Workbook.Worksheets.MoveBefore(worksheet_name, worksheet_name_target);
                                    }
                                    catch { }
                                }
                            }
                        }

                        ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                        SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                        excelPackage.Save();
                    }
                }
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }
        public string excel_whatif_recommendation(string seq, string file_fullpath_name, Boolean report_all, string seq_worksheet_def, string action_owner_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            try
            {
                #region Get Data
                List<SqlParameter> parameters = new List<SqlParameter>();

                sqlstr = @" select distinct
                        h.seq, h.pha_no, nl.id as id_list, g.pha_request_name
                        , nl.list, nl.list as list_check, nl.design_intent, nl.descriptions, nl.design_conditions, nl.list_boundary, nl.operating_conditions
                        , d.document_no, d.document_file_name
                        , nw.list_system, nw.list_sub_system, nw.causes, nw.consequences
                        , nw.category_type, nw.ram_befor_security, nw.ram_befor_likelihood, nw.ram_befor_risk
                        , nw.existing_safeguards, nw.recommendations, nw.recommendations_no, nw.responder_user_name, nw.responder_user_displayname
                        , nw.action_status
                        , nl.no as list_no, nw.no, nw.list_system_no, nw.list_sub_system_no, nw.causes_no, nw.consequences_no
                        , nw.seq as seq_worksheet
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_LIST nl on h.id = nl.id_pha 
                        left join EPHA_T_LIST_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_list 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_LIST_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_list
                        where h.seq = @seq and nw.responder_user_name is not null ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

                if (!string.IsNullOrEmpty(action_owner_name))
                {
                    sqlstr += @" and lower(nw.responder_user_name) = lower(@action_owner_name) ";
                    parameters.Add(new SqlParameter("@action_owner_name", SqlDbType.VarChar, 400) { Value = action_owner_name ?? "" });
                }

                sqlstr += @" order by cast(nl.no as int),cast(nw.list_system_no as int),cast(nw.list_sub_system_no as int),cast(nw.no as int), 
                        cast(nw.causes_no as int), cast(nw.consequences_no as int)";


                DataTable dtWorksheet = new DataTable();
                //dtWorksheet = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorksheet = new DataTable();
                        dtWorksheet = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorksheet.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable
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
                        dtWorksheet = new DataTable();
                        dtWorksheet = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorksheet.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                sqlstr = @" select distinct nl.no, nw.no, nw.seq, 0 as ref, nl.list, nl.list as list_check
                        , nw.ram_after_risk, nw.ram_after_risk_action, nw.recommendations, nw.recommendations_no, nw.action_status, nw.responder_user_name, nw.responder_user_displayname 
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_LIST nl on h.id = nl.id_pha 
                        left join EPHA_T_LIST_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_list 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_LIST_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_list    
                        where h.seq = @seq and nw.responder_user_name is not null ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

                if (!string.IsNullOrEmpty(action_owner_name))
                {
                    sqlstr += @" and lower(nw.responder_user_name) = lower(@action_owner_name) ";
                    parameters.Add(new SqlParameter("@action_owner_name", SqlDbType.VarChar, 400) { Value = action_owner_name ?? "" });
                }

                sqlstr += @" order by nl.no, nw.no ";

                DataTable dtTrack = new DataTable();
                //dtTrack = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtTrack = new DataTable();
                        dtTrack = _conn.ExecuteAdapter(command).Tables[0];
                        dtTrack.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable



                if (true)
                {
                    for (int t = 0; t < dtTrack?.Rows.Count; t++)
                    {
                        dtTrack.Rows[t]["ref"] = (t + 1);
                        dtTrack.AcceptChanges();
                    }
                }
                #endregion Get Data

                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "whatif";//sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                    {
                        if (!report_all)
                        {
                            var sheetsToDelete = excelPackage.Workbook.Worksheets
                              .Where(sheet => sheet.Name != "RecommTemplate" && sheet.Name != "TrackTemplate")
                              .ToList(); // ใช้ ToList เพื่อหลีกเลี่ยงปัญหา Collection Modified

                            foreach (var sheet in sheetsToDelete)
                            {
                                excelPackage.Workbook.Worksheets.Delete(sheet);
                            }
                        }

                        DataTable dt = new DataTable(); dt = dtWorksheet.Copy(); dt.AcceptChanges();
                        for (int i = 0; i < dt?.Rows.Count; i++)
                        {
                            #region Sheet
                            if (true)
                            {
                                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["RecommTemplate"];
                                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("RecommTemplate" + i, sourceWorksheet);

                                string ref_no = (i + 1).ToString();
                                worksheet.Name = "Response Sheet(Ref." + ref_no + ")";

                                string responder_user_name = (dt.Rows[i]["responder_user_name"] + "");
                                string responder_user_displayname = (dt.Rows[i]["responder_user_displayname"] + "");
                                string pha_request_name = (dt.Rows[i]["pha_request_name"] + "");
                                string pha_no = (dt.Rows[i]["pha_no"] + "");
                                string seq_worksheet = (dt.Rows[i]["seq_worksheet"] + "");


                                int startRows = 2;
                                if (true)
                                {
                                    string list = "";
                                    string drawing_doc = "";
                                    string list_system = "";
                                    string list_sub_system = "";
                                    string causes = "";
                                    string consequences = "";
                                    string existing_safeguards = "";
                                    string recommendations = "";
                                    string recommendations_no = "";
                                    int action_no = 0;

                                    #region loop drawing_doc 
                                    drawing_doc = (dt.Rows[i]["document_no"] + "");
                                    if ((dt.Rows[i]["document_file_name"] + "") != "")
                                    {
                                        drawing_doc += " (" + dt.Rows[i]["document_file_name"] + ")";
                                    }
                                    #endregion loop drawing_doc 

                                    #region loop workksheet
                                    //DataRow[] drWorksheet = dt.Select("seq_worksheet = '" + seq_worksheet + "'");
                                    var filterParameters = new Dictionary<string, object>();
                                    filterParameters.Add("seq_worksheet", seq_worksheet);
                                    var (drWorksheet, iMerge) = FilterDataTable(dt, filterParameters);
                                    if (drWorksheet != null)
                                    {
                                        if (drWorksheet?.Length > 0)
                                        {
                                            for (int n = 0; n < drWorksheet.Length; n++)
                                            {
                                                if ((drWorksheet[n]["list_system"] + "") != "")
                                                {
                                                    if (list_system != "") { list_system += ","; }
                                                    list_system += (drWorksheet[n]["list_system"] + "");
                                                }
                                                if ((drWorksheet[n]["list_sub_system"] + "") != "")
                                                {
                                                    if (list_sub_system != "") { list_sub_system += ","; }
                                                    list_sub_system += (drWorksheet[n]["list_sub_system"] + "");
                                                }
                                                if ((drWorksheet[n]["causes"] + "") != "")
                                                {
                                                    if (causes != "") { causes += ","; }
                                                    causes += (drWorksheet[n]["causes"] + "");
                                                }
                                                if ((drWorksheet[n]["consequences"] + "") != "")
                                                {
                                                    if (consequences != "") { consequences += ","; }
                                                    consequences += (drWorksheet[n]["consequences"] + "");
                                                }

                                                if ((drWorksheet[n]["existing_safeguards"] + "") != "")
                                                {
                                                    if (existing_safeguards.IndexOf((drWorksheet[n]["existing_safeguards"] + "")) > -1) { }
                                                    else
                                                    {
                                                        if (existing_safeguards != "") { existing_safeguards += ","; }
                                                        existing_safeguards += (drWorksheet[n]["existing_safeguards"] + "");
                                                    }
                                                }

                                                if ((drWorksheet[n]["recommendations"] + "") != "")
                                                {
                                                    if (recommendations != "") { recommendations += ","; }
                                                    recommendations += (drWorksheet[n]["recommendations"] + "");
                                                    action_no += 1;

                                                    if (recommendations_no != "") { recommendations_no += ","; }
                                                    recommendations_no += (drWorksheet[n]["recommendations_no"] + "");
                                                }

                                            }
                                        }
                                    }
                                    #endregion loop workksheet

                                    worksheet.Cells["A" + (startRows)].Value = "Project Title:" + pha_request_name;
                                    startRows += 1;
                                    worksheet.Cells["A" + (startRows)].Value = "Project No:" + pha_no;
                                    startRows += 1;
                                    worksheet.Cells["A" + (startRows)].Value = "List:" + list;
                                    startRows += 1;

                                    worksheet.Cells["B" + (startRows)].Value = responder_user_displayname;
                                    worksheet.Cells["E" + (startRows)].Value = responder_user_displayname;
                                    startRows += 1;

                                    worksheet.Cells["B" + (startRows)].Value = action_no;
                                    startRows += 1;

                                    worksheet.Cells["B" + (startRows)].Value = drawing_doc;
                                    startRows += 1;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = list_system;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = list_sub_system;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = causes;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = consequences;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = existing_safeguards;
                                    startRows += 1;
                                    worksheet.Cells["B" + (startRows)].Value = recommendations;
                                    startRows += 1;
                                }

                            }
                            #endregion Sheet

                        }

                        #region TrackTemplate
                        if (dtTrack?.Rows.Count > 0)
                        {
                            //ข้อมูลทั้งหมด
                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TrackTemplate"];
                            worksheet.Name = "Status Tracking Table";

                            int i = 0;
                            int startRows = 3;

                            dt = new DataTable(); dt = dtTrack.Copy(); dt.AcceptChanges();
                            if (dt?.Rows.Count > 0)
                            {
                                for (i = 0; i < dt?.Rows.Count; i++)
                                {
                                    worksheet.InsertRow(startRows, 1);
                                    worksheet.Cells["A" + (startRows)].Value = dt.Rows[i]["ref"].ToString();
                                    worksheet.Cells["B" + (startRows)].Value = dt.Rows[i]["list"].ToString();
                                    worksheet.Cells["C" + (startRows)].Value = dt.Rows[i]["ram_after_risk"].ToString();
                                    worksheet.Cells["D" + (startRows)].Value = dt.Rows[i]["recommendations"].ToString();
                                    worksheet.Cells["E" + (startRows)].Value = dt.Rows[i]["action_status"].ToString();
                                    worksheet.Cells["F" + (startRows)].Value = dt.Rows[i]["responder_user_displayname"].ToString();
                                    startRows++;
                                }

                                // วาดเส้นตาราง โดยใช้เซลล์ A1 ถึง C3
                                DrawTableBorders(worksheet, 3, 1, startRows - 1, 6);
                            }
                        }
                        #endregion Response Sheet


                        ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["RecommTemplate"];
                        SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                        excelPackage.Save();

                    }
                }

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }
        public string excle_template_data_whatif(string seq, string file_fullpath_name, Boolean report_all)
        {

            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            try
            {
                List<SqlParameter> parameters = new List<SqlParameter>();

                sqlstr = @" select distinct h.pha_no, g.pha_request_name, format(g.target_start_date,'dd MMM yyyy') as target_start_date
                        from epha_t_header h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        where h.seq = @seq ";
                sqlstr += @" order by g.pha_request_name";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

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



                if (!string.IsNullOrEmpty(file_fullpath_name))
                {
                    string file_fullpath_def = file_fullpath_name;
                    string folder = "whatif";//sub_software;


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

                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                {
                    FileInfo template_excel = new FileInfo(file_fullpath_name);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(template_excel))

                    {
                        if (!report_all)
                        {
                            var sheetsToDelete = excelPackage.Workbook.Worksheets
                              .Where(sheet => sheet.Name != "WorksheetTemplate")
                              .ToList(); // ใช้ ToList เพื่อหลีกเลี่ยงปัญหา Collection Modified

                            foreach (var sheet in sheetsToDelete)
                            {
                                excelPackage.Workbook.Worksheets.Delete(sheet);
                            }
                        }

                        ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                        ExcelWorksheet worksheet = sourceWorksheet;
                        try
                        {
                            if (dt?.Rows.Count > 0)
                            {
                                // c4
                                worksheet.Cells["C4"].Value = dt.Rows[0]["pha_request_name"].ToString();
                                // i4 = วันที่ทำการประเมิน (Date): 23/5/2566
                                worksheet.Cells["I4"].Value = "วันที่ทำการประเมิน (Date):" + dt.Rows[0]["target_start_date"].ToString();
                            }
                        }
                        catch { }
                        try
                        {
                            var startRows = 12;
                            var icol_end = 14;
                            for (int i = 0; i < 10; i++)
                            {
                                worksheet.InsertRow(startRows, 1);
                            }
                            DrawTableBorders(worksheet, startRows - 1, 2, startRows - 1, icol_end - 1);
                        }
                        catch { }
                        excelPackage.Save();
                    }
                }
            Next_Line:;
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }

        #endregion export excel what'if

        #region export excel hra

        // private string hra_query_workscope()
        // {
        //     cls = new ClassFunctions();
        //     sqlstr = @"  select h.seq, g.work_scope, h.pha_no, h.pha_version_text as pha_version, g.descriptions, format(g.target_start_date, 'dd MMM yyyy ') as target_start_date 
        //                  , case when ums.user_name is null  then h.request_user_displayname else case when ums.departments is null  then  ums.user_displayname else  ums.user_displayname + ' (' + ums.departments +')' end end request_user_displayname
        //                  , g.id_departments +' '+ g.id_sections as company, apu.descriptions as name_of_area,format(l.create_date, 'dd MMM yyyy ')  as assessment_date 
        //                  , '' as assessment_team_leader, ums_ta1.user_displayname as qmts_reviewer 
        //                  from epha_t_header h 
        //                  inner join EPHA_T_GENERAL g on h.id = g.id_pha 
        //left join EPHA_M_COMPANY com on g.id_toc = com.seq 
        //left join EPHA_M_BUSINESS_UNIT apu on g.id_unit_no = apu.id and apu.id_company = com.seq
        //left join EPHA_T_APPROVER ta1 on h.id = ta1.id_pha and ta1.approver_type = 'section_head'  
        //                  left join VW_EPHA_PERSON_DETAILS ums on lower(h.request_user_name) = lower(ums.user_name) 
        //                  left join VW_EPHA_PERSON_DETAILS ums_ta1 on lower(ta1.user_name) = lower(ums_ta1.user_name) 
        //                  left join ( select max(seq) as seq, id_pha, pha_status, create_date from epha_t_action_log group by id_pha, pha_status, create_date ) l on h.id = l.id_pha and l.pha_status = 21
        //                 ";
        //     return sqlstr;
        // }
        //private string hra_query_summary_report()
        //{
        //    cls = new ClassFunctions();
        //    //b.xxx =>> dummy field, replace after ever query
        //    sqlstr = @" select distinct b.responder_user_name, b.xxx from epha_t_header a  
        //                inner join EPHA_T_TABLE1_HAZARD hz on a.id  = hz.id_pha 
        //                inner join EPHA_T_TABLE2_TASKS ts on a.id  = ts.id_pha 
        //                inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha and hz.id = b.id_hazard and ts.id = b.id_tasks  
        //                left join EPHA_M_ACTIVITIES ma on b.id_activity = ma.id
        //                left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name)
        //                where a.seq = @seq ";
        //    sqlstr += " order by b.xxx";
        //    return sqlstr; 
        //}
        public string excel_hra_general(string seq, string file_fullpath_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";
            try
            {
                List<SqlParameter> parameters = new List<SqlParameter>();

                #region get data
                cls = new ClassFunctions();
                //string query_sqlstr = hra_query_workscope();
                sqlstr = "select t.* from VW_EPHA_DATA_HRA_WORKSCOPE t where t.seq = @seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtWorkScope = new DataTable();
                //dtWorkScope = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorkScope = new DataTable();
                        dtWorkScope = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorkScope.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = @"select distinct user_displayname from  EPHA_T_APPROVER where approver_type = 'approver' and id_pha = @seq order by user_displayname ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtApprover = new DataTable();
                //dtApprover = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtApprover = new DataTable();
                        dtApprover = _conn.ExecuteAdapter(command).Tables[0];
                        dtApprover.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (dtWorkScope?.Rows.Count > 0 && dtApprover?.Rows.Count > 0)
                {
                    for (int i = 0; i < dtWorkScope?.Rows.Count; i++)
                    {
                        if (i > 0) { dtWorkScope.Rows[0]["assessment_team_leader"] += ","; }
                        dtWorkScope.Rows[0]["assessment_team_leader"] += dtApprover.Rows[i]["user_displayname"]?.ToString();
                    }
                }
                #endregion get data

                if (dtWorkScope != null)
                {
                    if (dtWorkScope?.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(file_fullpath_name))
                        {
                            string file_fullpath_def = file_fullpath_name;
                            string folder = "hra";//sub_software;


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

                        if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                        {
                            FileInfo template_excel = new FileInfo(file_fullpath_name);
                            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                            using (ExcelPackage excelPackage = new ExcelPackage(template_excel))

                            {
                                //HAZOP Cover Page
                                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["HRA Cover Page"];  // Replace "SourceSheet" with the actual source sheet name
                                ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                                worksheet.Cells["B6"].Value = "Company: " + (dtWorkScope.Rows[0]["company"] + "");
                                worksheet.Cells["B7"].Value = "Name of Area: " + (dtWorkScope.Rows[0]["name_of_area"] + "");
                                worksheet.Cells["B8"].Value = "Assessment For: " + (dtWorkScope.Rows[0]["assessment_team_leader"] + "");
                                worksheet.Cells["B9"].Value = "Assessment Date: " + (dtWorkScope.Rows[0]["assessment_date"] + "");

                                excelPackage.Save();
                            }
                        }
                    }
                }

            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            return msg_error;
        }
        public string excel_hra_worksheet(string seq, string file_fullpath_name, Boolean report_all, Boolean part_recommendations)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }
            if (string.IsNullOrEmpty(file_fullpath_name)) { return "Invalid File Fullpath Name."; }

            string msg_error = "";

            try
            {
                List<SqlParameter> parameters = new List<SqlParameter>();

                #region get data
                cls = new ClassFunctions();
                //string query_sqlstr = hra_query_workscope();
                sqlstr = "select t.* from VW_EPHA_DATA_HRA_WORKSCOPE t where t.seq = @seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtWorkScope = new DataTable();
                //dtWorkScope = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorkScope = new DataTable();
                        dtWorkScope = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorkScope.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change, 0 as index_rows
                        from epha_t_header a inner join EPHA_T_TABLE1_SUBAREAS b on a.id  = b.id_pha
                        where a.seq = @seq  ";
                sqlstr += " order by a.seq,b.seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtSubareas = new DataTable();
                //dtSubareas = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtSubareas = new DataTable();
                        dtSubareas = _conn.ExecuteAdapter(command).Tables[0];
                        dtSubareas.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change, 0 as index_rows
                        from epha_t_header a inner join EPHA_T_TABLE1_SUBAREAS b on a.id  = b.id_pha
                        where a.seq = @seq  ";
                sqlstr += " order by a.seq,b.seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtHazard = new DataTable();
                //dtHazard = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtHazard = new DataTable();
                        dtHazard = _conn.ExecuteAdapter(command).Tables[0];
                        dtHazard.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable



                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change, 0 as index_rows
                        from epha_t_header a inner join EPHA_T_TABLE2_TASKS b on a.id  = b.id_pha
                        where a.seq = @seq  ";
                sqlstr += " order by a.seq,b.seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtTasks = new DataTable();
                //dtTasks = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtTasks = new DataTable();
                        dtTasks = _conn.ExecuteAdapter(command).Tables[0];
                        dtTasks.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                sqlstr = @" select b.* , 'update' as action_type, 0 as action_change, 0 as index_rows
                        from epha_t_header a inner join EPHA_T_TABLE2_WORKERS b on a.id  = b.id_pha
                        where a.seq = @seq  ";
                sqlstr += " order by a.seq,b.seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtWorker = new DataTable();
                //dtWorker = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorker = new DataTable();
                        dtWorker = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorker.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                #region worksheet

                sqlstr = @" select b.* , 0 as no   
                        , b.id_hazard as seq_hazard, b.id_tasks as seq_tasks
                        , vw.user_id as responder_user_id, vw.user_email as responder_user_email
                        , 'assets/img/team/avatar.webp' as responder_user_img
                        , hz.no as hazard_no, ts.no as tasks_no
                        , 0 as index_rows
                        , format(b.estimated_end_date, 'dd MMM yyyy') as due_date 
                        , ts.worker_group, ma.name  as activity, hz.type_hazard, hz.health_effect_rating
                        , b.standard_value + '' + b.standard_unit + '' + b.standard_desc as values_standard
                        from epha_t_header a  
                        inner join EPHA_T_TABLE1_HAZARD hz on a.id  = hz.id_pha 
                        inner join EPHA_T_TABLE2_TASKS ts on a.id  = ts.id_pha 
                        inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha and hz.id = b.id_hazard and ts.id = b.id_tasks  
                        left join EPHA_M_ACTIVITIES ma on b.id_activity = ma.id
                        left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name) 
                        where a.seq = @seq  ";
                sqlstr += " order by hz.no, ts.no, b.no";


                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtWorksheet = new DataTable();
                //dtWorksheet = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtWorksheet = new DataTable();
                        dtWorksheet = _conn.ExecuteAdapter(command).Tables[0];
                        dtWorksheet.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                sqlstr = @" select c.id_worksheet, c.recommendations   
                        from epha_t_header a  
                        inner join EPHA_T_TABLE1_HAZARD hz on a.id  = hz.id_pha 
                        inner join EPHA_T_TABLE2_TASKS ts on a.id  = ts.id_pha 
                        inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha and hz.id = b.id_hazard and ts.id = b.id_tasks  
                        inner join EPHA_T_RECOMMENDATIONS c on a.id = c.id_pha and b.id = c.id_worksheet 
                        where a.seq = @seq  ";
                sqlstr += " order by hz.no, ts.no, b.no";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtRecom = new DataTable();
                //dtRecom = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtRecom = new DataTable();
                        dtRecom = _conn.ExecuteAdapter(command).Tables[0];
                        dtRecom.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (dtRecom?.Rows.Count > 0)
                {
                    for (int i = 0; i < dtWorksheet?.Rows.Count; i++)
                    {
                        string id_worksheet = dtWorksheet.Rows[i]["id"]?.ToString() ?? "";
                        string recommendations = "";
                        //DataRow[] drSelect = dtRecom.Select("id_worksheet='" + id_worksheet + "'");
                        var filterParameters = new Dictionary<string, object>();
                        filterParameters.Add("id_worksheet", id_worksheet);
                        var (drSelect, iMerge) = FilterDataTable(dtRecom, filterParameters);
                        if (drSelect != null)
                        {
                            if (drSelect?.Length > 0)
                            {
                                for (int j = 0; j < drSelect.Length; j++)
                                {
                                    if (recommendations != "") { recommendations += ","; }
                                    recommendations += drSelect[j]["recommendations"]?.ToString() ?? "";
                                }
                            }
                        }
                        dtWorksheet.Rows[i]["recommendations"] = recommendations;
                    }
                }
                #endregion worksheet


                #region Summary Report
                //dtIndicator,dtIrr,dtHoc,dtRrr,dtRec
                //responder_user_displayname,initial_risk_rating,hierarchy_of_control,residual_risk_rating,recommendations

                //query_sqlstr = hra_query_summary_report();

                ////hra_query_summary_report(seq, "responder_user_displayname");
                //sqlstr = query_sqlstr.Replace("xxx", @"responder_user_displayname");
                sqlstr = @"  select distinct b.responder_user_name, b.responder_user_displayname
                             from epha_t_header a  
                                              inner join EPHA_T_TABLE1_HAZARD hz on a.id  = hz.id_pha 
                                              inner join EPHA_T_TABLE2_TASKS ts on a.id  = ts.id_pha 
                                              inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha and hz.id = b.id_hazard and ts.id = b.id_tasks  
                                              left join EPHA_M_ACTIVITIES ma on b.id_activity = ma.id
                                              left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name)
                             where a.seq = @seq  
                             order by b.responder_user_displayname ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtIndicator = new DataTable();
                //dtIndicator = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtIndicator = new DataTable();
                        dtIndicator = _conn.ExecuteAdapter(command).Tables[0];
                        dtIndicator.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                ////sqlstr = hra_query_summary_report(seq, "initial_risk_rating");
                //sqlstr = query_sqlstr.Replace("xxx", @"initial_risk_rating"); 
                sqlstr = @"  select distinct b.responder_user_name, b.initial_risk_rating
                             from epha_t_header a  
                                              inner join EPHA_T_TABLE1_HAZARD hz on a.id  = hz.id_pha 
                                              inner join EPHA_T_TABLE2_TASKS ts on a.id  = ts.id_pha 
                                              inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha and hz.id = b.id_hazard and ts.id = b.id_tasks  
                                              left join EPHA_M_ACTIVITIES ma on b.id_activity = ma.id
                                              left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name)
                             where a.seq = @seq  
                             order by b.initial_risk_rating ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtIrr = new DataTable();
                //dtIrr = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtIrr = new DataTable();
                        dtIrr = _conn.ExecuteAdapter(command).Tables[0];
                        dtIrr.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                ////sqlstr = hra_query_summary_report(seq, "hierarchy_of_control");
                //sqlstr = query_sqlstr.Replace("xxx", @"hierarchy_of_control");
                sqlstr = @"  select distinct b.responder_user_name, b.hierarchy_of_control
                             from epha_t_header a  
                                              inner join EPHA_T_TABLE1_HAZARD hz on a.id  = hz.id_pha 
                                              inner join EPHA_T_TABLE2_TASKS ts on a.id  = ts.id_pha 
                                              inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha and hz.id = b.id_hazard and ts.id = b.id_tasks  
                                              left join EPHA_M_ACTIVITIES ma on b.id_activity = ma.id
                                              left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name)
                             where a.seq = @seq  
                             order by b.hierarchy_of_control ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtHoc = new DataTable();
                //dtHoc = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtHoc = new DataTable();
                        dtHoc = _conn.ExecuteAdapter(command).Tables[0];
                        dtHoc.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                ////sqlstr = hra_query_summary_report(seq, "residual_risk_rating");
                //sqlstr = query_sqlstr.Replace("xxx", @"residual_risk_rating");
                sqlstr = @"  select distinct b.responder_user_name, b.residual_risk_rating
                             from epha_t_header a  
                                              inner join EPHA_T_TABLE1_HAZARD hz on a.id  = hz.id_pha 
                                              inner join EPHA_T_TABLE2_TASKS ts on a.id  = ts.id_pha 
                                              inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha and hz.id = b.id_hazard and ts.id = b.id_tasks  
                                              left join EPHA_M_ACTIVITIES ma on b.id_activity = ma.id
                                              left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name)
                             where a.seq = @seq  
                             order by b.residual_risk_rating ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtRrr = new DataTable();
                //dtRrr = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtRrr = new DataTable();
                        dtRrr = _conn.ExecuteAdapter(command).Tables[0];
                        dtRrr.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                //sqlstr = hra_query_summary_report(seq, "recommendations");
                sqlstr = @" select distinct b.responder_user_name, c.recommendations 
                        from epha_t_header a  
                        inner join EPHA_T_TABLE1_HAZARD hz on a.id  = hz.id_pha 
                        inner join EPHA_T_TABLE2_TASKS ts on a.id  = ts.id_pha 
                        inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha
                        inner join EPHA_T_RECOMMENDATIONS c on a.id = c.id_pha and b.id = c.id_worksheet
                        where a.seq = @seq   
                        order by b.responder_user_name";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtRec = new DataTable();
                //dtRec = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtRec = new DataTable();
                        dtRec = _conn.ExecuteAdapter(command).Tables[0];
                        dtRec.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                #endregion Summary Report

                #region mapping data 
                DataTable dtSummaryReport = dtWorksheet.Clone(); dtSummaryReport.AcceptChanges();
                for (int i = 0; i < dtIndicator?.Rows.Count; i++)
                {
                    //dtIrr,dtHoc,dtRrr,dtRec
                    //initial_risk_rating,hierarchy_of_control,residual_risk_rating,recommendations
                    string responder_user_name = dtIndicator.Rows[i]["responder_user_name"]?.ToString() ?? "";
                    string responder_user_displayname = dtIndicator.Rows[i]["responder_user_displayname"]?.ToString() ?? "";
                    string initial_risk_rating = "";
                    string hierarchy_of_control = "";
                    string residual_risk_rating = "";
                    string recommendations = "";

                    //DataRow[] drSelect = dtIrr.Select("responder_user_name='" + responder_user_name + "'");
                    var filterParameters = new Dictionary<string, object>();
                    filterParameters.Add("responder_user_name", responder_user_name);
                    var (drSelect, iMerge) = FilterDataTable(dtIrr, filterParameters);
                    if (drSelect != null)
                    {
                        if (drSelect?.Length > 0)
                        {
                            for (int j = 0; j < drSelect.Length; j++)
                            {
                                if (initial_risk_rating != "") { initial_risk_rating += ","; }
                                initial_risk_rating += drSelect[j]["initial_risk_rating"]?.ToString() ?? "";
                            }
                        }
                    }

                    //drSelect = dtHoc.Select("responder_user_name='" + responder_user_name + "'");
                    filterParameters = new Dictionary<string, object>();
                    filterParameters.Add("responder_user_name", responder_user_name);
                    (drSelect, iMerge) = FilterDataTable(dtHoc, filterParameters);
                    if (drSelect != null)
                    {
                        if (drSelect?.Length > 0)
                        {
                            for (int j = 0; j < drSelect.Length; j++)
                            {
                                if (hierarchy_of_control != "") { hierarchy_of_control += ","; }
                                hierarchy_of_control += drSelect[j]["hierarchy_of_control"]?.ToString() ?? "";

                            }
                        }
                    }

                    //drSelect = dtRrr.Select("responder_user_name='" + responder_user_name + "'");
                    filterParameters = new Dictionary<string, object>();
                    filterParameters.Add("responder_user_name", responder_user_name);
                    (drSelect, iMerge) = FilterDataTable(dtRrr, filterParameters);
                    if (drSelect != null)
                    {
                        if (drSelect?.Length > 0)
                        {
                            for (int j = 0; j < drSelect.Length; j++)
                            {
                                if (residual_risk_rating != "") { residual_risk_rating += ","; }
                                residual_risk_rating += drSelect[j]["residual_risk_rating"]?.ToString() ?? "";
                            }
                        }
                    }

                    //drSelect = dtRec.Select("responder_user_name='" + responder_user_name + "'");
                    filterParameters = new Dictionary<string, object>();
                    filterParameters.Add("responder_user_name", responder_user_name);
                    (drSelect, iMerge) = FilterDataTable(dtRec, filterParameters);
                    if (drSelect != null)
                    {
                        if (drSelect?.Length > 0)
                        {
                            for (int j = 0; j < drSelect.Length; j++)
                            {
                                if (recommendations != "") { recommendations += ","; }
                                recommendations += drSelect[j]["recommendations"]?.ToString() ?? "";
                            }
                        }
                    }

                    dtSummaryReport.Rows.Add(dtSummaryReport.NewRow());
                    dtSummaryReport.Rows[i]["responder_user_name"] = responder_user_name;
                    dtSummaryReport.Rows[i]["responder_user_displayname"] = responder_user_displayname;
                    dtSummaryReport.Rows[i]["initial_risk_rating"] = initial_risk_rating;
                    dtSummaryReport.Rows[i]["hierarchy_of_control"] = hierarchy_of_control;
                    dtSummaryReport.Rows[i]["residual_risk_rating"] = residual_risk_rating;
                    dtSummaryReport.Rows[i]["recommendations"] = recommendations;
                    dtSummaryReport.AcceptChanges();
                }
                #endregion mapping data 


                sqlstr = @" select distinct b.responder_user_name, b.responder_user_displayname 
                         , isnull(c.recommendations,'') as recommendations
                         , isnull(format(b.estimated_end_date, 'dd MMM yyyy'),'') as due_date
                         , case when isnull(b.effective,0) = 0 then 'Effective' else 'Ineffective' end action_status  
                         from epha_t_header a  
                         inner join EPHA_T_TABLE1_HAZARD hz on a.id  = hz.id_pha 
                         inner join EPHA_T_TABLE2_TASKS ts on a.id  = ts.id_pha 
                         inner join EPHA_T_TABLE3_WORKSHEET b on a.id  = b.id_pha and hz.id = b.id_hazard and ts.id = b.id_tasks  
                         left join EPHA_T_RECOMMENDATIONS c on a.id = c.id_pha and b.id = c.id_worksheet
                         left join EPHA_M_ACTIVITIES ma on b.id_activity = ma.id
                         left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name) 
                         where a.seq = @seq   
                         order by b.responder_user_name";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtActionResp = new DataTable();
                //dtActionResp = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtActionResp = new DataTable();
                        dtActionResp = _conn.ExecuteAdapter(command).Tables[0];
                        dtActionResp.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable


                #region Summary Risk 

                sqlstr = $" select t.* from VW_EPHA_DATA_HRA_SUMMARYRISK t where t.seq = @seq";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtSummaryRisk = new DataTable();
                //dtSummaryRisk = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtSummaryRisk = new DataTable();
                        dtSummaryRisk = _conn.ExecuteAdapter(command).Tables[0];
                        dtSummaryRisk.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable



                #endregion Summary Risk


                #endregion get data


                if (dtWorkScope != null)
                {
                    if (dtWorkScope?.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(file_fullpath_name))
                        {
                            string file_fullpath_def = file_fullpath_name;
                            string folder = "whatif";//sub_software;


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

                        if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                        {
                            FileInfo template_excel = new FileInfo(file_fullpath_name);
                            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                            using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                            {
                                //HAZOP Cover Page
                                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["HRA Cover Page"];  // Replace "SourceSheet" with the actual source sheet name
                                ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                                string company = (dtWorkScope.Rows[0]["company"] + "");
                                string name_of_area = (dtWorkScope.Rows[0]["name_of_area"] + "");
                                string assessment_date = (dtWorkScope.Rows[0]["assessment_date"] + "");
                                string assessment_team_leader = (dtWorkScope.Rows[0]["assessment_team_leader"] + "");
                                string qmts_reviewer = (dtWorkScope.Rows[0]["qmts_reviewer"] + "");

                                #region Summary Risk
                                if (true)
                                {
                                    worksheet = excelPackage.Workbook.Worksheets["Summary Risk"];
                                    if (worksheet != null)
                                    {
                                        int startRows = 6;
                                        int startRows_Def = startRows;

                                        if (dtSummaryRisk != null)
                                        {
                                            for (int i = 0; i < dtSummaryRisk?.Rows.Count; i++)
                                            {
                                                worksheet.Cells["G" + (startRows)].Value = dtSummaryRisk.Rows[i]["initial_risk_rating"] ?? "";
                                                worksheet.Cells["P" + (startRows)].Value = dtSummaryRisk.Rows[i]["residual_risk_rating"] ?? "";
                                                startRows++;
                                            }
                                        }
                                    }
                                }
                                #endregion Summary Risk


                                #region Summary Report
                                if (true)
                                {
                                    worksheet = excelPackage.Workbook.Worksheets["Summary Report"];

                                    worksheet.Cells["B3"].Value = "Company: " + company;
                                    worksheet.Cells["B4"].Value = "Name of Area: " + name_of_area;
                                    worksheet.Cells["B5"].Value = "Assessment Team Leader: " + assessment_team_leader;

                                    int startRows = 6;
                                    int icol_end = 4;
                                    //int startRows_Def = 6;

                                    //responder_user_displayname, b.initial_risk_rating, b.hierarchy_of_control, b.residual_risk_rating, b.recommendations
                                    for (int i = 0; i < dtSummaryReport?.Rows.Count; i++)
                                    {
                                        if (i > 0)
                                        {
                                            worksheet.InsertRow(startRows, 1, 9);
                                            worksheet.InsertRow(startRows, 1, 8);
                                            worksheet.InsertRow(startRows, 1, 7);
                                            worksheet.InsertRow(startRows, 1, 6);
                                        }

                                        //No.	Document Name	Drawing No	Document File	Comment
                                        //worksheet.InsertRow(startRows_Def, 1);
                                        worksheet.Cells["A" + (startRows)].Value = "Exposure Indicator " + (i + 1) + ":";
                                        worksheet.Cells["B" + (startRows)].Value = (dtSummaryReport.Rows[i]["responder_user_displayname"] + "");
                                        worksheet.Cells["C" + (startRows)].Value = ("Risk Ranking:");
                                        worksheet.Cells["D" + (startRows)].Value = (dtSummaryReport.Rows[i]["initial_risk_rating"] + "");
                                        startRows++;
                                        worksheet.Cells["B" + (startRows - 1)].Style.WrapText = true;

                                        worksheet.Cells["A" + (startRows)].Value = "Hierarchy of Control:";
                                        MergeCell(worksheet, "B" + (startRows), "D" + (startRows));
                                        worksheet.Cells["B" + (startRows)].Value = (dtSummaryReport.Rows[i]["hierarchy_of_control"] + "");
                                        startRows++;
                                        worksheet.Cells["B" + (startRows - 1)].Style.WrapText = true;

                                        worksheet.Cells["A" + (startRows)].Value = "Residual Risk Ranking:";
                                        MergeCell(worksheet, "B" + (startRows), "D" + (startRows));
                                        worksheet.Cells["B" + (startRows)].Value = (dtSummaryReport.Rows[i]["residual_risk_rating"] + "");
                                        startRows++;
                                        worksheet.Cells["B" + (startRows - 1)].Style.WrapText = true;

                                        worksheet.Cells["A" + (startRows)].Value = "Recommendations:";
                                        MergeCell(worksheet, "B" + (startRows), "D" + (startRows));
                                        worksheet.Cells["B" + (startRows)].Value = (dtSummaryReport.Rows[i]["recommendations"] + "");
                                        startRows++;
                                        worksheet.Cells["B" + (startRows - 1)].Style.WrapText = true;

                                    }
                                    // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                                    DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);
                                }
                                #endregion Summary Report

                                #region Action Response
                                if (true)
                                {
                                    worksheet = excelPackage.Workbook.Worksheets["Action Response"];

                                    int startRows = 6;
                                    int icol_end = 4;
                                    int startRows_Def = startRows;

                                    for (int i = 0; i < dtActionResp?.Rows.Count; i++)
                                    {
                                        //No.	Document Name	Drawing No	Document File	Comment
                                        //worksheet.InsertRow(startRows_Def + 1, 1);

                                        if ((dtActionResp.Rows[i]["action_status"] + "").ToLower() == "open")
                                        {
                                            worksheet.InsertRow(startRows, 1, 4);
                                        }
                                        else { worksheet.InsertRow(startRows, 1, 5); }

                                        worksheet.Cells["A" + (startRows)].Value = (dtActionResp.Rows[i]["recommendations"] + "");
                                        worksheet.Cells["B" + (startRows)].Value = (dtActionResp.Rows[i]["responder_user_displayname"] + "");
                                        worksheet.Cells["C" + (startRows)].Value = (dtActionResp.Rows[i]["due_date"] + "");
                                        worksheet.Cells["D" + (startRows)].Value = (dtActionResp.Rows[i]["action_status"] + "");
                                        startRows++;

                                    }
                                    // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                                    DrawTableBorders(worksheet, 1, 1, startRows, icol_end);

                                    worksheet.DeleteRow(5);
                                    worksheet.DeleteRow(4);
                                }
                                #endregion Action Response


                                if (part_recommendations)
                                {
                                    ExcelWorksheet worksheetToDelete = excelPackage.Workbook.Worksheets["HRA Cover Page"];
                                    if (worksheetToDelete != null)
                                    {
                                        excelPackage.Workbook.Worksheets.Delete(worksheetToDelete);
                                    }
                                    worksheetToDelete = excelPackage.Workbook.Worksheets["Table 1"];
                                    if (worksheetToDelete != null)
                                    {
                                        excelPackage.Workbook.Worksheets.Delete(worksheetToDelete);
                                    }
                                    worksheetToDelete = excelPackage.Workbook.Worksheets["Table 2"];
                                    if (worksheetToDelete != null)
                                    {
                                        excelPackage.Workbook.Worksheets.Delete(worksheetToDelete);
                                    }
                                    worksheetToDelete = excelPackage.Workbook.Worksheets["Table 3 Assess"];
                                    if (worksheetToDelete != null)
                                    {
                                        excelPackage.Workbook.Worksheets.Delete(worksheetToDelete);
                                    }
                                    worksheetToDelete = excelPackage.Workbook.Worksheets["Matrix Rating"];
                                    if (worksheetToDelete != null)
                                    {
                                        excelPackage.Workbook.Worksheets.Delete(worksheetToDelete);
                                    }
                                    worksheetToDelete = excelPackage.Workbook.Worksheets["Drawing PIDs & PFDs"];
                                    if (worksheetToDelete != null)
                                    {
                                        excelPackage.Workbook.Worksheets.Delete(worksheetToDelete);
                                    }
                                }
                                else
                                {
                                    #region Table 1
                                    if (true)
                                    {
                                        worksheet = excelPackage.Workbook.Worksheets["Table 1"];

                                        worksheet.Cells["A4"].Value = "Company: " + name_of_area;
                                        worksheet.Cells["A5"].Value = "Assessment Team Leader: " + assessment_team_leader;
                                        worksheet.Cells["A6"].Value = "Verified By: " + qmts_reviewer;

                                        int startRows = 9;
                                        int icol_end = 5;
                                        int startRows_Def = startRows;

                                        int irows_hazard = 1;
                                        if (dtHazard?.Rows.Count > 0)
                                        {
                                            irows_hazard = dtHazard?.Rows.Count ?? 0;
                                            for (int i = 0; i < dtHazard?.Rows.Count; i++)
                                            {
                                                //worksheet.InsertRow(startRows, 1);
                                                worksheet.InsertRow(startRows, 1, 9);
                                                worksheet.Cells["C" + (startRows)].Value = (dtHazard.Rows[i]["type_hazard"] + "");
                                                worksheet.Cells["D" + (startRows)].Value = (dtHazard.Rows[i]["health_hazard"] + "");
                                                worksheet.Cells["D" + (startRows)].Value = (dtHazard.Rows[i]["health_effect_rating"] + "");
                                                startRows++;
                                            }
                                        }

                                        //Name of Area
                                        MergeCell(worksheet, "A" + startRows_Def, "A" + (startRows_Def + irows_hazard));
                                        worksheet.Cells["A" + (startRows_Def)].Value = name_of_area;

                                        //Sub Area
                                        MergeCell(worksheet, "B" + startRows_Def, "B" + (startRows_Def + irows_hazard));
                                        string sub_area_text = "";
                                        for (int i = 0; i < dtSubareas?.Rows.Count; i++)
                                        {
                                            if (i > 0) { sub_area_text += Environment.NewLine; }
                                            sub_area_text += (i + 1) + ". " + dtSubareas.Rows[i]["sub_area"] + "(" + dtSubareas.Rows[i]["work_of_task"] + ")";
                                        }
                                        worksheet.Cells["A" + (startRows_Def)].Value = sub_area_text;


                                        worksheet.Cells["A" + (startRows_Def)].Style.WrapText = true;
                                        worksheet.Cells["B" + (startRows_Def)].Style.WrapText = true;

                                        // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                                        DrawTableBorders(worksheet, startRows_Def, 1, startRows, icol_end + 1);


                                    }
                                    #endregion Table 1

                                    #region Table 2
                                    if (true)
                                    {
                                        worksheet = excelPackage.Workbook.Worksheets["Table 2"];

                                        worksheet.Cells["A4"].Value = "Company : " + name_of_area;

                                        int startRows = 7;
                                        int icol_end = 3;
                                        int startRows_Def = startRows;

                                        if (dtTasks?.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < dtTasks?.Rows.Count; i++)
                                            {


                                                int irow_worker = 1;
                                                string worker_text = "";
                                                string id_tasks = (dtTasks.Rows[i]["id"] + "");
                                                //DataRow[] dr = dtWorker.Select("id_tasks = '" + id_tasks + "'");
                                                var filterParameters = new Dictionary<string, object>();
                                                filterParameters.Add("id_tasks", id_tasks);
                                                var (dr, iMerge) = FilterDataTable(dtWorker, filterParameters);
                                                if (dr != null)
                                                {
                                                    if (dr?.Length > 0)
                                                    {
                                                        irow_worker = dr.Length;
                                                        for (int j = 0; j < dr.Length; j++)
                                                        {
                                                            if (j > 0) { worker_text += Environment.NewLine; }
                                                            worker_text += "  " + dr[j]["user_displayname"];
                                                        }
                                                    }
                                                }

                                                worksheet.InsertRow(startRows, 1, 7);// Environment.NewLine
                                                worksheet.Cells["A" + (startRows)].Value = (dtTasks.Rows[i]["worker_group"] + "") + Environment.NewLine + worker_text;
                                                worksheet.Cells["B" + (startRows)].Value = (irow_worker);
                                                worksheet.Cells["C" + (startRows)].Value = (dtTasks.Rows[i]["work_or_task"] + "");
                                                startRows++;
                                            }
                                            // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                                            DrawTableBorders(worksheet, startRows, 1, startRows - 1, icol_end + 1);


                                        }

                                    }
                                    #endregion Table 2

                                    #region Table 3 Assess
                                    if (true)
                                    {
                                        worksheet = excelPackage.Workbook.Worksheets["Table 3 Assess"];

                                        //worksheet.Cells["A4"].Value = "Company : " + name_of_area;

                                        int startRows = 7;
                                        int icol_end = 15;
                                        int startRows_Def = startRows;

                                        if (dtWorksheet?.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < dtWorksheet?.Rows.Count; i++)
                                            {
                                                int irow_worker = 0;
                                                string values_standard = (dtWorksheet.Rows[i]["values_standard"] + "");

                                                if ((dtWorksheet.Rows[i]["row_type"] + "") == "tasks" || (dtWorksheet.Rows[i]["row_type"] + "") == "tasks")
                                                {
                                                    //worksheet.InsertRow(startRows_Def, 1); 
                                                    worksheet.InsertRow(startRows, 1, 7);// Environment.NewLine
                                                    worksheet.Cells["A" + (startRows)].Value = (dtWorksheet.Rows[i]["worker_group"] + "");
                                                    worksheet.Cells["B" + (startRows)].Value = (dtWorksheet.Rows[i]["activity"] + "");
                                                    worksheet.Cells["C" + (startRows)].Value = name_of_area;

                                                    worksheet.Cells["D" + (startRows)].Value = (dtWorksheet.Rows[i]["type_hazard"] + "");
                                                    worksheet.Cells["E" + (startRows)].Value = (dtWorksheet.Rows[i]["health_effect_rating"] + "");

                                                    startRows++;
                                                }

                                                //worksheet.InsertRow(startRows_Def, 1); 
                                                worksheet.InsertRow(startRows, 1, 7);// Environment.NewLine
                                                worksheet.Cells["D" + (startRows)].Value = (dtWorksheet.Rows[i]["responder_user_displayname"] + "");
                                                worksheet.Cells["F" + (startRows)].Value = (dtWorksheet.Rows[i]["frequency_level"] + "");
                                                worksheet.Cells["G" + (startRows)].Value = (dtWorksheet.Rows[i]["exposure_band"] + "");
                                                worksheet.Cells["H" + (startRows)].Value = values_standard;
                                                worksheet.Cells["I" + (startRows)].Value = (dtWorksheet.Rows[i]["exposure_level"] + "");
                                                worksheet.Cells["J" + (startRows)].Value = (dtWorksheet.Rows[i]["exposure_rating"] + "");
                                                worksheet.Cells["K" + (startRows)].Value = (dtWorksheet.Rows[i]["initial_risk_rating"] + "");
                                                worksheet.Cells["L" + (startRows)].Value = (dtWorksheet.Rows[i]["hierarchy_of_control"] + "");
                                                worksheet.Cells["M" + (startRows)].Value = (dtWorksheet.Rows[i]["effective"] + "");
                                                worksheet.Cells["N" + (startRows)].Value = (dtWorksheet.Rows[i]["residual_risk_rating"] + "");
                                                worksheet.Cells["O" + (startRows)].Value = (dtWorksheet.Rows[i]["recommendations"] + "");
                                                //try
                                                //{
                                                //    worksheet.Cells["O" + (startRows)].Style.Fill.SetBackground(OfficeOpenXml.Drawing.eThemeSchemeColor.Hyperlink);
                                                //}
                                                //catch { }
                                                startRows++;

                                            }
                                        }
                                        worksheet.Cells["A" + startRows_Def + ":O" + startRows].Style.WrapText = true;

                                    }
                                    #endregion Table 3 Assess
                                }

                            }
                        }
                    Next_Line:;
                    }
                }
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }
            return msg_error;
        }

        public string excel_potential_health_checklist_template(string seq, string file_fullpath_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string msg_error = "";

            try
            {
                List<SqlParameter> parameters = new List<SqlParameter>();

                cls = new ClassFunctions();
                sqlstr = @"  select distinct case when ums.user_name is null  then h.request_user_displayname else case when ums.departments is null  then  ums.user_displayname else  ums.user_displayname + ' (' + ums.departments +')' end end request_user_displayname
                     from epha_t_header h 
                     inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                     inner join vw_epha_person_details ums on lower(h.request_user_name) = lower(ums.user_name)  
                     where h.seq = @seq ";


                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                DataTable dtReq = new DataTable();
                //dtReq = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
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
                        dtReq = new DataTable();
                        dtReq = _conn.ExecuteAdapter(command).Tables[0];
                        dtReq.AcceptChanges();
                    }
                    catch { }
                    finally { _conn.CloseConnection(); }
                }
                catch { }
                #endregion Execute to Datable

                if (dtReq?.Rows.Count > 0)
                {
                    string request_user_displayname = dtReq.Rows[0]["request_user_displayname"]?.ToString() ?? "";

                    if (!string.IsNullOrEmpty(file_fullpath_name))
                    {
                        string file_fullpath_def = file_fullpath_name;
                        string folder = "hra";//sub_software;


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

                    if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(file_fullpath_name))
                    {
                        FileInfo template_excel = new FileInfo(file_fullpath_name);
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using (ExcelPackage excelPackage = new ExcelPackage(template_excel))
                        {
                            ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["HEALTH Check List"];
                            ExcelWorksheet worksheet = sourceWorksheet;

                            worksheet.Cells["A2"].Value = "Initiator : " + request_user_displayname;

                            excelPackage.Save();
                        }
                    }
                }
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }


            return msg_error;
        }

        public string export_potential_health_checklist_template(ReportModel param, Boolean res_fullpath = false)
        {
            string msg_error = "";
            DataSet _dsData = new DataSet();
            DataTable dtdef = (DataTable)ClassFile.DatatableFile();
            string _file_name = "";
            string _file_download_name = "";
            string _file_fullpath_name = "";

            try
            {
                if (param == null) { msg_error = $"The specified file does not param."; }
                else
                {
                    string seq = param?.seq ?? "";
                    string export_type = param?.export_type ?? "";
                    string sub_software = param.sub_software ?? "";
                    string folder = sub_software ?? "";
                    string file_name = "HRA Potential Health Checklist Template.xlsx";
                    if (!string.IsNullOrEmpty(seq) && !string.IsNullOrEmpty(sub_software))
                    {
                        // Save the workbook as PDF
                        if (export_type == "pdf")
                        {
                            file_name = "HRA Potential Health Checklist Template.pdf";
                        }
                        //copy template to new file report  
                        msg_error = ClassFile.copy_file_duplicate(file_name, ref _file_name, ref _file_download_name, ref _file_fullpath_name, folder);

                        if (string.IsNullOrEmpty(seq) || string.IsNullOrEmpty(_file_name) || string.IsNullOrEmpty(_file_download_name) || string.IsNullOrEmpty(_file_fullpath_name))
                        { msg_error = "Invalid folder."; }
                        else
                        {
                            if (!(export_type == "pdf"))
                            {
                                if (!string.IsNullOrEmpty(_file_fullpath_name))
                                {
                                    string file_fullpath_def = _file_fullpath_name;
                                    //string folder = "hra";//sub_software;


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
                                        _file_fullpath_name = fullPath;
                                    }
                                    else { _file_fullpath_name = ""; }
                                }

                                if (string.IsNullOrEmpty(msg_error) && !string.IsNullOrEmpty(_file_fullpath_name))
                                {
                                    msg_error = excel_potential_health_checklist_template(seq, _file_fullpath_name);
                                    if (!string.IsNullOrEmpty(msg_error)) { goto Next_Line; }
                                }
                            }
                        }
                    }

                Next_Line:;

                }
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            if (res_fullpath)
            {
                return _file_fullpath_name;
            }
            else
            {
                if (dtdef != null)
                {
                    ClassFile.AddRowToDataTable(ref dtdef, _file_name, _file_download_name, msg_error);
                    DataSet dsData = new DataSet();
                    _dsData.Tables.Add(dtdef.Copy());
                }
                return JsonConvert.SerializeObject(_dsData, Formatting.Indented);
            }
        }
        #endregion export excel what'if


        //public string test_export_to_pdf()
        //{
        //    string msg_error = "";
        //    string export_type = "pdf";
        //    string _file_fullpath_name = @"C:\Users\2bLove\source\repos\dotnet-epha-api-8\bin\Debug\net8.0\wwwroot\AttachedFileTemp\Hazop\Recommendation- 202410081548.xlsx";

        //    try
        //    {
        //        // ตรวจสอบการมีอยู่ของ LibreOffice ใน tools directory
        //        string libreOfficePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "LibreOffice", "program", "soffice.exe");
        //        if (!File.Exists(libreOfficePath))
        //        {
        //            return "LibreOffice executable not found in project directory.";
        //        }

        //        if (export_type == "pdf")
        //        {
        //            try
        //            {
        //                // ตรวจสอบว่ามีไฟล์และไม่ถูกล็อค
        //                FileInfo template = new FileInfo(_file_fullpath_name);
        //                if (!template.Exists || template.IsReadOnly)
        //                {
        //                    if (template.IsReadOnly)
        //                    {
        //                        template.IsReadOnly = false;
        //                    }
        //                    else
        //                    {
        //                        return "File permissions are not correctly set.";
        //                    }
        //                }

        //                // ใช้ LibreOffice ในการแปลงไฟล์
        //                string _file_fullpath_name_pdf = _file_fullpath_name.Replace(".xlsx", ".pdf");

        //                var process = new System.Diagnostics.Process();
        //                process.StartInfo.FileName = libreOfficePath;
        //                process.StartInfo.Arguments = $"--headless --convert-to pdf --outdir \"{Path.GetDirectoryName(_file_fullpath_name)}\" \"{_file_fullpath_name}\"";
        //                process.StartInfo.CreateNoWindow = true;
        //                process.StartInfo.UseShellExecute = false;
        //                process.StartInfo.RedirectStandardOutput = true;
        //                process.StartInfo.RedirectStandardError = true;  // เพิ่มการ redirect error output
        //                process.Start();

        //                // รอการแปลงไฟล์
        //                string output = process.StandardOutput.ReadToEnd();
        //                string errorOutput = process.StandardError.ReadToEnd();  // อ่าน error output
        //                process.WaitForExit();
        //                int exitCode = process.ExitCode;

        //                // ตรวจสอบผลลัพธ์จากการแปลง
        //                if (exitCode != 0)
        //                {
        //                    msg_error = $"Error in converting file to PDF. Exit code: {exitCode}, Error details: {errorOutput}";
        //                }
        //                else
        //                {
        //                    msg_error = "Conversion successful. Output: " + output;
        //                }
        //            }
        //            catch (Exception ex_pdf)
        //            {
        //                msg_error = "PDF Conversion Error: " + ex_pdf.Message.ToString();
        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        msg_error = "General Error: " + ex.Message.ToString();
        //    }

        //    return msg_error;
        //}

        //public string test_export_to_pdfx()
        //{
        //    string excelFilePath = @"C:\Users\2bLove\source\repos\dotnet-epha-api-8\bin\Debug\net8.0\wwwroot\AttachedFileTemp\Hazop\Recommendation- 202410081548.xlsx";
        //    string libreOfficePath = @"C:\Program Files\LibreOffice\program\soffice.exe";  // ระบุ path ที่แท้จริงของ LibreOffice
        //    string outputDirectory = Path.GetDirectoryName(excelFilePath);  // โฟลเดอร์ที่ต้องการบันทึก PDF
        //    string pdfFilePath = Path.ChangeExtension(excelFilePath, ".pdf");

        //    try
        //    {
        //        // ตรวจสอบว่า LibreOffice และไฟล์ Excel มีอยู่จริง
        //        if (!File.Exists(libreOfficePath))
        //        {
        //            return "LibreOffice executable not found.";
        //        }

        //        if (!File.Exists(excelFilePath))
        //        {
        //            return "Excel file not found.";
        //        }

        //        // ตั้งค่าการทำงานของ Process สำหรับ LibreOffice
        //        var process = new Process();
        //        process.StartInfo.FileName = libreOfficePath;
        //        process.StartInfo.Arguments = $"--headless --convert-to pdf --outdir \"{outputDirectory}\" \"{excelFilePath}\"";
        //        process.StartInfo.CreateNoWindow = true;
        //        process.StartInfo.UseShellExecute = false;
        //        process.StartInfo.RedirectStandardOutput = true;
        //        process.StartInfo.RedirectStandardError = true;

        //        process.Start();

        //        // อ่านข้อมูลการทำงานและข้อผิดพลาด (ถ้ามี)
        //        string output = process.StandardOutput.ReadToEnd();
        //        string errorOutput = process.StandardError.ReadToEnd();
        //        process.WaitForExit();

        //        // ตรวจสอบ ExitCode ของ Process เพื่อดูว่าแปลงไฟล์สำเร็จหรือไม่
        //        if (process.ExitCode != 0)
        //        {
        //            return $"Error in converting file to PDF. Exit code: {process.ExitCode}. Error details: {errorOutput}";
        //        }

        //        // ตรวจสอบว่ามีไฟล์ PDF ที่แปลงแล้วอยู่ในโฟลเดอร์
        //        if (!File.Exists(pdfFilePath))
        //        {
        //            return "PDF file not created.";
        //        }

        //        return "File converted successfully to PDF: " + pdfFilePath;
        //    }
        //    catch (Exception ex)
        //    {
        //        return "An error occurred: " + ex.Message;
        //    }
        //}


    }
}