

namespace dotnet_epha_api.Class
{
    using Microsoft.Exchange.WebServices.Data;
    using System.Data;
    public class ClassFile
    {

        public static DataTable refMsg(string status, string remark, string? seq_new = "")
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_new");
            dtMsg.TableName = "msg";
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            dtMsg.Rows[0]["seq_new"] = seq_new;
            return dtMsg;
        }
        public static DataTable refMsgSave(string status, string remark, string? seq_new = "", string? pha_seq = "", string? pha_no = "", string? pha_status = "")
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_new");
            dtMsg.Columns.Add("pha_seq");
            dtMsg.Columns.Add("pha_no");
            dtMsg.Columns.Add("pha_status");
            dtMsg.TableName = "msg";
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            dtMsg.Rows[0]["seq_new"] = seq_new;
            dtMsg.Rows[0]["pha_seq"] = pha_seq;
            dtMsg.Rows[0]["pha_no"] = pha_no;
            dtMsg.Rows[0]["pha_status"] = pha_status;
            return dtMsg;
        }
        public static DataTable DatatableMsg()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("status");
            return dt;
        }
        public static DataTable DatatableFile()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("STATUS");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.TableName = "msg";
            return dt;
        }
        public static void AddRowToDataTable(ref DataTable dtdef, string fileName, string filePath, string msgError)
        {
            // ตรวจสอบว่า DataTable ไม่เป็น null
            if (dtdef == null) throw new ArgumentNullException(nameof(dtdef));

            if (dtdef != null &&
                dtdef.Columns.Contains("ATTACHED_FILE_NAME") &&
                dtdef.Columns.Contains("ATTACHED_FILE_PATH") &&
                dtdef.Columns.Contains("IMPORT_DATA_MSG") &&
                dtdef.Columns.Contains("STATUS"))
            {
                // สร้างแถวใหม่
                DataRow newRow = dtdef.NewRow();
                newRow["ATTACHED_FILE_NAME"] = fileName ?? "";
                newRow["ATTACHED_FILE_PATH"] = filePath ?? "";
                newRow["IMPORT_DATA_MSG"] = msgError ?? "";
                newRow["STATUS"] = string.IsNullOrEmpty(msgError) ? "true" : "error";

                // เพิ่มแถวใหม่ลงใน DataTable
                dtdef.Rows.Add(newRow);
            }

        }
        public static string copy_file_data_to_server(
      ref string file_name, ref string file_download_name, ref string _file_fullpath_name,
      IFormFileCollection? files,
      string? folder = "_temp",
      string? file_part = "import",
      string? file_doc = "docno",
      bool tempFile = false,
      bool folderCopyFile = false)
        {
            if (files == null || files.Count == 0)
            {
                return "Invalid files.";
            }

            if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
            {
                return "Invalid folder.";
            }

            if (string.IsNullOrWhiteSpace(file_part) || file_part.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || file_part.Contains("..") || Path.IsPathRooted(file_part))
            {
                return "Invalid part.";
            }

            try
            {
                IFormFile file = files[0];
                string safeFileTemp = Path.GetFileName(file.FileName);
                if (string.IsNullOrEmpty(safeFileTemp))
                {
                    return "Invalid file name.";
                }

                // ตรวจสอบตัวอักษรที่อนุญาตในชื่อไฟล์
                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                    .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return "Input file name contains invalid characters.";
                }

                string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                if (string.IsNullOrEmpty(extension))
                {
                    return "File does not have a valid extension.";
                }

                // อนุญาตเฉพาะไฟล์ที่กำหนด
                string[] allowedExtensionsExcel = { ".xlsx", ".xls" };
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };

                if (tempFile && !allowedExtensionsExcel.Contains(extension))
                {
                    return "Invalid file type. Only Excel files are allowed.";
                }
                else if (!tempFile && !allowedExtensions.Contains(extension))
                {
                    return "Invalid file type.";
                }

                // ตรวจสอบการมีอยู่ของไดเรกทอรีหลัก
                string templatewwwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                if (!Directory.Exists(templatewwwwRootDir))
                {
                    return "Folder directory not found.";
                }

                string templateRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                if (!Directory.Exists(templateRootDir))
                {
                    return "Folder directory not found.";
                }

                // กำหนดไดเรกทอรีสุดท้ายสำหรับเก็บไฟล์
                string finalRootDir = Path.Combine(templateRootDir, folder);
                if (folderCopyFile)
                {
                    finalRootDir = Path.Combine(finalRootDir, "copy");
                }
                if (!Directory.Exists(finalRootDir))
                {
                    Directory.CreateDirectory(finalRootDir); // สร้างไดเรกทอรีถ้าไม่พบ
                    //return "Folder Module directory not found.";
                }

                // สร้างชื่อไฟล์ใหม่
                var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
                string retFileName = $"{file_doc}-{file_part}-{datetime_run}";
                string sourceFile = $"{retFileName}{extension}";

                if (sourceFile == "")
                {
                    return ("Invalid file type. Only Excel files are allowed.");
                }
                if (sourceFile.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return ("Input path contains invalid characters.");
                }
                if (string.IsNullOrWhiteSpace(sourceFile) || sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || sourceFile.Contains("..") || sourceFile.Contains($"\\"))
                {
                    return ("Invalid fileName.");
                }

                // สร้างเส้นทางไฟล์ปลายทางแบบสัมบูรณ์และตรวจสอบทีละชั้น
                string newFileNameFullPath = Path.Combine(finalRootDir, sourceFile);
                newFileNameFullPath = Path.GetFullPath(newFileNameFullPath);
                if (!newFileNameFullPath.StartsWith(finalRootDir, StringComparison.OrdinalIgnoreCase))
                {
                    return "Attempt to access unauthorized path.";
                }

                _file_fullpath_name = newFileNameFullPath;

                // คัดลอกไฟล์ไปยังเซิร์ฟเวอร์
                using (var fileStream = new FileStream(newFileNameFullPath, FileMode.Create))
                {
                    file.CopyTo(fileStream);
                }

                // ตั้งค่าตัวแปรผลลัพธ์
                file_name = safeFileTemp;
                file_download_name = Path.GetRelativePath(templatewwwwRootDir, newFileNameFullPath);

                if (file_download_name.Contains(".."))
                {
                    return "The resulting relative path is attempting to access outside the intended directory.";
                }
            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request. {ex.Message}";
            }

            return "";
        }


        public static string copy_file_excel_template(
     ref string file_name, ref string file_download_name, ref string _file_fullpath_name,
     string? folder = "other", string? file_part = "Report", string? file_doc = "template")
        {
            if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
            {
                return "Invalid folder.";
            }

            if (string.IsNullOrWhiteSpace(file_part) || file_part.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || file_part.Contains("..") || Path.IsPathRooted(file_part))
            {
                return "Invalid part.";
            }

            try
            {
                // สร้างชื่อไฟล์เทมเพลต
                string defFileTemp = $"{folder?.ToUpper()} {file_part}";
                string safeFileTemp = $"{defFileTemp} Template.xlsx";

                // ตรวจสอบตัวอักษรที่อนุญาตในชื่อไฟล์
                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                    .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return "Input file name contains invalid characters.";
                }

                // ตรวจสอบนามสกุลไฟล์
                string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                string[] allowedExtensionsExcel = { ".xlsx", ".xls" };
                if (!allowedExtensionsExcel.Contains(extension))
                {
                    return "Invalid file type. Only Excel files are allowed.";
                }

                // ตรวจสอบไดเรกทอรีหลัก
                string templatewwwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                if (!Directory.Exists(templatewwwwRootDir))
                {
                    return "Folder directory not found.";
                }
                string templateRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                if (!Directory.Exists(templateRootDir))
                {
                    return "Folder directory not found.";
                }

                // สร้างเส้นทางเต็มของไฟล์เทมเพลต
                string fileNameFullPath = Path.Combine(templateRootDir, safeFileTemp);
                fileNameFullPath = Path.GetFullPath(fileNameFullPath);
                if (!fileNameFullPath.StartsWith(templateRootDir, StringComparison.OrdinalIgnoreCase))
                {
                    return "Temp directory is outside of the Template directory.";
                }

                // ตรวจสอบไดเรกทอรีย่อย
                string templateModuleDir = Path.Combine(templateRootDir, folder);
                if (!Directory.Exists(templateModuleDir))
                {
                    return "Folder Module directory not found.";
                }

                // สร้างชื่อไฟล์ใหม่พร้อม timestamp
                var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
                string retFileName = $"{file_name}-{file_doc} {datetime_run}";
                string sourceFile = $"{retFileName}{extension}";

                if (sourceFile == "")
                {
                    return ("Invalid file type. Only Excel files are allowed.");
                }
                if (sourceFile.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return ("Input path contains invalid characters.");
                }
                if (string.IsNullOrWhiteSpace(sourceFile) || sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || sourceFile.Contains("..") || sourceFile.Contains($"\\"))
                {
                    return ("Invalid fileName.");
                }


                // สร้างเส้นทางเต็มของไฟล์ใหม่
                string newFileNameFullPath = Path.Combine(templateModuleDir, sourceFile);
                newFileNameFullPath = Path.GetFullPath(newFileNameFullPath);
                if (!newFileNameFullPath.StartsWith(templateModuleDir, StringComparison.OrdinalIgnoreCase))
                {
                    return "Attempt to access unauthorized path.";
                }

                _file_fullpath_name = newFileNameFullPath;

                // คัดลอกไฟล์จากเทมเพลตไปยังไฟล์ใหม่
                File.Copy(fileNameFullPath, newFileNameFullPath, overwrite: true);

                // ตรวจสอบว่าไฟล์ใหม่ถูกสร้างขึ้นและไม่ใช่ ReadOnly
                var checkFileInfo = new FileInfo(newFileNameFullPath);
                if (!checkFileInfo.Exists || checkFileInfo.IsReadOnly)
                {
                    return "File permissions are not correctly set.";
                }

                // กำหนดค่าตัวแปรผลลัพธ์เพื่อส่งออกข้อมูล
                file_name = sourceFile;
                file_download_name = Path.GetRelativePath(templatewwwwRootDir, newFileNameFullPath);

                // ตรวจสอบว่าเส้นทางที่ได้ไม่สามารถเข้าถึงไดเรกทอรีที่ไม่ได้รับอนุญาต
                if (file_download_name.Contains(".."))
                {
                    return "The resulting relative path is attempting to access outside the intended directory.";
                }
            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request. {ex.Message}";
            }

            return "";
        }


        public static string copy_file_duplicate(string file_name, ref string _file_name, ref string _file_download_name, ref string _file_fullpath_name, string? _folder = "other")
        {
            if (string.IsNullOrWhiteSpace(_folder) || _folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || _folder.Contains("..") || Path.IsPathRooted(_folder))
            {
                return "Invalid folder.";
            }
            if (string.IsNullOrWhiteSpace(file_name))
            {
                return "Invalid file name.";
            }
            try
            {
                _file_name = "";
                _file_download_name = "";

                // ดึงเฉพาะชื่อไฟล์โดยไม่รวมพาธ
                string safeFileTemp = Path.GetFileName(file_name);
                if (string.IsNullOrEmpty(safeFileTemp))
                {
                    return "Invalid file safe file temp.";
                }

                // ตรวจสอบตัวอักษรที่อนุญาตในชื่อไฟล์
                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                    .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return "Input file name contains invalid characters.";
                }

                // ตรวจสอบนามสกุลไฟล์
                string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                if (!allowedExtensions.Contains(extension))
                {
                    return "Invalid file type.";
                }

                // ตรวจสอบว่ามีไดเรกทอรี wwwroot อยู่หรือไม่
                string templatewwwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                if (!Directory.Exists(templatewwwwRootDir))
                {
                    return "Folder directory not found.";
                }

                // ตรวจสอบว่ามีไดเรกทอรี AttachedFileTemp อยู่หรือไม่
                string templateRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                if (!Directory.Exists(templateRootDir))
                {
                    return "Folder directory not found.";
                }

                // ตรวจสอบว่ามีไดเรกทอรี Module (เช่น "other") อยู่หรือไม่
                string templateModuleDir = Path.Combine(templateRootDir, $"{_folder}");
                if (!Directory.Exists(templateModuleDir))
                {
                    return "Folder Module directory not found.";
                }

                // สร้างเส้นทางเต็มของไฟล์ต้นฉบับ
                //string fileNameFullPath = Path.IsPathRooted(file_name) ? file_name : Path.Combine(templateModuleDir, safeFileTemp);
                string fileNameFullPath = Path.Combine(templateModuleDir, safeFileTemp);
                fileNameFullPath = Path.GetFullPath(fileNameFullPath);

                // ตรวจสอบว่าไฟล์อยู่ภายในไดเรกทอรีที่อนุญาต
                if (!fileNameFullPath.StartsWith(templateModuleDir, StringComparison.OrdinalIgnoreCase))
                {
                    return "Temp directory is outside of the Template directory.";
                }

                // สร้างชื่อไฟล์ใหม่ด้วย timestamp
                var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
                string defFileTemp = Path.GetFileNameWithoutExtension(file_name);
                string retFileName = $"{defFileTemp} {datetime_run}";
                string sourceFile = $"{retFileName}{extension}";

                if (sourceFile == "")
                {
                    return ("Invalid file type. Only Excel files are allowed.");
                }
                if (sourceFile.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return ("Input path contains invalid characters.");
                }
                if (string.IsNullOrWhiteSpace(sourceFile) || sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || sourceFile.Contains("..") || sourceFile.Contains($"\\"))
                {
                    return ("Invalid fileName.");
                }

                // สร้างเส้นทางเต็มของไฟล์ใหม่
                string newFileNameFullPath = Path.Combine(templateModuleDir, sourceFile);
                newFileNameFullPath = Path.GetFullPath(newFileNameFullPath);

                // ตรวจสอบว่าไฟล์ใหม่อยู่ภายในไดเรกทอรีที่อนุญาต
                if (!newFileNameFullPath.StartsWith(templateModuleDir, StringComparison.OrdinalIgnoreCase))
                {
                    return "Attempt to access unauthorized path.";
                }

                // คัดลอกไฟล์
                File.Copy(fileNameFullPath, newFileNameFullPath, overwrite: true);

                // ตรวจสอบว่าไฟล์ใหม่ถูกสร้างขึ้นและไม่ใช่ ReadOnly
                var checkFileInfo = new FileInfo(newFileNameFullPath);
                if (!checkFileInfo.Exists || checkFileInfo.IsReadOnly)
                {
                    return "File permissions are not correctly set.";
                }

                // กำหนดค่าตัวแปรอ้างอิงเพื่อส่งออกข้อมูล
                _file_name = sourceFile;
                _file_fullpath_name = newFileNameFullPath;
                _file_download_name = Path.GetRelativePath(templatewwwwRootDir, newFileNameFullPath);

                // ตรวจสอบว่าเส้นทางที่ได้ไม่สามารถเข้าถึงไดเรกทอรีที่ไม่ได้รับอนุญาต
                if (_file_download_name.Contains(".."))
                {
                    return "The resulting relative path is attempting to access outside the intended directory.";
                }
            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request. {ex.Message}";
            }

            return "";
        }


        public static string check_file_other(string file_name, ref string _file_fullpath_name, ref string _file_download_name, string? folder = "")
        {
            if (string.IsNullOrEmpty(file_name))
            {
                return "Invalid file name.";
            }
            try
            {
                _file_fullpath_name = "";
                _file_download_name = "";

                string safeFileTemp = Path.GetFileName(file_name);
                if (string.IsNullOrEmpty(safeFileTemp))
                {
                    return "Invalid file safe file temp.";
                }

                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                    .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return "Input file name contains invalid characters.";
                }

                string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                if (!allowedExtensions.Contains(extension))
                {
                    return "Invalid file type.";
                }

                string templatewwwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                if (!Directory.Exists(templatewwwwRootDir))
                {
                    return "Folder directory not found.";
                }

                string finalRootDir = "";

                if (string.IsNullOrWhiteSpace(folder))
                {
                    finalRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                    if (!Directory.Exists(finalRootDir))
                    {
                        return "Folder directory not found.";
                    }
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
                    {
                        return "Invalid file folder.";
                    }
                    else
                    {
                        string templateRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                        if (!Directory.Exists(templateRootDir))
                        {
                            return "Folder directory not found.";
                        }
                        if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || folder.Contains($"\\") || Path.IsPathRooted(folder))
                        {
                            return ("Invalid folder.");
                        }
                        finalRootDir = Path.Combine(templateRootDir, folder);
                        if (!Directory.Exists(finalRootDir))
                        {
                            Directory.CreateDirectory(finalRootDir);
                            //return "Folder directory not found.";
                        }
                    }
                }

                string sourceFile = $"{safeFileTemp}";
                if (sourceFile == "")
                {
                    return ("Invalid file type. Only Excel files are allowed.");
                }
                if (sourceFile.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return ("Input path contains invalid characters.");
                }
                if (string.IsNullOrWhiteSpace(sourceFile) || sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || sourceFile.Contains("..") || sourceFile.Contains($"\\"))
                {
                    return ("Invalid fileName.");
                }

                string fileNameFullPath = Path.Combine(finalRootDir, sourceFile);
                fileNameFullPath = Path.GetFullPath(fileNameFullPath);
                if (!fileNameFullPath.StartsWith(finalRootDir, StringComparison.OrdinalIgnoreCase))
                {
                    return "Temp directory is outside of the Template directory.";
                }

                _file_fullpath_name = fileNameFullPath;
                _file_download_name = Path.GetRelativePath(templatewwwwRootDir, fileNameFullPath);
                if (_file_download_name.Contains(".."))
                {
                    return "The resulting relative path is attempting to access outside the intended directory.";
                }
            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request. {ex.Message}";
            }

            return "";
        }
        public static string check_file_on_server(string? _file_fullpath_name, string? folder, ref string _file_fullpath_name_new)
        {
            if (string.IsNullOrEmpty(_file_fullpath_name))
            {
                return "Invalid file fullpath.";
            }

            try
            {
                // ตรวจสอบและรับชื่อไฟล์จากพาธเต็ม
                string safeFileTemp = Path.GetFileName(_file_fullpath_name);
                if (string.IsNullOrEmpty(safeFileTemp) || safeFileTemp.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                {
                    return "Invalid file name.";
                }

                // ตรวจสอบว่าชื่อไฟล์มีเฉพาะตัวอักษรที่อนุญาตเท่านั้น
                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                    .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return "Input file name contains invalid characters.";
                }

                // ตรวจสอบนามสกุลไฟล์
                string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                if (!allowedExtensions.Contains(extension))
                {
                    return "Invalid file type.";
                }

                // ตรวจสอบการมีอยู่ของไดเรกทอรี wwwroot
                string templatewwwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                if (!Directory.Exists(templatewwwwRootDir))
                {
                    return "Folder directory not found.";
                }

                // ตรวจสอบการมีอยู่ของไดเรกทอรี AttachedFileTemp
                string finalRootDir = "";
                // ตรวจสอบว่าไดเรกทอรีที่กำหนดโดย folder อยู่ภายในไดเรกทอรีที่ปลอดภัย
                if (!string.IsNullOrEmpty(folder))
                {
                    if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || folder.Contains("..") || Path.IsPathRooted(folder))
                    {
                        return "Invalid file folder.";
                    }
                    else
                    {
                        string templateRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                        if (!Directory.Exists(templateRootDir))
                        {
                            return "Folder directory not found.";
                        }

                        finalRootDir = Path.Combine(templateRootDir, folder);
                        if (!Directory.Exists(templateRootDir))
                        {
                            return "Folder directory not found.";
                        }
                    }
                }
                else
                {
                    finalRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                    if (!Directory.Exists(finalRootDir))
                    {
                        return "Folder directory not found.";
                    }
                }

                string sourceFile = $"{safeFileTemp}";
                if (sourceFile == "")
                {
                    return ("Invalid file type. Only Excel files are allowed.");
                }
                if (sourceFile.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return ("Input path contains invalid characters.");
                }
                if (string.IsNullOrWhiteSpace(sourceFile) || sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || sourceFile.Contains("..") || sourceFile.Contains($"\\"))
                {
                    return ("Invalid fileName.");
                }

                string fileNameFullPath = Path.Combine(finalRootDir, sourceFile);
                fileNameFullPath = Path.GetFullPath(fileNameFullPath);
                if (!fileNameFullPath.StartsWith(finalRootDir, StringComparison.OrdinalIgnoreCase))
                {
                    return "Temp directory is outside of the Template directory.";
                }

                // ตรวจสอบว่าไฟล์มีอยู่และจัดการสถานะ Read-Only
                FileInfo fileInfo = new FileInfo(fileNameFullPath);
                if (!fileInfo.Exists)
                {
                    return "File not found.";
                }
                if (fileInfo.IsReadOnly)
                {
                    try
                    {
                        fileInfo.IsReadOnly = false; // ยกเลิกสถานะ Read-Only ถ้าจำเป็น
                    }
                    catch (Exception ex)
                    {
                        return $"Failed to modify file attributes: {ex.Message}";
                    }
                }

                // ตรวจสอบขนาดไฟล์เพื่อความปลอดภัย (ถ้าต้องการ)
                long maxSizeInBytes = 50 * 1024 * 1024; // 50 MB
                if (fileInfo.Length > maxSizeInBytes)
                {
                    return "File size exceeds the allowed limit.";
                }

                // ตรวจสอบสิทธิ์ในการเข้าถึงไฟล์ (ถ้าต้องการ)
                try
                {
                    using (FileStream fs = fileInfo.Open(FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        // สามารถอ่านไฟล์ได้
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    return "Access to the file is denied.";
                }

                _file_fullpath_name_new = fileNameFullPath;
            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request. {ex.Message}";
            }

            return ""; // ถ้าไม่มีข้อผิดพลาด
        }

        public static string check_format_file_name(string file_name)
        {
            try
            {
                // ตรวจสอบว่า file_name ไม่เป็นค่าว่าง
                if (string.IsNullOrEmpty(file_name))
                {
                    return "Invalid file name.";
                }

                // ใช้ Path.GetFileName เพื่อดึงเฉพาะชื่อไฟล์โดยไม่รวมพาธ
                string safeFileTemp = Path.GetFileName(file_name);
                if (string.IsNullOrEmpty(safeFileTemp))
                {
                    return "Invalid file name.";
                }

                // ตรวจสอบว่า safeFileTemp ไม่มีองค์ประกอบที่อาจทำให้เกิด Path Traversal
                if (safeFileTemp.Contains("..") || safeFileTemp.Contains("/") || safeFileTemp.Contains("\\"))
                {
                    return "Invalid file name.";
                }

                // ตรวจสอบ Format ชื่อไฟล์ตามที่กำหนด
                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                    .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return "Input file name contains invalid characters.";
                }

                // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                if (string.IsNullOrEmpty(extension))
                {
                    return "File does not have a valid extension.";
                }

                // อนุญาตเฉพาะไฟล์ที่มีนามสกุลกำหนดไว้
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                if (!allowedExtensions.Contains(extension))
                {
                    return "Invalid file type.";
                }
            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request. {ex.Message}";
            }

            return ""; // หากชื่อไฟล์ถูกต้องตามเงื่อนไขที่กำหนด
        }


    }
}
