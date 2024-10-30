
using System.Data;
using System.Data.SqlClient;
using System;
using System.IO;


namespace dotnet6_epha_api.Class
{
    public class ClassFunctions
    {
        #region Utility Functions

        public string ChkSqlNum(object str, string nType)
        {
            if (str == null || Convert.IsDBNull(str) || (str?.ToString() ?? "").ToUpper() == "NULL")
                return "NULL";

            try
            {
                return nType switch
                {
                    "N" => Convert.ToInt64(str).ToString(),
                    "D" => Convert.ToDouble(str).ToString(),
                    _ => "NULL"
                };
            }
            catch
            {
                return "NULL";
            }
        }

        public string ChkSqlNum(object str, string nType, int iLength)
        {
            if (str == null || Convert.IsDBNull(str) || (str?.ToString() ?? "").ToUpper() == "NULL")
                return "NULL";

            try
            {
                double num = Convert.ToDouble(str);
                return nType switch
                {
                    "N" => Convert.ToInt64(num).ToString(),
                    "D" => num.ToString($"F{iLength}"),
                    _ => "NULL"
                };
            }
            catch
            {
                return "NULL";
            }
        }

        public string ChkSqlStr(object str, int length)
        {
            try
            {
                // ตรวจสอบว่า str เป็น null, DBNull หรือว่างเปล่า
                if (str == null || Convert.IsDBNull(str) || string.IsNullOrWhiteSpace(str.ToString()))
                    return "null";

                // แปลงค่า str เป็น string และแทนที่ ' ด้วย ''
                string str1 = str.ToString().Replace("'", "''");

                // ตัดให้ความยาวไม่เกิน length และคืนค่า
                return $"'{(str1.Length > length ? str1.Substring(0, length) : str1)}'";
            }
            catch { return "null"; }
        }
        public string ChkSqlDateYYYYMMDD(object sDate)
        {
            if (sDate == null || Convert.IsDBNull(sDate) || string.IsNullOrWhiteSpace(sDate.ToString()))
                return "NULL";

            try
            {
                string[] dateParts = sDate.ToString().Split('-');
                if (dateParts.Length == 3)
                {
                    sDate = $"{dateParts[0]}{dateParts[1].PadLeft(2, '0')}{dateParts[2].PadLeft(2, '0')}";
                }

                DateTime tsDate = DateTime.ParseExact(sDate.ToString(), "yyyyMMdd", null);

                if (tsDate.Year > 2500)
                {
                    tsDate = tsDate.AddYears(-543);
                }
                if (tsDate.Year < 2000)
                {
                    tsDate = tsDate.AddYears(543);
                }

                return $"CONVERT(date, '{tsDate:yyyyMMdd}')";
            }
            catch
            {
                return "NULL";
            }
        }

        #endregion Utility Functions
    }
}

