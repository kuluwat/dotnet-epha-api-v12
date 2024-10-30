using Newtonsoft.Json;
using System.Data;

namespace Class
{
    public class ClassJSON
    { 
        public static string SetJSONresultRef(DataTable _dtJson)
        {
            string JSONresult;
            DataSet ds = new DataSet();
            ds.Tables.Add(_dtJson.Copy());
            ds.Tables[0].TableName = "msg";

            JSONresult = JsonConvert.SerializeObject(ds);
            JSONresult = JsonConvert.SerializeObject(_dtJson);
            return JSONresult;
        }
        public string SetJSONresult(DataTable? _dtJson)
        {
            string JSONresult;
            JSONresult = JsonConvert.SerializeObject(_dtJson);
            return JSONresult;
        }
        public DataTable ConvertJSONresult(string user_name, string role_type, String jsper)
        {
            DataTable _dtJson = new DataTable();
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!ClassLogin.IsAuthorized(user_name)) { return _dtJson; }

            try
            {
                // ตรวจสอบว่า JSON ไม่ใช่ค่าว่างก่อนทำการแปลง
                if (!string.IsNullOrEmpty(jsper))
                {
                    _dtJson = (DataTable)JsonConvert.DeserializeObject(jsper, typeof(DataTable));

                    if (_dtJson != null && _dtJson.Rows.Count > 0)
                    {
                        // ตรวจสอบค่าในแถวและคอลัมน์ก่อนทำการลบแถว
                        if (_dtJson.Columns.Contains("json_check_null"))
                        {
                            if (_dtJson.Rows[0]["json_check_null"].ToString() == "true")
                            {
                                if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                                {
                                    _dtJson.Rows[0].Delete();
                                    _dtJson.AcceptChanges();
                                }
                            }
                            if (ClassLogin.IsAuthorizedRole(user_name, role_type))
                            {
                                // ลบคอลัมน์ "json_check_null" ถ้ามี
                                _dtJson.Columns.Remove("json_check_null");
                            }
                            _dtJson.AcceptChanges();
                        }
                    }
                }
            }
            catch (JsonException jsonEx)
            {
                // จัดการข้อผิดพลาดที่เกี่ยวข้องกับการแปลง JSON
                throw new Exception("Error in converting JSON to DataTable: " + jsonEx.Message);
            }
            catch (Exception ex)
            {
                // จัดการข้อผิดพลาดทั่วไป
                throw new Exception("An error occurred in ConvertJSONresult: " + ex.Message);
            }

            return _dtJson;
        }


    }
}
