using Class;
using Newtonsoft.Json;
using System.Data;
using System.Data.SqlClient;

namespace dotnet6_epha_api.Class
{
    public class ClassTransactionLog
    {
        String ConnStrSQL = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionConfig")["ConnString"];
       
        public string insert_log(string user_name,  string role_type,  string module, string sub_software, string jsontext, ref string _token)
        {
            string token = Guid.NewGuid().ToString();
            try
            {
                string query = "INSERT INTO EPHA_T_TRANSACTIONLOG (token, module, sub_software, jsontext, create_date) VALUES (@token, @module, @sub_software, @jsontext, getdate())";

                using (SqlConnection connection = new SqlConnection(ConnStrSQL))
                {
                    SqlCommand command = new SqlCommand(query, connection);
                    command.Parameters.AddWithValue("@token", token ?? "");
                    command.Parameters.AddWithValue("@module", module ?? "");
                    command.Parameters.AddWithValue("@sub_software", sub_software ?? "");
                    command.Parameters.AddWithValue("@jsontext", jsontext ?? "");
                    if (!ClassLogin.IsAuthorized(user_name))
                    {
                        return "Parameters must be provided for a parameterized SQL query.";
                    }
                    else
                    {
                        connection.Open();
                        int result = command.ExecuteNonQuery();

                        if (result > 0)
                        {
                            _token = token ?? "";
                            return "";
                        }
                    }
                }
            }
            catch (Exception e) { token = ""; return e.Message.ToString(); }

            return "false";
        }
        public string load_log(string _token)
        {
            try
            {
                string query = "SELECT * FROM EPHA_T_TRANSACTIONLOG WHERE token = @token";

                using (SqlConnection connection = new SqlConnection(ConnStrSQL))
                {
                    SqlCommand command = new SqlCommand(query, connection);
                    command.Parameters.AddWithValue("@token", _token ?? "");

                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataSet dsRead = new DataSet();
                    adapter.Fill(dsRead);

                    return JsonConvert.SerializeObject(dsRead, Formatting.Indented);
                }
            }
            catch (Exception e)
            {
                return e.Message.ToString();
            }
        }

    }
}
