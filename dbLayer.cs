using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;


namespace NetPositive_FT_KWH
{
    class dbLayer
    {

        private string SqlCon;

        public dbLayer()
        {
            this.SqlCon = ConfigurationManager.AppSettings["Sqlconnstring"].ToString();
           // SqlCon = Form1.connectionString;
        }

        public int ExecSqlNonQuery(
    string strSQL,
    CommandType cmdType,
    List<SqlParameter> ListSqlParams)
        {
            // Using a 'using' statement ensures the connection and resources are properly disposed
            using (SqlConnection sqlConnection = new SqlConnection(this.SqlCon))
            {
                SqlCommand cmd = new SqlCommand();
                try
                {
                    // Set up the command object
                    this.getSqlPara(ListSqlParams, cmd);
                    cmd.CommandType = cmdType;
                    cmd.Connection = sqlConnection;

                    // Open the connection
                    sqlConnection.Open();

                    // Set the command text (SQL or stored procedure name)
                    cmd.CommandText = strSQL;

                    // Execute the command and return the result
                    return cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    // Log the exception or throw a custom exception with more context if needed
                    throw new Exception("Error executing SQL command.", ex);
                }
            }
        }

        private void getSqlPara(List<SqlParameter> parameters, SqlCommand cmd)
        {
            if (parameters == null)
                return;

            // Add each parameter to the SqlCommand
            foreach (var parameter in parameters)
            {
                cmd.Parameters.Add(parameter);
            }
        }
        public List<string> GetTagsFromDatabase()
        {
            List<string> tags = new List<string>();

            using (SqlConnection conn = new SqlConnection(SqlCon))
            {
                conn.Open();
                string query = "SELECT TagName FROM Tags"; // Assume you have a table named 'Tags' with TagName column

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        tags.Add(reader["TagName"].ToString());
                    }
                }
            }

            return tags;
        }
    }
}

