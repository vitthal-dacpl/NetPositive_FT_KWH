
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;

namespace NetPositive_FT_KWH
{

    class BusinessLayer
    {

        private dbLayer dbl = new dbLayer();


        public void FileWriter(string s)
        {
            try
            {
                string str1 = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location).ToString();
                DateTime today = DateTime.Today;
                string str2 = today.ToString("dd - MMM - yyyy").Substring(today.ToString("dd - MMM - yyyy").Length - 4);
                string str3 = today.ToString("MMM");
                string str4 = today.ToString("dd-MMM-yyyy");
                using (StreamWriter streamWriter = new StreamWriter(Directory.CreateDirectory(str1 + "\\" + str2 + "\\" + str3 + "\\" + str4).FullName + "\\ AutoBackupLogFile_" + str4 + ".txt", true))
                {
                    streamWriter.WriteLine("                                                                             ");
                    streamWriter.WriteLine("--------------------------------------------------------------------------------------------------------------");
                    streamWriter.WriteLine("--------------------------------------------------------------------------------------------------------------");
                    streamWriter.WriteLine("                                                                        ");
                    streamWriter.WriteLine(DateTime.Now.ToString() + "  " + s);
                    streamWriter.Flush();
                }
            }
            catch (Exception ex)
            {
                this.tbIlnsertLogData((object)("Exception at " + (object)DateTime.Now + " " + (object)ex));
                this.FileWriter("                                                                        ");
                this.FileWriter("--------------------------------------------------------------------------------------------------------------");
                this.FileWriter("                                                                        ");
            }
        }


        public int InsertData(
     object EM1, object EM2, object EM3, object EM4, object EM5, object EM6, object EM7, object EM8, object EM9, object EM10,
     object EM11, object EM12, object EM13, object EM14, object EM15, object EM16, object EM17, object EM18, object EM19, object EM20,
     object EM21, object EM22, object EM23, object EM24, object EM25, object EM26, object EM27, object EM28, object EM29, object EM30,
     object EM31, object EM32, object EM33, object EM34, object EM35, object EM36, object EM37, object EM38, object EM39, object EM40,
     object EM41, object EM42, object EM43, object EM44, object EM45, object EM46, object EM47, object EM48, object EM49, object EM50
 )
        {
            // Initialize the list to hold the parameters
            List<SqlParameter> sqlParams = new List<SqlParameter>();

            // Loop through the 50 parameters and add them to the list
            for (int i = 1; i <= 50; i++)
            {
                // Use reflection to get the value of the parameter by its name
                var paramValue = this.GetType().GetField($"EM{i}")?.GetValue(this);

                // Create a parameter with the appropriate name and value
                string paramName = $"@em{i}";  // Dynamic parameter name (@em1, @em2, ..., @em50)
                sqlParams.Add(new SqlParameter(paramName, paramValue ?? DBNull.Value));  // Add parameter to the list
            }

            // Execute the SQL query with the generated parameters
            return this.dbl.ExecSqlNonQuery("USP_INSERT_DATA", CommandType.StoredProcedure, sqlParams);
        }


        public List<string> Gettags()
        {
            return dbl.GetTagsFromDatabase();
        }

        public int tbIlnsertLogData(object LogData) => this.dbl.ExecSqlNonQuery("sp_InsertLogData", CommandType.StoredProcedure, new List<SqlParameter>()
    {
      new SqlParameter("@LogData", LogData)
    });


        public int InsertData(Dictionary<string , string> Values)
        {
            // Initialize the list to hold the parameters
            List<SqlParameter> sqlParams = new List<SqlParameter>();

            // Loop through the 50 parameters and add them to the list
            for (int i = 0; i < Values.Count; i++)
            {
                // Use reflection to get the value of the parameter by its name
                var paramValue = this.GetType().GetField($"EM{i+1}")?.GetValue(this);

                // Create a parameter with the appropriate name and value
                string paramName = $"@em{i+1}";  // Dynamic parameter name (@em1, @em2, ..., @em50)
                sqlParams.Add(new SqlParameter(paramName, paramValue ?? DBNull.Value));  // Add parameter to the list
            }

            // Execute the SQL query with the generated parameters
            return this.dbl.ExecSqlNonQuery("USP_INSERT_DATA", CommandType.StoredProcedure, sqlParams);
        }
    }
}

