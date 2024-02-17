using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;

namespace ExcelToWord.Classes
{
    class Status
    {
        public static List<Status> Statuses = new List<Status>();
        public static int MaxID = 0;
        public int ID { get; set; }
        public string Name { get; set; }
        public Status(string status, int id = -1)
        {
            if (id == -1)
                ID = MaxID++;
            else
            {
                ID = id;
                if (id > MaxID)
                    MaxID = id + 1;
            }
            Name = status;
        }

        public static void SyncToDB(string connectionString)
        {
            Statuses.Clear();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM Account_status", connection);
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Statuses.Add(new Status(reader.GetValue(1).ToString(), (int)reader.GetValue(0)));
                    }
                }
            }
        }
    }
}
