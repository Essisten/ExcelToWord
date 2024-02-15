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
                SqlCommand command = new SqlCommand("SELECT * FROM Status", connection);
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Statuses.Add(new Status(reader.GetValue(1).ToString(), (int)reader.GetValue(0)));
                    }
                }
                string query = "INSERT INTO Status (ID, Status) VALUES";
                for (int i = 0; i < Statuses.Count; i++)
                {
                    query += $" ({Statuses[i].ID}, {Statuses[i].Name})";
                    if (i + 1 == Statuses.Count)
                        query += ";";
                    else
                        query += ",";
                }
            }
        }
    }
}
