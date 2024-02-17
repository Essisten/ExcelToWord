using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;

namespace ExcelToWord.Classes
{
    class Gender
    {
        public static List<Gender> Genders = new List<Gender>();
        public static int MaxID = 0;
        public int ID { get; set; }
        public string Name { get; set; }
        public Gender(string gender, int id = -1)
        {
            if (id == -1)
                ID = MaxID++;
            else
            {
                ID = id;
                if (id > MaxID)
                    MaxID = id + 1;
            }
            Name = gender;
        }

        public static void SyncToDB(string connectionString)
        {
            Genders.Clear();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM Account_gender", connection);
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Genders.Add(new Gender(reader.GetValue(1).ToString(), (int)reader.GetValue(0)));
                    }
                }
                string query = "INSERT INTO Account_gender (ID, Gender) VALUES";
                for (int i = 0; i < Genders.Count; i++)
                {
                    query += $" ({Genders[i].ID}, {Genders[i].Name})";
                    if (i + 1 == Genders.Count)
                        query += ";";
                    else
                        query += ",";
                }
            }
        }
    }
}
