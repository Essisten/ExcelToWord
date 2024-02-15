using System;
using System.Collections.Generic;
using System.Windows;

namespace ExcelToWord.Classes
{
    class Account
    {
        public static List<Account> Accounts = new List<Account>();
        public static int MaxID = 0;
        public int ID { get; set; }
        public string Firstname { get; set; }
        public string Secondname { get; set; }
        public int Age { get; set; }
        public int Gender { get; set; }
        public int Status { get; set; }
        public float Salary { get; set; }
        public Account(string fn, string sn, string gender, int age, string status, float salary)
        {
            ID = MaxID++;
            Firstname = fn;
            Secondname = sn;
            Age = age;
            Salary = salary;
            Gender tmp = Classes.Gender.Genders.Find(_ => _.Name.Equals(gender));
            if (tmp == null)
            {
                MessageBox.Show("Не удалось найти идентификатор для пола " + gender, "Возникла ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                Gender = 0;
            }
            else
                Gender = tmp.ID;
            Status tmp2 = Classes.Status.Statuses.Find(_ => _.Name.Equals(status));
            if (tmp2 == null)
            {
                MessageBox.Show("Не удалось найти идентификатор для статуса " + status, "Возникла ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                Status = 0;
            }
            else
                Status = tmp2.ID;
        }
    }
}
