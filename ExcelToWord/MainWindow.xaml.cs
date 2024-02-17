using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Configuration;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Xceed.Words.NET;
using ExcelToWord.Classes;

namespace ExcelToWord
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string connectionString;
        public MainWindow()
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        }

        private void ExcelButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel Documents (*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == false || dialog.FileName.Length == 0)
                return;
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(dialog.FileName);
            try
            {
                Excel.Worksheet sheet = wb.Worksheets[1];
                Gender.SyncToDB(connectionString);
                Status.SyncToDB(connectionString);
                Account.Accounts.Clear();
                string query_status = "", query_gender = "";
                for (int i = 2; i > 0; i++)
                {
                    if (sheet.Cells[i, 1].Value == null)
                        break;
                    if (Gender.Genders.Find(g => g.Name == sheet.Cells[i, 3].Value.ToString()) == null)
                    {
                        Gender g = new Gender(sheet.Cells[i, 3].Value.ToString());
                        Gender.Genders.Add(g);
                        query_gender += $"INSERT INTO Account_gender (ID, Gender) VALUES ({g.ID}, '{g.Name}');";
                    }
                    if (Status.Statuses.Find(s => s.Name == sheet.Cells[i, 5].Value.ToString()) == null)
                    {
                        Status s = new Status(sheet.Cells[i, 5].Value.ToString());
                        Status.Statuses.Add(s);
                        query_status += $"INSERT INTO Account_status (ID, Status) VALUES ({s.ID}, '{s.Name}');";
                    }
                }
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(query_gender + query_status, connection);
                    int result = 0;
                    if (command.CommandText.Length > 0)
                        result = command.ExecuteNonQuery();
                    for (int i = 2; i > 0; i++)
                    {
                        if (sheet.Cells[i, 1].Value == null)
                            break;
                        Account.Accounts.Add(new Account(sheet.Cells[i, 1].Value.ToString(), sheet.Cells[i, 2].Value.ToString(),
                            sheet.Cells[i, 3].Value.ToString(), Convert.ToInt32(sheet.Cells[i, 4].Value), sheet.Cells[i, 5].Value.ToString(),
                            (float)sheet.Cells[i, 6].Value));
                    }
                    string query_account = "INSERT INTO Accounts (ID, Firstname, Secondname, Gender, Age, Status, Salary) VALUES";
                    for (int i = 0; i < Account.Accounts.Count; i++)
                    {
                        Account a = Account.Accounts[i];
                        query_account += $" ({a.ID}, '{a.Firstname}', '{a.Secondname}', {a.Gender}, {a.Age}, {a.Status}, {a.Salary})";
                        if (i + 1 == Account.Accounts.Count)
                            query_account += ";";
                        else
                            query_account += ",";
                    }
                    command.CommandText = query_account;
                    if (command.CommandText.Length > 0)
                        result += command.ExecuteNonQuery();
                    MessageBox.Show($"Добавлено {result} записей в БД", "Завершение операции", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "Возникла ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                wb.Close();
                excel.Quit();
            }
        }

        private void WordButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Word Documents (*.docx)|*.docx";
            if (dialog.ShowDialog() == false || dialog.FileName.Length == 0)
                return;
            try
            {
                using (DocX file = DocX.Create(dialog.FileName))
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();

                        //Мучины и женщины
                        Gender.SyncToDB(connectionString);
                        int male_id = -1, female_id = -1;
                        Gender g = Gender.Genders.Find(_ => _.Name == "м");
                        if (g != null)
                            male_id = g.ID;
                        g = Gender.Genders.Find(_ => _.Name == "ж");
                        if (g != null)
                            female_id = g.ID;

                        SqlCommand command = new SqlCommand($"SELECT COUNT(*) FROM Accounts WHERE Gender = {male_id};", connection);
                        SqlDataReader reader = command.ExecuteReader();
                        reader.Read();
                        int male_amount = reader.GetInt32(0);
                        reader.Close();

                        command.CommandText = $"SELECT COUNT(*) FROM Accounts WHERE Gender = {female_id};";
                        reader = command.ExecuteReader();
                        reader.Read();
                        int female_amount = reader.GetInt32(0);
                        reader.Close();

                        var p1 = file.InsertParagraph();
                        p1.Append($"Мужчин: {male_amount}. Женщин: {female_amount}.");

                        //Мужчины в возрасте 30-40 лет
                        command.CommandText = $"SELECT COUNT(*) FROM Accounts WHERE Gender = {male_id} AND Age BETWEEN 30 AND 40;";
                        reader = command.ExecuteReader();
                        reader.Read();
                        male_amount = reader.GetInt32(0);
                        reader.Close();

                        var p2 = file.InsertParagraph();
                        p2.Append($"Мужчин в возрасте 30-40 лет: {male_amount}");

                        //Стандартные и премиум аккаунты
                        Status.SyncToDB(connectionString);
                        int standart_id = -1, premium_id = 0;
                        Status s = Status.Statuses.Find(_ => _.Name == "стандарт");
                        if (s != null)
                            standart_id = s.ID;
                        s = Status.Statuses.Find(_ => _.Name == "премиум");
                        if (s != null)
                            premium_id = s.ID;

                        command.CommandText = $"SELECT COUNT(*) FROM Accounts WHERE Status = {standart_id};";
                        reader = command.ExecuteReader();
                        reader.Read();
                        int standart_amount = reader.GetInt32(0);
                        reader.Close();

                        command.CommandText = $"SELECT COUNT(*) FROM Accounts WHERE Status = {premium_id};";
                        reader = command.ExecuteReader();
                        reader.Read();
                        int premium_amount = reader.GetInt32(0);
                        reader.Close();

                        var p3 = file.InsertParagraph();
                        p3.Append($"Стандартных аккаунтов: {standart_amount}. Премиум: {premium_amount}.");

                        //Женщины с премиум аккаунтами от 30 лет
                        command.CommandText = $"SELECT COUNT(*) FROM Accounts WHERE Gender = {female_id} AND Age >= 30;";
                        reader = command.ExecuteReader();
                        reader.Read();
                        female_amount = reader.GetInt32(0);
                        reader.Close();

                        var p4 = file.InsertParagraph();
                        p4.Append($"Женщин с премиум аккаунтом от 30 лет: {female_amount}");

                        //Женщины с большим окладом
                        string women_list = "";
                        command.CommandText = $"SELECT Firstname, Secondname FROM Accounts " +
                            $"WHERE Gender = {female_id} AND Age BETWEEN 23 AND 35 ORDER BY Salary DESC;";
                        var p5 = file.InsertParagraph();
                        p5.Append($"Женщины с лучшим окладом 23-35 лет:");
                        var list5 = file.AddList();
                        reader = command.ExecuteReader();
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                file.AddListItem(list5, $"{reader.GetValue(0)} {reader.GetValue(1)}");
                            }
                        }
                        reader.Close();
                        file.InsertList(list5);
                        p5.LineSpacingBefore = 10;

                        //3 мужчины и 3 женщины с мин зп
                        command.CommandText = $"SELECT TOP 3 Firstname, Secondname FROM Accounts " +
                            $"WHERE Gender = {female_id} ORDER BY Salary;";
                        reader = command.ExecuteReader();
                        var p6 = file.InsertParagraph();
                        p6.Append($"Женщины с худшим окладом:");
                        var list6 = file.AddList();
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                file.AddListItem(list6, $"{reader.GetValue(0)} {reader.GetValue(1)}");
                            }
                        }
                        reader.Close();
                        p6.LineSpacingBefore = 10;
                        file.InsertList(list6);

                        command.CommandText = $"SELECT TOP 3 Firstname, Secondname FROM Accounts " +
                            $"WHERE Gender = {male_id} ORDER BY Salary;";
                        reader = command.ExecuteReader();
                        var p7 = file.InsertParagraph();
                        p7.Append($"Мужчины с худшим окладом:");
                        var list7 = file.AddList();
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                file.AddListItem(list7, $"{reader.GetValue(0)} {reader.GetValue(1)}");
                            }
                        }
                        p7.LineSpacingBefore = 10;
                        reader.Close();
                        file.InsertList(list7);

                        file.Save();
                        MessageBox.Show("Статистика сохранена", "Завершение операции", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "Возникла ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand("DELETE FROM Accounts;DELETE FROM Account_gender;DELETE FROM Account_status;", connection);
                    int result = command.ExecuteNonQuery();
                    MessageBox.Show($"Удалено {result} записей в БД");
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "Возникла ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
