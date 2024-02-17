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
                    MessageBox.Show($"Добавлено {result} записей в БД");
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
