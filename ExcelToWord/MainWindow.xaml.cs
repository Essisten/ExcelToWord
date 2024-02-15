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
        SqlDataAdapter adapter;
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
                Excel.Worksheet sheet = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                Excel.Range range = sheet.UsedRange;
                Gender.SyncToDB(connectionString);
                Status.SyncToDB(connectionString);
                for (int i = 0; i < range.Rows.Count; i++)
                {
                    if (Gender.Genders.Find(g => g.Name == range.Cells[i, 2]) == null)
                        Gender.Genders.Add(new Gender(range.Cells[i, 2]));
                    if (Status.Statuses.Find(s => s.Name == range.Cells[i, 4]) == null)
                        Status.Statuses.Add(new Status(range.Cells[i, 4]));
                }
                for (int i = 0; i < range.Rows.Count; i++)
                {
                    Account.Accounts.Add(new Account(range.Cells[i, 0], range.Cells[i, 1], range.Cells[i, 2],
                                                    range.Cells[i, 3], range.Cells[i, 4], range.Cells[i, 5]));
                }
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "INSERT INTO Accounts (ID, Firstname, Secondname, Gender, Age, Status, Salary) VALUES";
                    for (int i = 0; i < Account.Accounts.Count; i++)
                    {
                        Account a = Account.Accounts[i];
                        query += $" ({a.ID}, {a.Firstname}, {a.Secondname}, {a.Gender}, {a.Age}, {a.Status}, {a.Salary})";
                        if (i + 1 == Account.Accounts.Count)
                            query += ";";
                        else
                            query += ",";
                    }
                }
                MessageBox.Show("Успешно выгружено");
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

        }
    }
}
