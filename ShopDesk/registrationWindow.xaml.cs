using System;
using System.Data.OleDb;
using System.Windows;

namespace ShopDesk
{
    public partial class registrationWindow : Window
    {
        // Строка подключения на уровне класса, чтобы была доступна в обработчике
        private readonly string connectingString =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\repos\ShopDesk\ShopDesk\database\employees.accdb";

        public registrationWindow()
        {
            InitializeComponent();
        }

        private void registrationbtn_Click(object sender, RoutedEventArgs e)
        {
            string name = nametxt.Text.Trim();
            string post = posttxt.Text.Trim();
            string login = logintxt.Text.Trim();
            string password = passwordtxt.Text.Trim();

            // Простейшая проверка на пустые поля
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(login) ||
                string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }

            // ВАЖНО: замени Employees и имена столбцов на свои из БД
            string query = "INSERT INTO Employees ([ФИО], [Должность], [Логин], [Пароль]) " +
                           "VALUES (?, ?, ?, ?)";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectingString))
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Для Access (OLE DB) параметры идут в том же порядке, что и в запросе,
                        // и через знак "?", а не через "@name"
                        command.Parameters.AddWithValue("?", name);
                        command.Parameters.AddWithValue("?", post);
                        command.Parameters.AddWithValue("?", login);
                        command.Parameters.AddWithValue("?", password);

                        int rows = command.ExecuteNonQuery();
                        MessageBox.Show($"Добавлено строк: {rows}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении: " + ex.Message);
            }
        }
    }
}