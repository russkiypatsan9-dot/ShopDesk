using System.Data.OleDb;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ShopDesk
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string connectingString =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\repos\ShopDesk\ShopDesk\database\employees.accdb";
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnregistration_Click(object sender, RoutedEventArgs e)
        {
            registrationWindow newWindow = new registrationWindow();
            newWindow.Show();
        }

        private void loginbtn_Click(object sender, RoutedEventArgs e)
        {
            string login = txtlogin.Text.Trim();
            string password = txtpassword.Text.Trim();

            if (string.IsNullOrEmpty(login) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Введите логин и пароль!");
                return;
            }

            string query = "SELECT COUNT(*) FROM employees WHERE [Логин] = ? AND [Пароль] = ?";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectingString))
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("?", login);
                        command.Parameters.AddWithValue("?", password);

                        int count = (int)command.ExecuteScalar();

                        if (count > 0)
                        {
                            MessageBox.Show("Добро пожаловать!");

                            productsWindow products = new productsWindow();
                            products.Show();
                            this.Close();
                        }
                        else
                        {
                            MessageBox.Show("Неверный логин или пароль!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при входе: " + ex.Message);
            }
        }
    }
}