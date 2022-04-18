using System;
using System.Drawing;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using Tulpep.NotificationWindow;

namespace AccountingOfCourtCases
{
    public partial class Регистрация : Form
    {
        private SqlConnection sqlConnection = null;
        private PopupNotifier popup = null;
        public Регистрация()
        {
            InitializeComponent();
        }

        private void Регистрация_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private async void Button1_Click(object sender, EventArgs e)
        {
            if (this.логинTextBox.Text == "" || this.парольTextBox.Text == "")
            {
                label1.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Пользователи (Логин, ФИО, Пароль, Email, Телефон) VALUES (@Логин, @ФИО, @Пароль, @Email, @Телефон)", sqlConnection);
                command.Parameters.AddWithValue("Логин", логинTextBox.Text);
                command.Parameters.AddWithValue("ФИО", фИОTextBox.Text);
                command.Parameters.AddWithValue("Пароль", парольTextBox.Text);
                command.Parameters.AddWithValue("Email", emailTextBox.Text);
                command.Parameters.AddWithValue("Телефон", телефонTextBox.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Регистрация",
                    ContentText = "Регистрация успешно завершина! Теперь, можете войти."
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                this.Hide();
                Вход вход = new Вход();
                вход.Show();
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Вход вход = new Вход();
            вход.Show();
        }

        private void Panel1_MouseMove(object sender, MouseEventArgs e)
        {
            label1.Hide();
        }

        private void ПарольTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }
    }
}
