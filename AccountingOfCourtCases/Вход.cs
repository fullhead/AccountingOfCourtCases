using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace AccountingOfCourtCases
{
    public partial class Вход : Form
    {
        private SqlConnection sqlConnection;
        public Вход()
        {
            InitializeComponent();
        }

        private void Вход_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlDataAdapter Tablet = new SqlDataAdapter("Select Count (*) Login From Пользователи Where Логин ='" + textBox1.Text + "'and Пароль = '" + textBox2.Text + "'", sqlConnection);
            DataTable dt = new DataTable();
            Tablet.Fill(dt);
            if (dt.Rows[0][0].ToString() == "1")

            {
                Пользователи пользователи = new Пользователи();
                пользователи.Show();
                this.Hide();

            }
            else
            {
                label3.Show();
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Регистрация регистрация = new Регистрация();
            регистрация.Show();
        }

        private void TextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void Panel1_MouseMove(object sender, MouseEventArgs e)
        {
            label3.Hide();
        }
    }
}
