using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using Tulpep.NotificationWindow;
using System.IO;

namespace AccountingOfCourtCases
{
    public partial class Пользователи : Form
    {
        private SqlConnection sqlConnection = null;
        private PopupNotifier popup = null;
        private SqlDataAdapter adapter = null;
        private DataTable table = null;
        public Пользователи()
        {
            InitializeComponent();
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        }

        private void Пользователи_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //STATUS DB
        private void Пользователи_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Пользователи". При необходимости она может быть перемещена или удалена.
            this.пользователиTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Пользователи);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Пользователи", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;

            if (sqlConnection.State == ConnectionState.Open)
                pictureBox1.Image = Properties.Resources.connected;

            else
                pictureBox1.Image = Properties.Resources.disconnect;
        }

        //REFRESH DB
        private void TabControl1_Selected(object sender, TabControlEventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Пользователи", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;
        }

        //SEARCH
        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();

                adapter = new SqlDataAdapter("SELECT * from Пользователи where Логин like'%" + textBox1.Text + "%'", sqlConnection);

                table = new DataTable();

                adapter.Fill(table);
                dataGridView1.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        //UPDATE
        private async void Button1_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.Text == "" || this.логинTextBox.Text == "" || this.парольTextBox.Text == "")
            {
                label4.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlDataAdapter Tablet = new SqlDataAdapter("Select Count (*) Login From Пользователи Where Код_пользователя ='" + comboBox1.Text + "'", sqlConnection);
                DataTable dt = new DataTable();
                Tablet.Fill(dt);
                if (dt.Rows[0][0].ToString() == "1")
                {
                    sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                    sqlConnection.Open();
                    SqlCommand command = new SqlCommand("UPDATE Пользователи SET Логин=@Логин, ФИО=@ФИО, Роль=@Роль, Пароль=@Пароль," +
                        " Email=@Email, Телефон=@Телефон, Примечание=@Примечание WHERE Код_пользователя=@Код_пользователя", sqlConnection);
                    command.Parameters.AddWithValue("Код_пользователя", comboBox1.Text);
                    command.Parameters.AddWithValue("Логин", логинTextBox.Text);
                    command.Parameters.AddWithValue("ФИО", фИОTextBox.Text);
                    command.Parameters.AddWithValue("Роль", Convert.ToBoolean(рольCheckBox.Checked));
                    command.Parameters.AddWithValue("Пароль", парольTextBox.Text);
                    command.Parameters.AddWithValue("Email", emailTextBox.Text);
                    command.Parameters.AddWithValue("Телефон", телефонTextBox.Text);
                    command.Parameters.AddWithValue("Примечание", примечаниеTextBox.Text);
                    popup = new PopupNotifier
                    {
                        Image = Properties.Resources.connected,
                        ImageSize = new Size(96, 96),
                        TitleText = "Пользователи",
                        ContentText = "Данные успешно обновлены!"
                    };
                    popup.Popup();

                    await command.ExecuteNonQueryAsync();
                }
                else
                {
                    label1.Show();
                }
            }
        }

        //Conditions Texboxes for UPDATE
        private void Panel4_MouseMove(object sender, MouseEventArgs e)
        {
            label1.Hide();
            label4.Hide();
        }

        private void ComboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }
        private void ПарольTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        //INSERT
        private async void Button2_Click(object sender, EventArgs e)
        {
            if (this.логинTextBox1.Text == "" || this.парольTextBox1.Text == "")
            {
                label5.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Пользователи (Логин, ФИО, Роль, Пароль, Email, Телефон, Примечание) VALUES (@Логин, @ФИО, @Роль, @Пароль, @Email, @Телефон, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Логин", логинTextBox1.Text);
                command.Parameters.AddWithValue("ФИО", фИОTextBox1.Text);
                command.Parameters.AddWithValue("Роль", Convert.ToBoolean(рольCheckBox1.Checked));
                command.Parameters.AddWithValue("Пароль", парольTextBox1.Text);
                command.Parameters.AddWithValue("Email", emailTextBox1.Text);
                command.Parameters.AddWithValue("Телефон", телефонTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox1.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Пользователи",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
            }
        }

        //Conditions Texboxes for INSERT
        private void Panel5_MouseMove(object sender, MouseEventArgs e)
        {
            label5.Hide();
        }

        //DELETE
        private async void Button3_Click(object sender, EventArgs e)
        {
            if (this.comboBox2.Text == "")
            {
                label6.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlDataAdapter Tablet = new SqlDataAdapter("Select Count (*) Login From Пользователи Where Код_пользователя ='" + comboBox2.Text + "'", sqlConnection);
                DataTable dt = new DataTable();
                Tablet.Fill(dt);
                if (dt.Rows[0][0].ToString() == "1")
                {
                    sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                    sqlConnection.Open();
                    SqlCommand command = new SqlCommand("DELETE FROM Пользователи WHERE Код_пользователя=@Код_пользователя", sqlConnection);
                    command.Parameters.AddWithValue("Код_пользователя", comboBox2.Text);
                    popup = new PopupNotifier
                    {
                        Image = Properties.Resources.connected,
                        ImageSize = new Size(96, 96),
                        TitleText = "Пользователи",
                        ContentText = "Данные успешно удалены!"
                    };
                    popup.Popup();

                    await command.ExecuteNonQueryAsync();
                }
                else
                {
                    label7.Show();
                }
            }
        }

        //Conditions Texboxes for DELETE
        private void Panel6_MouseMove(object sender, MouseEventArgs e)
        {
            label6.Hide();
            label7.Hide();
        }

        private void ComboBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        //SAVE FOR CSV
        private void СохранитьКакCVFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Пользователи", sqlConnection);
                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
            string path = "";
            using (var path_dialog = new FolderBrowserDialog())
                if (path_dialog.ShowDialog() == DialogResult.OK)
                {
                    //Путь к директории
                    path = path_dialog.SelectedPath;
                };
            sqlConnection.Close();
            ToCSV(dt, path + @"\" + @"Отчёт_документы.csv");
        }
        public static void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.UTF8);
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(';'))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        private void ПечатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printDocument1.Print();
        }

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bmp = new Bitmap(dataGridView1.Size.Width + 10, dataGridView1.Size.Height + 10);
            dataGridView1.DrawToBitmap(bmp, dataGridView1.Bounds);
            e.Graphics.DrawImage(bmp, 0, 0);
        }

        //OTHER DB's FORMS
        private void ОПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            О_программе о_программе = new О_программе();
            о_программе.Show();
        }


    }
}
