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
    public partial class Судьи : Form
    {
        private SqlConnection sqlConnection = null;
        private PopupNotifier popup = null;
        private SqlDataAdapter adapter = null;
        private DataTable table = null;
        public Судьи()
        {
            InitializeComponent();
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        }

        private void Судьи_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //STATUS DB
        private void Судьи_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Судьи". При необходимости она может быть перемещена или удалена.
            this.судьиTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Судьи);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Судьи". При необходимости она может быть перемещена или удалена.
            this.судьиTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Судьи);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Судьи", sqlConnection);
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
            adapter = new SqlDataAdapter("SELECT * FROM Судьи", sqlConnection);
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
                adapter = new SqlDataAdapter("SELECT * from Судьи where ФИО like'%" + textBox1.Text + "%'", sqlConnection);
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

        private void СохранитьКакCVFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Судьи", sqlConnection);
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

        //PRINT
        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printDocument1.Print();
        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bmp = new Bitmap(dataGridView1.Size.Width + 10, dataGridView1.Size.Height + 10);
            dataGridView1.DrawToBitmap(bmp, dataGridView1.Bounds);
            e.Graphics.DrawImage(bmp, 0, 0);
        }

        //UPDATE
        private async void button1_Click(object sender, EventArgs e)
        {
            if (фИОTextBox.Text == "" || this.comboBox1.Text == "" || this.comboBox2.Text == "" || this.comboBox3.Text == "")
            {
                label1.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Судьи SET ФИО=@ФИО, Возраст=@Возраст, Пол=@Пол, Телефон=@Телефон, Адрес=@Адрес, Паспортные_данные=@Паспортные_данные, Примечание=@Примечание WHERE Код_судьи=@Код_судьи", sqlConnection);
                command.Parameters.AddWithValue("Код_судьи", comboBox1.Text);
                command.Parameters.AddWithValue("ФИО", фИОTextBox.Text);
                command.Parameters.AddWithValue("Возраст", comboBox2.Text);
                command.Parameters.AddWithValue("Пол", comboBox3.Text);
                command.Parameters.AddWithValue("Телефон", телефонTextBox.Text);
                command.Parameters.AddWithValue("Адрес", адресTextBox.Text);
                command.Parameters.AddWithValue("Паспортные_данные", паспортные_данныеTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Судьи",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
            }
        }

        private void panel4_MouseMove(object sender, MouseEventArgs e)
        {
            label1.Hide();
        }

        private void фИОTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.' && l != ' ')
            {
                e.Handled = true;
            }
        }

        private void телефонTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void адресTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        private void паспортные_данныеTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        //INSERT
        private async void button2_Click(object sender, EventArgs e)
        {
            if (фИОTextBox1.Text == "" || this.comboBox4.Text == "" || this.comboBox5.Text == "")
            {
                label4.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Судьи (ФИО, Возраст, Пол, Телефон, Адрес, Паспортные_данные, Примечание) VALUES (@ФИО, @Возраст, @Пол, @Телефон, @Адрес, @Паспортные_данные, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("ФИО", фИОTextBox1.Text);
                command.Parameters.AddWithValue("Возраст", comboBox4.Text);
                command.Parameters.AddWithValue("Пол", comboBox5.Text);
                command.Parameters.AddWithValue("Телефон", телефонTextBox1.Text);
                command.Parameters.AddWithValue("Адрес", адресTextBox1.Text);
                command.Parameters.AddWithValue("Паспортные_данные", паспортные_данныеTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox1.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Судьи",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();

                adapter = new SqlDataAdapter("SELECT * FROM Судьи", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                comboBox1.DataSource = table;
                comboBox6.DataSource = table;
            }
        }

        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            label4.Hide();
        }

        private void фИОTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.' && l != ' ')
            {
                e.Handled = true;
            }
        }

        private void телефонTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void адресTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        private void паспортные_данныеTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        //DELETE
        private async void button3_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Судьи WHERE Код_судьи=@Код_судьи", sqlConnection);
            command.Parameters.AddWithValue("Код_судьи", comboBox6.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Судьи",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            adapter = new SqlDataAdapter("SELECT * FROM Судьи", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            comboBox1.DataSource = table;
            comboBox6.DataSource = table;
        }

        private void адвокатыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Адвокаты адвокаты = new Адвокаты();
            адвокаты.Show();
        }

        private void обвинителиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Обвинители обвинители = new Обвинители();
            обвинители.Show();
        }

        private void обвиняемыеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Обвиняемые обвиняемые = new Обвиняемые();
            обвиняемые.Show();
        }

        private void пользователиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Пользователи пользователи = new Пользователи();
            пользователи.Show();
        }

        private void статьиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Статьи статьи = new Статьи();
            статьи.Show();
        }

        private void уголовныеДелаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Уголовные_дела уголовные_дела = new Уголовные_дела();
            уголовные_дела.Show();
        }

        private void уликиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Улики улики = new Улики();
            улики.Show();
        }

        private void экспертизыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Экспертизы экспертизы = new Экспертизы();
            экспертизы.Show();
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            О_программе о_программе = new О_программе();
            о_программе.Show();
        }
    }
}
