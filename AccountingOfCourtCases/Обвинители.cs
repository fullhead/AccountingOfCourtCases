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
    public partial class Обвинители : Form
    {
        private SqlConnection sqlConnection = null;
        private PopupNotifier popup = null;
        private SqlDataAdapter adapter = null;
        private DataTable table = null;
        public Обвинители()
        {
            InitializeComponent();
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        }

        private void Обвинители_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //STATUS DB
        private void Обвинители_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Обвинители". При необходимости она может быть перемещена или удалена.
            this.обвинителиTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Обвинители);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Обвинители". При необходимости она может быть перемещена или удалена.
            this.обвинителиTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Обвинители);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Обвинители", sqlConnection);
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
            adapter = new SqlDataAdapter("SELECT * FROM Обвинители", sqlConnection);
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
                adapter = new SqlDataAdapter("SELECT * from Обвинители where ФИО like'%" + textBox1.Text + "%'", sqlConnection);
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

        //SAVE FOR .CSV
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

        //PRINT
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
        //UPDATE
        private async void Button1_Click(object sender, EventArgs e)
        {
            if (фИОTextBox.Text == "" || this.comboBox3.Text == "" || this.comboBox7.Text == "" || this.comboBox5.Text == "" || this.отделение_прокуратурыTextBox.Text == "")
            {
                label4.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Обвинители SET ФИО=@ФИО, Пол=@Пол, Возраст=@Возраст, Адрес=@Адрес, Телефон=@Телефон, Паспортные_данные=@Паспортные_данные, Отделение_прокуратуры=@Отделение_прокуратуры, Примечание=@Примечание WHERE Код_обвинителя=@Код_обвинителя", sqlConnection);
                command.Parameters.AddWithValue("Код_обвинителя", comboBox3.Text);
                command.Parameters.AddWithValue("ФИО", фИОTextBox.Text);
                command.Parameters.AddWithValue("Пол", comboBox5.Text);
                command.Parameters.AddWithValue("Возраст", comboBox7.Text);
                command.Parameters.AddWithValue("Адрес", адресTextBox.Text);
                command.Parameters.AddWithValue("Телефон", телефонTextBox.Text);
                command.Parameters.AddWithValue("Паспортные_данные", паспортные_данныеTextBox.Text);
                command.Parameters.AddWithValue("Отделение_прокуратуры", отделение_прокуратурыTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Обвинители",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
            }
        }

        //Conditions Texboxes for UPDATE
        private void panel4_MouseMove(object sender, MouseEventArgs e)
        {
            label4.Hide();
        }

        private void фИОTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.' && l != ' ')
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

        private void отделение_прокуратурыTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        //INSERT
        private async void button2_Click(object sender, EventArgs e)
        {
            if (фИОTextBox1.Text == "" || this.comboBox1.Text == "" || this.comboBox2.Text == "" || this.отделение_прокуратурыTextBox1.Text == "")
            {
                label5.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Обвинители (ФИО, Пол, Возраст, Адрес, Телефон, Паспортные_данные, Отделение_прокуратуры, Примечание) VALUES (@ФИО, @Пол, @Возраст, @Адрес, @Телефон, @Паспортные_данные, @Отделение_прокуратуры, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("ФИО", фИОTextBox1.Text);
                command.Parameters.AddWithValue("Пол", comboBox2.Text);
                command.Parameters.AddWithValue("Возраст", comboBox1.Text);
                command.Parameters.AddWithValue("Адрес", адресTextBox1.Text);
                command.Parameters.AddWithValue("Телефон", телефонTextBox1.Text);
                command.Parameters.AddWithValue("Паспортные_данные", паспортные_данныеTextBox1.Text);
                command.Parameters.AddWithValue("Отделение_прокуратуры", отделение_прокуратурыTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox1.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Обвинители",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();

                adapter = new SqlDataAdapter("SELECT * FROM Обвинители", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                comboBox3.DataSource = table;
                comboBox4.DataSource = table;
            }
        }
        //Conditions Texboxes for INSERT
        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            label5.Hide();
        }

        private void фИОTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != ' ')
            {
                e.Handled = true;
            }
        }

        private void отделение_прокуратурыTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
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
        private void телефонTextBox1_KeyPress(object sender, KeyPressEventArgs e)
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
            SqlCommand command = new SqlCommand("DELETE FROM Обвинители WHERE Код_обвинителя=@Код_обвинителя", sqlConnection);
            command.Parameters.AddWithValue("Код_обвинителя", comboBox4.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Обвинители",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();

            adapter = new SqlDataAdapter("SELECT * FROM Обвинители", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            comboBox3.DataSource = table;
            comboBox4.DataSource = table;
        }

        private void адвокатыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Адвокаты адвокаты = new Адвокаты();
            адвокаты.Show();
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

        private void судьиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Судьи судьи = new Судьи();
            судьи.Show();
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
