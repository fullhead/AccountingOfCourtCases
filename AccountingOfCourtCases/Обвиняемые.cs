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
    public partial class Обвиняемые : Form
    {
        private SqlConnection sqlConnection = null;
        private PopupNotifier popup = null;
        private SqlDataAdapter adapter = null;
        private DataTable table = null;
        public Обвиняемые()
        {
            InitializeComponent();
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        }

        private void Обвиняемые_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //STATUS DB
        private void Обвиняемые_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Статьи". При необходимости она может быть перемещена или удалена.
            this.статьиTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Статьи);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Обвиняемые". При необходимости она может быть перемещена или удалена.
            this.обвиняемыеTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Обвиняемые);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Статьи". При необходимости она может быть перемещена или удалена.
            this.статьиTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Статьи);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Обвиняемые". При необходимости она может быть перемещена или удалена.
            this.обвиняемыеTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Обвиняемые);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Обвиняемые", sqlConnection);
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
            adapter = new SqlDataAdapter("SELECT * FROM Обвиняемые", sqlConnection);
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
                adapter = new SqlDataAdapter("SELECT * from Обвиняемые where ФИО like'%" + textBox1.Text + "%'", sqlConnection);
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
                adapter = new SqlDataAdapter("SELECT * FROM Обвиняемые", sqlConnection);
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
            if (фИОTextBox.Text == "" || this.comboBox1.Text == "" || this.comboBox2.Text == "" || this.comboBox5.Text == "" || this.comboBox7.Text == "")
            {
                label4.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Обвиняемые SET Код_статьи=@Код_статьи, ФИО=@ФИО, Пол=@Пол," +
                " Возраст=@Возраст, Адрес=@Адрес, Телефон=@Телефон, Паспортные_данные=@Паспортные_данные, Примечание=@Примечание WHERE Код_обвиняемого=@Код_обвиняемого", sqlConnection);
                command.Parameters.AddWithValue("Код_обвиняемого", comboBox2.Text);
                command.Parameters.AddWithValue("Код_статьи", comboBox1.Text);
                command.Parameters.AddWithValue("ФИО", фИОTextBox.Text);
                command.Parameters.AddWithValue("Пол", comboBox5.Text);
                command.Parameters.AddWithValue("Возраст", comboBox7.Text);
                command.Parameters.AddWithValue("Телефон", телефонTextBox.Text);
                command.Parameters.AddWithValue("Адрес", адресTextBox.Text);
                command.Parameters.AddWithValue("Паспортные_данные", паспортные_данныеTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Обвиняемые",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
            }
        }
        //Conditions Texboxes for UPDATE
        private void ФИОTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.')
            {
                e.Handled = true;
            }
        }
        private void Panel4_MouseMove(object sender, MouseEventArgs e)
        {
            label4.Hide();
        }
        //INSERT
        private async void Button2_Click(object sender, EventArgs e)
        {
            if (фИОTextBox1.Text == "" || this.comboBox3.Text == "" || this.comboBox6.Text == "" || this.comboBox8.Text == "")
            {
                label5.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Обвиняемые (Код_статьи, ФИО, Пол, Возраст, Адрес, Телефон, Паспортные_данные, Примечание) VALUES (@Код_статьи, @ФИО, @Пол, @Возраст, @Адрес, @Телефон, @Паспортные_данные, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Код_статьи", comboBox3.Text);
                command.Parameters.AddWithValue("ФИО", фИОTextBox1.Text);
                command.Parameters.AddWithValue("Пол", comboBox8.Text);
                command.Parameters.AddWithValue("Возраст", comboBox6.Text);
                command.Parameters.AddWithValue("Телефон", телефонTextBox1.Text);
                command.Parameters.AddWithValue("Адрес", адресTextBox1.Text);
                command.Parameters.AddWithValue("Паспортные_данные", паспортные_данныеTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox1.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Обвиняемые",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();

                adapter = new SqlDataAdapter("SELECT * FROM Обвиняемые", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                comboBox2.DataSource = table;
                comboBox4.DataSource = table;
            }
        }

        //Conditions Texboxes for INSERT
        private void ФИОTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.' && l != ' ')
            {
                e.Handled = true;
            }
        }
        private void Panel5_MouseMove(object sender, MouseEventArgs e)
        {
            label5.Hide();
        }

        //DELETE
        private async void Button3_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Обвиняемые WHERE Код_обвиняемого=@Код_обвиняемого", sqlConnection);
            command.Parameters.AddWithValue("Код_обвиняемого", comboBox2.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Обвиняемые",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();

            adapter = new SqlDataAdapter("SELECT * FROM Обвиняемые", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            comboBox2.DataSource = table;
            comboBox4.DataSource = table;

        }
        //OTHER DB's FORMS
        private void АдвокатыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Адвокаты адвокаты = new Адвокаты();
            адвокаты.Show();
        }

        private void ОбвинителиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Обвинители обвинители = new Обвинители();
            обвинители.Show();
        }

        private void ПользователиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Пользователи пользователи = new Пользователи();
            пользователи.Show();
        }

        private void СтатьиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Статьи статьи = new Статьи();
            статьи.Show();
        }

        private void СудьиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Судьи судьи = new Судьи();
            судьи.Show();
        }

        private void УголовныеДелаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Уголовные_дела уголовные_дела = new Уголовные_дела();
            уголовные_дела.Show();
        }

        private void УликиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Улики улики = new Улики();
            улики.Show();
        }

        private void ЭкспертизыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Экспертизы экспертизы = new Экспертизы();
            экспертизы.Show();
        }

        private void ОПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            О_программе о_программе = new О_программе();
            о_программе.Show();
        }
    }
}
