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
    public partial class Уголовные_дела : Form
    {
        private SqlConnection sqlConnection = null;
        private PopupNotifier popup = null;
        private SqlDataAdapter adapter = null;
        private DataTable table = null;
        public Уголовные_дела()
        {
            InitializeComponent();
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        }

        private void Уголовные_дела_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //STATUS DB
        private void Уголовные_дела_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Экспертизы". При необходимости она может быть перемещена или удалена.
            this.экспертизыTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Экспертизы);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Обвинители". При необходимости она может быть перемещена или удалена.
            this.обвинителиTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Обвинители);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Адвокаты". При необходимости она может быть перемещена или удалена.
            this.адвокатыTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Адвокаты);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Обвиняемые". При необходимости она может быть перемещена или удалена.
            this.обвиняемыеTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Обвиняемые);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Судьи". При необходимости она может быть перемещена или удалена.
            this.судьиTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Судьи);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Уголовные_дела". При необходимости она может быть перемещена или удалена.
            this.уголовные_делаTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Уголовные_дела);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Уголовные_дела". При необходимости она может быть перемещена или удалена.
            this.уголовные_делаTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Уголовные_дела);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Уголовные_дела". При необходимости она может быть перемещена или удалена.
            this.уголовные_делаTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Уголовные_дела);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Уголовные_дела", sqlConnection);
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
            adapter = new SqlDataAdapter("SELECT * FROM Уголовные_дела", sqlConnection);
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
                adapter = new SqlDataAdapter("SELECT * from Уголовные_дела where Решение_суда like'%" + textBox1.Text + "%'", sqlConnection);
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
                adapter = new SqlDataAdapter("SELECT * FROM Уголовные_дела", sqlConnection);
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
            if (решение_судаTextBox.Text == "")
            {
                label1.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Уголовные_дела SET Код_судьи=@Код_судьи, Код_обвиняемого=@Код_обвиняемого, " +
                    "Код_адвоката=@Код_адвоката, Код_обвинителя=@Код_обвинителя, Код_экспертизы=@Код_экспертизы, Решение_суда=@Решение_суда, " +
                    "Примечание=@Примечание WHERE Код_уголовного_дела=@Код_уголовного_дела", sqlConnection);
                command.Parameters.AddWithValue("Код_уголовного_дела", comboBox1.Text);
                command.Parameters.AddWithValue("Код_судьи", comboBox2.Text);
                command.Parameters.AddWithValue("Код_обвиняемого", comboBox3.Text);
                command.Parameters.AddWithValue("Код_адвоката", comboBox4.Text);
                command.Parameters.AddWithValue("Код_обвинителя", comboBox5.Text);
                command.Parameters.AddWithValue("Код_экспертизы", comboBox6.Text);
                command.Parameters.AddWithValue("Решение_суда", решение_судаTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Уголовные_дела",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
            }
        }

        private void решение_судаTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        private void panel4_MouseMove(object sender, MouseEventArgs e)
        {
            label1.Hide();
        }

        //INSERT
        private async void button2_Click(object sender, EventArgs e)
        {
            if (решение_судаTextBox1.Text == "")
            {
                label4.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Уголовные_дела (Код_судьи, Код_обвиняемого, Код_адвоката, Код_обвинителя, Код_экспертизы, Решение_суда, Примечание) VALUES (@Код_судьи, @Код_обвиняемого, @Код_адвоката, @Код_обвинителя, @Код_экспертизы, @Решение_суда, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Код_судьи", comboBox7.Text);
                command.Parameters.AddWithValue("Код_обвиняемого", comboBox8.Text);
                command.Parameters.AddWithValue("Код_адвоката", comboBox9.Text);
                command.Parameters.AddWithValue("Код_обвинителя", comboBox10.Text);
                command.Parameters.AddWithValue("Код_экспертизы", comboBox11.Text);
                command.Parameters.AddWithValue("Решение_суда", решение_судаTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox1.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Уголовные дела",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();

                adapter = new SqlDataAdapter("SELECT * FROM Уголовные_дела", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                comboBox1.DataSource = table;
                comboBox12.DataSource = table;
            }

        }

        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            label4.Hide();
        }

        private void решение_судаTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        //DELETE
        private async void button3_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Уголовные_дела WHERE Код_уголовного_дела=@Код_уголовного_дела", sqlConnection);
            command.Parameters.AddWithValue("Код_уголовного_дела", comboBox12.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Уголовные дела",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            adapter = new SqlDataAdapter("SELECT * FROM Уголовные_дела", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            comboBox1.DataSource = table;
            comboBox12.DataSource = table;
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

        private void судьиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Судьи судьи = new Судьи();
            судьи.Show();
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
