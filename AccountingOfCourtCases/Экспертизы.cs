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
    public partial class Экспертизы : Form
    {
        private SqlConnection sqlConnection = null;
        private PopupNotifier popup = null;
        private SqlDataAdapter adapter = null;
        private DataTable table = null;
        public Экспертизы()
        {
            InitializeComponent();
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        }

        private void Экспертизы_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //STATUS DB
        private void Экспертизы_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Экспертизы". При необходимости она может быть перемещена или удалена.
            this.экспертизыTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Экспертизы);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Экспертизы". При необходимости она может быть перемещена или удалена.
            this.экспертизыTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Экспертизы);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Экспертизы", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;

            if (sqlConnection.State == ConnectionState.Open)
                pictureBox1.Image = Properties.Resources.connected;

            else
                pictureBox1.Image = Properties.Resources.disconnect;
        }

        //REFRESH DB
        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Экспертизы", sqlConnection);
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
                adapter = new SqlDataAdapter("SELECT * from Экспертизы where Заключение like'%" + textBox1.Text + "%'", sqlConnection);
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
                adapter = new SqlDataAdapter("SELECT * FROM Экспертизы", sqlConnection);
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
            if (заключениеTextBox.Text == "" || this.comboBox1.Text == "")
            {
                label1.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Экспертизы SET Заключение=@Заключение, Примечание=@Примечание WHERE Код_экспертизы=@Код_экспертизы", sqlConnection);
                command.Parameters.AddWithValue("Код_экспертизы", comboBox1.Text);
                command.Parameters.AddWithValue("Заключение", заключениеTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Экспертизы",
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

        private void заключениеTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        //INSERT
        private async void button2_Click(object sender, EventArgs e)
        {
            if (заключениеTextBox1.Text == "")
            {
                label4.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Экспертизы (Заключение, Примечание) VALUES (@Заключение, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Заключение", заключениеTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox1.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Экспертизы",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();

                adapter = new SqlDataAdapter("SELECT * FROM Экспертизы", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                comboBox1.DataSource = table;
                comboBox2.DataSource = table;
            }
        }

        private void заключениеTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }
        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            label4.Hide();
        }
        private async void button3_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Экспертизы WHERE Код_экспертизы=@Код_экспертизы", sqlConnection);
            command.Parameters.AddWithValue("Код_экспертизы", comboBox2.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Экспертизы",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            adapter = new SqlDataAdapter("SELECT * FROM Экспертизы", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            comboBox1.DataSource = table;
            comboBox2.DataSource = table;
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
        private void судьиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Судьи судьи = new Судьи();
            судьи.Show();
        }
        private void угловныеДелаToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            О_программе о_программе = new О_программе();
            о_программе.Show();
        }


    }
}
