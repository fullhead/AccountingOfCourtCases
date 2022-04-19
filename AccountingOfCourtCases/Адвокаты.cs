﻿using System;
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
    public partial class Адвокаты : Form
    {
        private SqlConnection sqlConnection = null;
        private PopupNotifier popup = null;
        private SqlDataAdapter adapter = null;
        private DataTable table = null;
        public Адвокаты()
        {
            InitializeComponent();
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

        }

        private void Адвокаты_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //STATUS DB
        private void Адвокаты_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Обвиняемые". При необходимости она может быть перемещена или удалена.
            this.обвиняемыеTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Обвиняемые);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfCourtCasesDataSet.Адвокаты". При необходимости она может быть перемещена или удалена.
            this.адвокатыTableAdapter.Fill(this.accountingOfCourtCasesDataSet.Адвокаты);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Адвокаты", sqlConnection);
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
            adapter = new SqlDataAdapter("SELECT * FROM Адвокаты", sqlConnection);
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
                adapter = new SqlDataAdapter("SELECT * from Адвокаты where ФИО like'%" + textBox1.Text + "%'", sqlConnection);
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

        //SAVE FOR CSV
        private void СохранитьКакCVFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Адвокаты", sqlConnection);
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
            if (comboBox1.Text == "" || this.comboBox2.Text == "" || this.фИОTextBox.Text == "" || this.comboBox3.Text == "" || this.comboBox7.Text == "" || this.компанияTextBox.Text == "")
            {
                label4.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlDataAdapter Tablet = new SqlDataAdapter("Select Count (*) Login From Адвокаты Where Код_адвоката ='" + comboBox1.Text + "'", sqlConnection);
                DataTable dt = new DataTable();
                Tablet.Fill(dt);
                if (dt.Rows[0][0].ToString() == "1")
                {
                    sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                    sqlConnection.Open();
                    SqlDataAdapter Tablet1 = new SqlDataAdapter("Select Count (*) Login From Обвиняемые Where Код_обвиняемого ='" + comboBox2.Text + "'", sqlConnection);
                    DataTable dt1 = new DataTable();
                    Tablet1.Fill(dt1);
                    if (dt1.Rows[0][0].ToString() == "1")
                    {
                        sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                        sqlConnection.Open();
                        SqlCommand command = new SqlCommand("UPDATE Адвокаты SET Код_обвиняемого=@Код_обвиняемого, ФИО=@ФИО, Пол=@Пол," +
                        " Возраст=@Возраст, Телефон=@Телефон, Адрес=@Адрес, Паспортные_данные=@Паспортные_данные, Компания=@Компания, Примечание=@Примечание WHERE Код_адвоката=@Код_адвоката", sqlConnection);
                        command.Parameters.AddWithValue("Код_адвоката", comboBox1.Text);
                        command.Parameters.AddWithValue("Код_обвиняемого", comboBox2.Text);
                        command.Parameters.AddWithValue("ФИО", фИОTextBox.Text);
                        command.Parameters.AddWithValue("Пол", comboBox3.Text);
                        command.Parameters.AddWithValue("Возраст", comboBox7.Text);
                        command.Parameters.AddWithValue("Телефон", телефонTextBox.Text);
                        command.Parameters.AddWithValue("Адрес", адресTextBox.Text);
                        command.Parameters.AddWithValue("Паспортные_данные", паспортные_данныеTextBox.Text);
                        command.Parameters.AddWithValue("Компания", компанияTextBox.Text);
                        command.Parameters.AddWithValue("Примечание", примечаниеTextBox.Text);
                        popup = new PopupNotifier
                        {
                            Image = Properties.Resources.connected,
                            ImageSize = new Size(96, 96),
                            TitleText = "Адвокаты",
                            ContentText = "Данные успешно обновлены!"
                        };
                        popup.Popup();
                        await command.ExecuteNonQueryAsync();
                    }
                    else
                    {
                        label8.Show();
                    }
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
            label8.Hide();
        }
        private void ФИОTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.')
            {
                e.Handled = true;
            }
        }

        private void КомпанияTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.')
            {
                e.Handled = true;
            }
        }

        //INSERT
        private async void Button2_Click(object sender, EventArgs e)
        {
            if (comboBox4.Text == "" || this.фИОTextBox1.Text == "" || this.comboBox5.Text == "" || this.comboBox8.Text == "" || this.компанияTextBox1.Text == "")
            {
                label5.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlDataAdapter Tablet1 = new SqlDataAdapter("Select Count (*) Login From Обвиняемые Where Код_обвиняемого ='" + comboBox4.Text + "'", sqlConnection);
                DataTable dt1 = new DataTable();
                Tablet1.Fill(dt1);
                if (dt1.Rows[0][0].ToString() == "1")
                {
                    sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                    sqlConnection.Open();
                    SqlCommand command = new SqlCommand("INSERT INTO Адвокаты (Код_обвиняемого, ФИО, Пол, Возраст, Телефон, Адрес, Паспортные_данные, Компания, Примечание) VALUES (@Код_обвиняемого, @ФИО, @Пол, @Возраст, @Телефон, @Адрес, @Паспортные_данные, @Компания, @Примечание)", sqlConnection);
                    command.Parameters.AddWithValue("Код_обвиняемого", comboBox4.Text);
                    command.Parameters.AddWithValue("ФИО", фИОTextBox1.Text);
                    command.Parameters.AddWithValue("Пол", comboBox5.Text);
                    command.Parameters.AddWithValue("Возраст", comboBox8.Text);
                    command.Parameters.AddWithValue("Телефон", телефонTextBox1.Text);
                    command.Parameters.AddWithValue("Адрес", адресTextBox1.Text);
                    command.Parameters.AddWithValue("Паспортные_данные", паспортные_данныеTextBox1.Text);
                    command.Parameters.AddWithValue("Компания", компанияTextBox1.Text);
                    command.Parameters.AddWithValue("Примечание", примечаниеTextBox1.Text);
                    popup = new PopupNotifier
                    {
                        Image = Properties.Resources.connected,
                        ImageSize = new Size(96, 96),
                        TitleText = "Адвокаты",
                        ContentText = "Данные успешно добавлены!"
                    };
                    popup.Popup();
                    await command.ExecuteNonQueryAsync();
                }
                else
                {
                    label9.Show();
                }

            }
        }

        //Conditions Texboxes for INSERT
        private void Panel5_MouseMove(object sender, MouseEventArgs e)
        {
            label5.Hide();
            label9.Hide();
        }
        private void ФИОTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.')
            {
                e.Handled = true;
            }
        }

        private void КомпанияTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != '.')
            {
                e.Handled = true;
            }
        }

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
                SqlDataAdapter Tablet = new SqlDataAdapter("Select Count (*) Login From Адвокаты Where Код_адвоката ='" + comboBox6.Text + "'", sqlConnection);
                DataTable dt = new DataTable();
                Tablet.Fill(dt);
                if (dt.Rows[0][0].ToString() == "1")
                {
                    sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                    sqlConnection.Open();
                    SqlCommand command = new SqlCommand("DELETE FROM Адвокаты WHERE Код_адвоката=@Код_адвоката", sqlConnection);
                    command.Parameters.AddWithValue("Код_адвоката", comboBox2.Text);
                    popup = new PopupNotifier
                    {
                        Image = Properties.Resources.connected,
                        ImageSize = new Size(96, 96),
                        TitleText = "Адвокаты",
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
    }
}
