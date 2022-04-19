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
        private void button1_Click(object sender, EventArgs e)
        {

        }


    }
}
