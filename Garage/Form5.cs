using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Word = Microsoft.Office.Interop.Word;

namespace Garage
{
    public partial class Form5 : Form
    {
        string repair = "repair";
        string temp = "";

        public Form5()
        {
            InitializeComponent();
            InitializeInterface();
            DataGridREView(repair);
            GetNoteCount(repair);
        }

        private void InitializeInterface()
        {
            this.Text = "Просмотр и удаление заказов";
            button1.Text = "Назад";
            button2.Text = "Удалить";
            button3.Text = "";
            button3.Text = "Отчет";
            label1.Text = "Кол-во записей: ";
            button2.Enabled = false;
            button3.Enabled = false;

            GetNoteCount(repair);
        }

        private void GetNoteCount(string table)
        {
            try
            {
                string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection oleDbConn = new OleDbConnection(con);
                oleDbConn.Open();
                OleDbCommand sql = new OleDbCommand("SELECT COUNT(*) FROM " + table + ";", oleDbConn);
                sql.Connection = oleDbConn;
                sql.ExecuteNonQuery();

                label1.Text = "Кол-во записей: " + (int)sql.ExecuteScalar();

                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void DataGridREView(string table)
        {
            try
            {
                string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection oleDbConn = new OleDbConnection(con);
                DataTable dt = new DataTable();
                oleDbConn.Open();
                OleDbCommand sql = new OleDbCommand("SELECT repair_id AS [id], (SELECT mechanic_surname FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [фамилия мастера], (SELECT mechanic_name FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [имя мастера], (SELECT mechanic_patronymic FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [отчество], (SELECT car_name FROM car WHERE repair.car_id = car.car_id) AS [модель авто], (SELECT car_mark FROM car WHERE repair.car_id = car.car_id) AS [марка авто], repair_date AS [дата работы], repair_cost AS [цена] FROM " + table + ";", oleDbConn);
                sql.Connection = oleDbConn;
                sql.ExecuteNonQuery();

                OleDbDataAdapter da = new OleDbDataAdapter(sql);
                da.Fill(dt);

                dataGridView1.DataSource = dt;

                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form MainForm = new Form1();
            MainForm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно желаете удалить запись?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                DeleteData(repair);
                DataGridREView(repair);
                GetNoteCount(repair);
                MessageBox.Show("Запись удалена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            button2.Enabled = false;
            button3.Enabled = false;
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            button2.Enabled = true;
            button3.Enabled = true;

            try
            {
                temp = dataGridView1.SelectedCells[0].Value.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void DeleteData(string table)
        {
            try
            {
                string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection oleDbConn = new OleDbConnection(con);
                DataTable dt = new DataTable();
                oleDbConn.Open();

                OleDbCommand sql = new OleDbCommand("DELETE FROM " + table + " WHERE repair_id=" + Convert.ToInt32(temp) + ";");
                sql.Connection = oleDbConn;
                sql.ExecuteNonQuery();

                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void WordExport()
        {
            string con1 = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
            OleDbConnection OleDbcon1 = new OleDbConnection(con1);
            OleDbcon1.Open();
            OleDbCommand sql1 = new OleDbCommand("SELECT repair_id AS [id], (SELECT mechanic_surname FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [фамилия мастера], (SELECT mechanic_name FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [имя мастера], (SELECT mechanic_patronymic FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [отчество], (SELECT car_name FROM car WHERE repair.car_id = car.car_id) AS [модель авто], (SELECT car_mark FROM car WHERE repair.car_id = car.car_id) AS [марка авто], repair_date AS [дата работы], repair_cost AS [цена] FROM repair WHERE repair_id = " + temp +";", OleDbcon1);
            sql1.ExecuteNonQuery();
            OleDbDataReader reader = sql1.ExecuteReader();
            while (reader.Read())
            {
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Add();
            wordApp.Visible = true;

            Word.Range range = doc.Range();
            doc.SaveAs();
                range.Text = "Номер сделки: "+ reader["id"].ToString() +
                "\nФамилия мастера: " + reader["фамилия мастера"].ToString() +
                "\nИмя мастера: " + reader["имя мастера"].ToString() +
                "\nОтчество: " + reader["отчество"].ToString() +
                "\nМодель авто: " + reader["модель авто"].ToString() +
                "\nМарка авто: " + reader["марка авто"].ToString() +
                "\nДата работы: " + reader["дата работы"].ToString() +
                "\nЦена: " + reader["цена"].ToString() + " руб";
            }
            reader.Close();
            OleDbcon1.Close();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                WordExport();
                button2.Enabled = false;
                button3.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
           
        }
    }
}
