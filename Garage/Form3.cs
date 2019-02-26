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

namespace Garage
{
    public partial class Form3 : Form
    {

        string car = "car";
        string mechanic = "mechanic";
        string temp = "";

        public Form3()
        {
            InitializeComponent();
            InitializeInterface();
            DataGridREView(car);
            DataGridREView(mechanic);
        }

        private void InitializeInterface()
        {
            this.Text = "Редактирование и удаление данных";
            tabPage1.Text = "Механики";
            tabPage2.Text = "Машины";

            button1.Text = "Назад";
            button2.Text = "Редактировать";
            button3.Text = "Очистить поля ввода";
            button8.Text = "Удалить";
            label1.Text = "Номер механика";
            label2.Text = "Фамилия";
            label3.Text = "Имя";
            label4.Text = "Отчество";
            label5.Text = "Стаж, года";
            label6.Text = "Разряд";
            label12.Text = "месяца";
            label13.Text = "дни";
            string[] ranks = { "1", "2", "3", "4", "5", "6" };
            comboBox1.Items.AddRange(ranks);

            button6.Text = "Назад";
            button5.Text = "Редактировать";
            button4.Text = "Очистить поля ввода";
            button7.Text = "Удалить";
            label7.Text = "Номер авто";
            label8.Text = "Марка";
            label9.Text = "Модель";
            label10.Text = "Год выпуска";
            label11.Text = "Тип кузова";
            string[] car_types = { "Седан", "Хэтчбэк", "Универсал", "Лифтбэк", "Купе", "Кабриолет", "Родстер", "Тарга" };
            comboBox2.Items.AddRange(car_types);

            button2.Enabled = false;
            button8.Enabled = false;
            button5.Enabled = false;
            button7.Enabled = false;

            GetNoteCount(car);
            GetNoteCount(mechanic);
        }

        private void GetNoteCount(string table)
        {
            if (table == "mechanic")
            {
                try
                {
                    string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                    OleDbConnection oleDbConn = new OleDbConnection(con);
                    oleDbConn.Open();
                    OleDbCommand sql = new OleDbCommand("SELECT COUNT(*) FROM " + table + ";", oleDbConn);
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();

                    label14.Text = "Кол-во записей: " + (int)sql.ExecuteScalar();

                    oleDbConn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            if (table == "car")
            {
                try
                {
                    string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                    OleDbConnection oleDbConn = new OleDbConnection(con);
                    oleDbConn.Open();
                    OleDbCommand sql = new OleDbCommand("SELECT COUNT(*) FROM " + table + ";", oleDbConn);
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();

                    label15.Text = "Кол-во записей: " + (int)sql.ExecuteScalar();

                    oleDbConn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void ClearTextBox(string table)
        {
            if (table == "mechanic")
            {
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox10.Text = "";
                textBox11.Text = "";
                comboBox1.Text = "";
            }
            else if (table == "car")
            {
                textBox6.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                comboBox2.Text = "";
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
                if (table == "mechanic")
                {
                    OleDbCommand sql = new OleDbCommand("DELETE FROM " + table + " WHERE mechanic_id=" + temp + ";");
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();
                }
                else if (table == "car")
                {
                    OleDbCommand sql = new OleDbCommand("DELETE FROM " + table + " WHERE car_id=" + temp + ";");
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();
                }

                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            GetNoteCount(car);
            GetNoteCount(mechanic);
        }

        private string[] GetExpArray(string exp)
        {
            string[] ExpArray = exp.Split(new Char[] { ' ' });
            return ExpArray;
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                temp = dataGridView1.SelectedCells[0].Value.ToString();
                textBox1.Text = dataGridView1.SelectedCells[1].Value.ToString();
                textBox2.Text = dataGridView1.SelectedCells[2].Value.ToString();
                textBox3.Text = dataGridView1.SelectedCells[3].Value.ToString();
                textBox4.Text = dataGridView1.SelectedCells[4].Value.ToString();
                textBox5.Text = GetExpArray(dataGridView1.SelectedCells[5].Value.ToString())[0];
                textBox10.Text = GetExpArray(dataGridView1.SelectedCells[5].Value.ToString())[1];
                textBox11.Text = GetExpArray(dataGridView1.SelectedCells[5].Value.ToString())[2];
                comboBox1.Text = dataGridView1.SelectedCells[6].Value.ToString();
                button2.Enabled = true;
                button8.Enabled = true;
            }
            catch (Exception)
            {

            }
        }

        private void dataGridView2_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                temp = dataGridView2.SelectedCells[0].Value.ToString();
                textBox6.Text = dataGridView2.SelectedCells[1].Value.ToString();
                textBox7.Text = dataGridView2.SelectedCells[2].Value.ToString();
                textBox8.Text = dataGridView2.SelectedCells[3].Value.ToString();
                comboBox2.Text = dataGridView2.SelectedCells[4].Value.ToString();
                textBox9.Text = dataGridView2.SelectedCells[5].Value.ToString();
                button5.Enabled = true;
                button7.Enabled = true;
            }
            catch (Exception)
            {

            }
        }

        private string GetMechanicExp()
        {
            string time = textBox5.Text + " " + textBox10.Text + " " + textBox11.Text;
            return time;
        }

        private void UpdateData(string table)
        {
            try
            {
                string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection oleDbConn = new OleDbConnection(con);
                DataTable dt = new DataTable();
                oleDbConn.Open();
                if (table == "mechanic")
                {
                    OleDbCommand sql = new OleDbCommand("UPDATE " + table + " SET mechanic_number='" + textBox1.Text + "', mechanic_surname='" + textBox2.Text + "', mechanic_name='" + textBox3.Text + "' , mechanic_patronymic='" + textBox4.Text + "' , mechanic_exp= '" + GetMechanicExp() + "', mechanic_rank='" + comboBox1.Text + "' WHERE mechanic_id=" + temp + ";");
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();
                }
                else if (table == "car")
                {
                    OleDbCommand sql = new OleDbCommand("UPDATE " + table + " SET car_number='" + textBox6.Text + "',  car_mark='" + textBox7.Text + "', car_name='" + textBox8.Text + "' , car_type='" + comboBox2.Text + "', car_year=" + Convert.ToInt32(textBox9.Text) + "  WHERE car_id=" + temp + ";");
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();
                }
                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            GetNoteCount(car);
            GetNoteCount(mechanic);
        }

        private void DataGridREView(string table)
        {
            try
            {
                string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection oleDbConn = new OleDbConnection(con);
                DataTable dt = new DataTable();
                oleDbConn.Open();
                OleDbCommand sql = new OleDbCommand("SELECT * FROM " + table + ";");
                sql.Connection = oleDbConn;
                sql.ExecuteNonQuery();

                OleDbDataAdapter da = new OleDbDataAdapter(sql);
                da.Fill(dt);

                if (table == "car")
                {
                    dt.Columns["car_number"].ColumnName = "номер авто";
                    dt.Columns["car_mark"].ColumnName = "марка";
                    dt.Columns["car_name"].ColumnName = "модель";
                    dt.Columns["car_type"].ColumnName = "тип";
                    dt.Columns["car_year"].ColumnName = "год выпуска";
                    dataGridView2.DataSource = dt;
                }

                if (table == "mechanic")
                {
                    dt.Columns["mechanic_number"].ColumnName = "номер";
                    dt.Columns["mechanic_surname"].ColumnName = "фамилия";
                    dt.Columns["mechanic_name"].ColumnName = "имя";
                    dt.Columns["mechanic_patronymic"].ColumnName = "отчество";
                    dt.Columns["mechanic_exp"].ColumnName = "стаж";
                    dt.Columns["mechanic_rank"].ColumnName = "разряд";
                    dataGridView1.DataSource = dt;
                }

                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void GoBackToMainForm()
        {
            this.Close();
            Form MainForm = new Form1();
            MainForm.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GoBackToMainForm();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            GoBackToMainForm();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно очистить поля ввода?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                ClearTextBox(mechanic);
                MessageBox.Show("Поля очищены", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно очистить поля ввода?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                ClearTextBox(car);
                MessageBox.Show("Поля очищены", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var result = TextCheck(mechanic);
            if (result == 1)
            {
                return;
            }
            else if (result == 0)
            {
                UpdateData(mechanic);
                ClearTextBox(mechanic);
                DataGridREView(mechanic);
                MessageBox.Show("Запись изменена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            button2.Enabled = false;
            button8.Enabled = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно желаете удалить запись?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                DeleteData(mechanic);
                ClearTextBox(mechanic);
                DataGridREView(mechanic);
                MessageBox.Show("Запись удалена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            button2.Enabled = false;
            button8.Enabled = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var result = TextCheck(car);
            if (result == 1)
            {
                return;
            }
            else if (result == 0)
            {
                UpdateData(car);
                ClearTextBox(car);
                DataGridREView(car);
                MessageBox.Show("Запись изменена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            button5.Enabled = false;
            button7.Enabled = false;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно желаете удалить запись?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                DeleteData(car);
                ClearTextBox(car);
                DataGridREView(car);
                MessageBox.Show("Запись удалена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            button5.Enabled = false;
            button7.Enabled = false;
        }

        private void FullNoInput(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;

            if (!Char.IsDigit(ch) && ch != 9)
            {
                e.Handled = true;
            }

            if (!Char.IsDigit(e.KeyChar)) return;
            else
                e.Handled = true;
        }

        private void TextNoInput(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 9)
            {
                e.Handled = true;
            }
        }

        private void NumberNoInput(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar)) return;
            else
                e.Handled = true;
        }

        public int TextCheck(string table)
        {
            if (table == "mechanic")
            {
                if (textBox1.Text == "")
                {
                    MessageBox.Show("Заполните корректно поле Номер механика!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (textBox2.Text == "")
                {
                    MessageBox.Show("Заполните корректно поле Фамилия!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (textBox3.Text == "")
                {
                    MessageBox.Show("Заполните корректно поле Имя!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (textBox4.Text == "")
                {
                    MessageBox.Show("Заполните корректно поле Отчество!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (textBox5.Text == "" || Convert.ToInt32(textBox5.Text) > 120 || textBox10.Text == "" || Convert.ToInt32(textBox10.Text) > 12 || textBox11.Text == "" || Convert.ToInt32(textBox11.Text) > 31)
                {
                    MessageBox.Show("Заполните корректно поле Стаж!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (comboBox1.Text == "")
                {
                    MessageBox.Show("Заполните поле Разряд!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }

                return 0;
            }
            else if (table == "car")
            {
                if (textBox6.Text == "")
                {
                    MessageBox.Show("Заполните корректно поле Номер авто!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (textBox7.Text == "")
                {
                    MessageBox.Show("Заполните корректно поле Марка!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (textBox8.Text == "")
                {
                    MessageBox.Show("Заполните корректно поле Модель!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (textBox9.Text == "" || Convert.ToInt32(textBox9.Text) < 1805 || Convert.ToInt32(textBox9.Text) > 2019)
                {
                    MessageBox.Show("Заполните корректно поле Год выпуска!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (comboBox2.Text == "")
                {
                    MessageBox.Show("Заполните поле Тип кузова!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }

                return 0;
            }

            return 1;
        }


    }
}
