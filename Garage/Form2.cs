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
    public partial class Form2 : Form
    {
        string car = "car";
        string mechanic = "mechanic";

        public Form2()
        {
            InitializeComponent();
            InitializeInterface();
            DataGridREView(car);
            DataGridREView(mechanic);
        }

        private void InitializeInterface()
        {
            this.Text = "Добавление данных";
            tabPage1.Text = "Механики";
            tabPage2.Text = "Машины";

            button1.Text = "Назад";
            button2.Text = "Добавить";
            button3.Text = "Очистить поля ввода";
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
            button5.Text = "Добавить";
            button4.Text = "Очистить поля ввода";
            label7.Text = "Номер авто";
            label8.Text = "Марка";
            label9.Text = "Модель";
            label10.Text = "Год выпуска";
            label11.Text = "Тип кузова";
            string[] car_types = { "Седан", "Хэтчбэк", "Универсал", "Лифтбэк", "Купе", "Кабриолет", "Родстер", "Тарга" };
            comboBox2.Items.AddRange(car_types);

            GetNoteCount(car);
            GetNoteCount(mechanic);
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

        private string GetMechanicExp()
        {
            string time = textBox5.Text + " " + textBox10.Text + " " + textBox11.Text;
            return time;
        }

        private void AddData(string table)
        {
            try
            {
                Random rnd = new Random();
                int temp = rnd.Next(0, 99999);

                string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection oleDbConn = new OleDbConnection(con);
                DataTable dt = new DataTable();
                oleDbConn.Open();
                if (table == "mechanic")
                {
                    OleDbCommand sql = new OleDbCommand("INSERT INTO " + table + " (mechanic_id, mechanic_number, mechanic_surname,  mechanic_name, mechanic_patronymic, mechanic_exp, mechanic_rank) VALUES (" + Convert.ToInt32(temp) + " , '" + textBox1.Text + "' , '" + textBox2.Text + "' , '" + textBox3.Text + "' , '" + textBox4.Text + "' , '" + GetMechanicExp() + "' , '" + comboBox1.Text + "')");
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();
                }
                else if (table == "car")
                {
                    OleDbCommand sql = new OleDbCommand("INSERT INTO " + table + " (car_id, car_number,  car_mark, car_name , car_type, car_year) VALUES (" + Convert.ToInt32(temp) + " , '" + textBox6.Text + "' , '" + textBox7.Text + "' , '" + textBox8.Text + "' , '" + comboBox2.Text + "' , '" + textBox9.Text + "')");
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

        private void button1_Click(object sender, EventArgs e)
        {
            GoBackToMainForm();
        }

        private void GoBackToMainForm()
        {
            this.Close();
            Form MainForm = new Form1();
            MainForm.Show();
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

        private void button2_Click(object sender, EventArgs e)
        {

            var result = TextCheck(mechanic);
            if (result == 1)
            {
                return;
            }
            else if (result == 0)
            {
                AddData(mechanic);
                ClearTextBox(mechanic);
                DataGridREView(mechanic);
                MessageBox.Show("Запись добавлена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
                AddData(car);
                ClearTextBox(car);
                DataGridREView(car);
                MessageBox.Show("Запись добавлена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            GoBackToMainForm();
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

        private void NumberNoInput (object sender, KeyPressEventArgs e)
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
