﻿using System;
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
    public partial class Form4 : Form
    {
        string repair = "repair";
        string repair_date = "";

        public Form4()
        {
            InitializeComponent();
            InitializeInterface();
            comboboxInput();
        }

        private void InitializeInterface()
        {
            this.Text = "Составление заказа";
            button1.Text = "Назад";
            button2.Text = "Добавить";
            button3.Text = "Очистить поля ввода";
            label1.Text = "Выберите механика";
            label2.Text = "Выберите авто";
            label3.Text = "Выберите дату работы";
            label4.Text = "Введите стоимость";
            label5.Text = "Введите время ремонта";
            label6.Text = "дни";
            label7.Text = "часы";
            label8.Text = "минуты";
        }

        private void ClearTextBox(string table)
        {
            if (table == "repair")
            {
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                comboBox1.Text = "";
                comboBox2.Text = "";
                dateTimePicker1.Text = "";
            }
        }

        private void AddData(string table)
        {
            try
            {
                string A = comboBox1.Text;
                string[] a = A.Split(new Char[] { ' ' });
                string B = comboBox2.Text;
                string[] b = B.Split(new Char[] { ' ' });
                int days = Convert.ToInt32(textBox2.Text);
                int hours = Convert.ToInt32(textBox3.Text);
                int minuts = Convert.ToInt32(textBox4.Text);
                string repair_time = GetRepairTime(days, hours, minuts);
                string repair_cost = textBox1.Text;

                int id = 0;
                Random rnd = new Random();
                id = rnd.Next(1, 999999);

                DateTime SelectedDate = dateTimePicker1.Value;
                var DateRepair = SelectedDate;
                repair_date = DateRepair.ToString("dd/mm/yy");

                string con1 = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection OleDbcon1 = new OleDbConnection(con1);
                OleDbcon1.Open();

                OleDbCommand sql1 = new OleDbCommand("INSERT INTO " + table
                    + " (repair_id, mechanic_id, car_id, repair_date, repair_time, repair_cost) VALUES (" +
                    Convert.ToInt32(id) + "," + Convert.ToInt32(a[0]) + "," + Convert.ToInt32(b[0]) + ",'" + repair_date + "','" + repair_time + "'," +
                    Convert.ToInt32(repair_cost) + ");", OleDbcon1);

                sql1.ExecuteNonQuery();
                OleDbcon1.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void comboboxInput()
        {
            try
            {
                string con1 = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection OleDbcon1 = new OleDbConnection(con1);
                OleDbcon1.Open();
                OleDbCommand sql1 = new OleDbCommand("SELECT mechanic_id AS [id], mechanic_surname AS [фамилия], mechanic_name AS [имя], mechanic_patronymic AS [отчество] FROM mechanic;", OleDbcon1);
                sql1.ExecuteNonQuery();
                OleDbDataReader reader = sql1.ExecuteReader();
                while (reader.Read())
                {
                    comboBox1.Items.Add(reader["id"] + " " + reader["фамилия"] + " " + reader["имя"].ToString() + " " + reader["отчество"].ToString());
                }
                reader.Close();
                OleDbcon1.Close();

                string con2 = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection OleDbcon2 = new OleDbConnection(con2);
                OleDbcon2.Open();
                OleDbCommand sql2 = new OleDbCommand("SELECT car_id AS [id], car_number AS [номер], car_mark AS [марка], car_name AS [модель] FROM car;", OleDbcon2);
                sql2.ExecuteNonQuery();
                OleDbDataReader reader2 = sql2.ExecuteReader();
                while (reader2.Read())
                {
                    comboBox2.Items.Add(reader2["id"] + " " + reader2["номер"] + " " + reader2["марка"].ToString() + " " + reader2["модель"].ToString());
                }
                reader2.Close();
                OleDbcon2.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private string GetRepairTime(int days, int hours, int minuts)
        {
            string minut_time = days.ToString() + "д" + hours.ToString() + "ч" + minuts.ToString() + "м";
            return minut_time;
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

        private void button2_Click(object sender, EventArgs e)
        {
            var result = TextCheck(repair);
            if (result == 1)
            {
                return;
            }
            else if (result == 0)
            {
                AddData(repair);
                ClearTextBox(repair);
                MessageBox.Show("Запись добавлена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно очистить поля ввода?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                ClearTextBox(repair);
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

        private void NumberNoInput(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar)) return;
            else
                e.Handled = true;
        }

        public int TextCheck(string table)
        {
            if (table == "repair")
            {
                if (comboBox1.Text == "")
                {
                    MessageBox.Show("Заполните поле Выберите механика!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (comboBox2.Text == "")
                {
                    MessageBox.Show("Заполните поле Выберите авто!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (dateTimePicker1.Text == "")
                {
                    MessageBox.Show("Заполните корректно поле Выберите дату работы!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (textBox1.Text == "")
                {
                    MessageBox.Show("Заполните корректно поле Введите стоимость!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }
                if (textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || Convert.ToInt32(textBox2.Text) > 31 || Convert.ToInt32(textBox3.Text) > 24 || Convert.ToInt32(textBox4.Text) > 60)
                {
                    MessageBox.Show("Заполните корректно поля Введите время ремонта!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }

                return 0;
            }

            return 1;
        }


    }
}
