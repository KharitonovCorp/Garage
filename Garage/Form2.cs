// This is an open source non-commercial project. Dear PVS-Studio, please check it.
// PVS-Studio Static Code Analyzer for C, C++, C#, and Java: http://www.viva64.com
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
    //Форма добавления данных
    public partial class Form2 : Form
    {
        //Инициализация переменных используемых в дальнейшем
        string car = "car";
        string mechanic = "mechanic";
        //Инициализия элементов и компонентов формы
        public Form2()
        {
            InitializeComponent();
            InitializeInterface();
            DataGridREView(car);
            DataGridREView(mechanic);
        }

        //Метод инициализации элементов интерфейса
        private void InitializeInterface()
        {
            //Названия формы и контролов
            this.Text = "Добавление данных";
            tabPage1.Text = "Механики";
            tabPage2.Text = "Машины";
            //Контрол механики
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
            //Контрол машины
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
            //Получение количества записей
            GetNoteCount(car);
            GetNoteCount(mechanic);
        }

        //Метод очистки текстовых полей
        private void ClearTextBox(string table)
        {
            //Если контрол механики
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
            //Если контрол машины
            else if (table == "car")
            {
                textBox6.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                comboBox2.Text = "";
            }
        }

        //Метод получения количества записей
        private void GetNoteCount(string table)
        {
            //Если контрол мехиники
            if (table == "mechanic")
            {
                try
                {
                    //Подключение к бд
                    string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                    OleDbConnection oleDbConn = new OleDbConnection(con);
                    oleDbConn.Open();
                    //Создание запроса на получение количества записей
                    OleDbCommand sql = new OleDbCommand("SELECT COUNT(*) FROM " + table + ";", oleDbConn);
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();
                    //Ввод количества записей в текстовый элемент на форму
                    label14.Text = "Кол-во записей: " + (int)sql.ExecuteScalar();

                    oleDbConn.Close();
                }
                catch (Exception ex)
                {
                    //Уведомление об ошибке
                    MessageBox.Show(ex.ToString());
                }
            }
            //Если контрол машины
            if (table == "car")
            {
                try
                {
                    //Подключение к бд
                    string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                    OleDbConnection oleDbConn = new OleDbConnection(con);
                    oleDbConn.Open();
                    //Создание запроса на получение количества записей
                    OleDbCommand sql = new OleDbCommand("SELECT COUNT(*) FROM " + table + ";", oleDbConn);
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();
                    //Ввод количества записей в текстовый элемент на форме
                    label15.Text = "Кол-во записей: " + (int)sql.ExecuteScalar();

                    oleDbConn.Close();
                }
                catch (Exception ex)
                {
                    //Уведомление об ошибке
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        //Преобразование стажа в единый текст
        private string GetMechanicExp()
        {
            string time = textBox5.Text + " " + textBox10.Text + " " + textBox11.Text;
            return time;
        }

        //Метод добавления записей
        private void AddData(string table)
        {
            try
            {
                //Получение случайного числа, используемого, как id
                Random rnd = new Random();
                int temp = rnd.Next(0, 99999);
                //Подключегие к бд
                string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection oleDbConn = new OleDbConnection(con);
                DataTable dt = new DataTable();
                oleDbConn.Open();
                //Если контрол механики
                if (table == "mechanic")
                {
                    //Создание запроса на добавление записи
                    OleDbCommand sql = new OleDbCommand("INSERT INTO " + table + " (mechanic_id, mechanic_number, mechanic_surname,  mechanic_name, mechanic_patronymic, mechanic_exp, mechanic_rank) VALUES (" + Convert.ToInt32(temp) + " , '" + textBox1.Text + "' , '" + textBox2.Text + "' , '" + textBox3.Text + "' , '" + textBox4.Text + "' , '" + GetMechanicExp() + "' , '" + comboBox1.Text + "')");
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();
                }
                //Елси контрол машины
                else if (table == "car")
                {
                    //Создание запроса на добавление записиы
                    OleDbCommand sql = new OleDbCommand("INSERT INTO " + table + " (car_id, car_number,  car_mark, car_name , car_type, car_year) VALUES (" + Convert.ToInt32(temp) + " , '" + textBox6.Text + "' , '" + textBox7.Text + "' , '" + textBox8.Text + "' , '" + comboBox2.Text + "' , '" + textBox9.Text + "')");
                    sql.Connection = oleDbConn;
                    sql.ExecuteNonQuery();
                }
                //Закрытие подключения
                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                //Уведомление об ошибке
                MessageBox.Show(ex.ToString());
            }
            //Обновление количества записей
            GetNoteCount(car);
            GetNoteCount(mechanic);
        }

        //Получение данных в DataGridView
        private void DataGridREView(string table)
        {
            try
            {
                //Подключение к бд
                string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection oleDbConn = new OleDbConnection(con);
                DataTable dt = new DataTable();
                oleDbConn.Open();
                //Создание запроса на получение информации
                OleDbCommand sql = new OleDbCommand("SELECT * FROM " + table + ";");
                sql.Connection = oleDbConn;
                sql.ExecuteNonQuery();

                OleDbDataAdapter da = new OleDbDataAdapter(sql);
                da.Fill(dt);
                //Ввод информации в контрол машины
                if (table == "car")
                {
                    dt.Columns["car_id"].ColumnName = "id";
                    dt.Columns["car_number"].ColumnName = "номер авто";
                    dt.Columns["car_mark"].ColumnName = "марка";
                    dt.Columns["car_name"].ColumnName = "модель";
                    dt.Columns["car_type"].ColumnName = "тип";
                    dt.Columns["car_year"].ColumnName = "год выпуска";
                    dataGridView2.DataSource = dt;
                }
                //Ввод информации в контрол мехники
                if (table == "mechanic")
                {
                    dt.Columns["mechanic_id"].ColumnName = "id";
                    dt.Columns["mechanic_number"].ColumnName = "номер";
                    dt.Columns["mechanic_surname"].ColumnName = "фамилия";
                    dt.Columns["mechanic_name"].ColumnName = "имя";
                    dt.Columns["mechanic_patronymic"].ColumnName = "отчество";
                    dt.Columns["mechanic_exp"].ColumnName = "стаж";
                    dt.Columns["mechanic_rank"].ColumnName = "разряд";
                    dataGridView1.DataSource = dt;
                }
                //Закрытие подключения
                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                //Уведомление об ошибке
                MessageBox.Show(ex.ToString());
            }
        }

        //Кнопка перехода на главную форму
        private void button1_Click(object sender, EventArgs e)
        {
            GoBackToMainForm();
        }

        //Метод перехода на голавную форму
        private void GoBackToMainForm()
        {
            this.Close();
            Form MainForm = new Form1();
            MainForm.Show();
        }

        //Кнопка очистки полей ввода
        private void button3_Click(object sender, EventArgs e)
        {
            //Подтверждение запроса
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно очистить поля ввода?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                //Очистка полей вода, уведомление
                ClearTextBox(mechanic);
                MessageBox.Show("Поля очищены", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //Кнопка добавления записи
        private void button2_Click(object sender, EventArgs e)
        {
            //Проверка заполненности полей
            var result = TextCheck(mechanic);
            if (result == 1)
            {
                return;
            }
            else if (result == 0)
            {
                //Добавление записи, очистка полей, обновление DataGridView, уведомление
                AddData(mechanic);
                ClearTextBox(mechanic);
                DataGridREView(mechanic);
                MessageBox.Show("Запись добавлена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //Кнопка добавление записи
        private void button5_Click(object sender, EventArgs e)
        {
            //Проверка заполненности полей
            var result = TextCheck(car);
            if (result == 1)
            {
                return;
            }
            else if (result == 0)
            {
                //Добавление записи в бд, очистка полей ввода, обновление DataGridView, уведомление
                AddData(car);
                ClearTextBox(car);
                DataGridREView(car);
                MessageBox.Show("Запись добавлена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //Кнопка возврата к главному меню
        private void button6_Click(object sender, EventArgs e)
        {
            GoBackToMainForm();
        }

        //Кнопка очистки полей ввода
        private void button4_Click(object sender, EventArgs e)
        {
            //Подтверждение запроса
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно очистить поля ввода?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                //Очистка полей ввода, уведомление
                ClearTextBox(car);
                MessageBox.Show("Поля очищены", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //Метод запрета ввода
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

        //Метод запрета ввода текстовых символов
        private void TextNoInput(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 9)
            {
                e.Handled = true;
            }
        }

        //Метод запрета ввода чисел
        private void NumberNoInput (object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar)) return;
            else
                e.Handled = true;
        }

        //Метод проверки заполненности текстовых полей
        public int TextCheck(string table)
        {
            //Елси контрол механик
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
            //Если контрол машина
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
