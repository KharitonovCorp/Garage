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
using System.IO;

namespace Garage
{
    //Форма вывода в word и удаления заказов
    public partial class Form5 : Form
    {
        private readonly string FileName = Directory.GetCurrentDirectory() + @"\\template.docx";

        //Инициализация переменных, используемых в дальнейшем
        string repair = "repair";
        string temp = "";

        //Инициализация компонентов формы и интерфейса
        public Form5()
        {
            InitializeComponent();
            InitializeInterface();
            DataGridREView(repair);
            GetNoteCount(repair);
        }

        //Метод инициализации интерфейса
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
        }

        //Метод получения количество записей
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
                //Закрытие подключения
                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                //Сообщение об ошибке
                MessageBox.Show(ex.ToString());
            }
        }

        //Метод получения данных для DataGridView
        private void DataGridREView(string table)
        {
            try
            {
                //Подключение к бд
                string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection oleDbConn = new OleDbConnection(con);
                DataTable dt = new DataTable();
                oleDbConn.Open();
                //Отправление запроса
                OleDbCommand sql = new OleDbCommand("SELECT repair_id AS [id], (SELECT mechanic_surname FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [фамилия мастера], (SELECT mechanic_name FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [имя мастера], (SELECT mechanic_patronymic FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [отчество], (SELECT car_name FROM car WHERE repair.car_id = car.car_id) AS [модель авто], (SELECT car_mark FROM car WHERE repair.car_id = car.car_id) AS [марка авто], repair_date AS [дата работы], repair_cost AS [цена] FROM " + table + ";", oleDbConn);
                sql.Connection = oleDbConn;
                sql.ExecuteNonQuery();
                //Заполнение полей DataGridView
                OleDbDataAdapter da = new OleDbDataAdapter(sql);
                da.Fill(dt);

                dataGridView1.DataSource = dt;
                //Закрытие подключения
                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                //Сообщение об ошибке
                MessageBox.Show(ex.ToString());
            }
        }

        //Кнопка перехода к главному меню
        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form MainForm = new Form1();
            MainForm.Show();
        }

        //Кнопка удаления данных
        private void button2_Click(object sender, EventArgs e)
        {
            //Получение согласия на удаление данных
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно желаете удалить запись?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                //Удаление данных, обновление DataGridView и счетчика записей в таблице
                DeleteData(repair);
                DataGridREView(repair);
                GetNoteCount(repair);
                MessageBox.Show("Запись удалена", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            button2.Enabled = false;
            button3.Enabled = false;
        }

        //Метод получение id записи и активация кнопок
        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //Активация кнопок
            button2.Enabled = true;
            button3.Enabled = true;

            try
            {
                //Получение id
                temp = dataGridView1.SelectedCells[0].Value.ToString();
            }
            catch (Exception)
            {
                
            }
        }

        //Метод удаления записи из таблицы
        private void DeleteData(string table)
        {
            try
            {
                //Подключение к бд
                string con = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
                OleDbConnection oleDbConn = new OleDbConnection(con);
                DataTable dt = new DataTable();
                oleDbConn.Open();
                //Отправление запроса на удаление записи
                OleDbCommand sql = new OleDbCommand("DELETE FROM " + table + " WHERE repair_id=" + Convert.ToInt32(temp) + ";");
                sql.Connection = oleDbConn;
                sql.ExecuteNonQuery();

                oleDbConn.Close();
            }
            catch (Exception ex)
            {
                //Уведомление об ошибке
                MessageBox.Show(ex.ToString());
            }
        }

        //Медот вывода данных в word
        private void WordExport()
        {
            //Подключение к бд
            string con1 = "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=Garage.mdb;";
            OleDbConnection OleDbcon1 = new OleDbConnection(con1);
            OleDbcon1.Open();
            //Отправка запроса на получение данных
            OleDbCommand sql1 = new OleDbCommand("SELECT repair_id AS [id], (SELECT mechanic_surname FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [фамилия мастера], (SELECT mechanic_name FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [имя мастера], (SELECT mechanic_patronymic FROM mechanic WHERE repair.mechanic_id = mechanic.mechanic_id) AS [отчество], (SELECT car_name FROM car WHERE repair.car_id = car.car_id) AS [модель авто], (SELECT car_mark FROM car WHERE repair.car_id = car.car_id) AS [марка авто], repair_date AS [дата работы], repair_cost AS [цена] FROM repair WHERE repair_id = " + temp +";", OleDbcon1);
            sql1.ExecuteNonQuery();
            OleDbDataReader reader = sql1.ExecuteReader();
            //Инициализация элемента word и его визуализация
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(FileName);
            while (reader.Read())
            {
                //Ввод полученной из бд информации в word
                wordDocument.Bookmarks["id"].Range.Text = reader["id"].ToString();
                wordDocument.Bookmarks["surname"].Range.Text = reader["фамилия мастера"].ToString();
                wordDocument.Bookmarks["name"].Range.Text = reader["имя мастера"].ToString();
                wordDocument.Bookmarks["patronimyc"].Range.Text = reader["отчество"].ToString();
                wordDocument.Bookmarks["model"].Range.Text = reader["модель авто"].ToString();
                wordDocument.Bookmarks["mark"].Range.Text = reader["марка авто"].ToString();
                wordDocument.Bookmarks["date"].Range.Text = reader["дата работы"].ToString();
                wordDocument.Bookmarks["cost"].Range.Text = reader["цена"].ToString();
            }
            //Просмотр документа
            wordApp.Visible = true;
            //Прекращение ввода
            reader.Close();
            //Закрытие подключения
            OleDbcon1.Close();
        }

        //Кнопка вывода в word
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                WordExport();
                //Деактивация кнопок
                button2.Enabled = false;
                button3.Enabled = false;
            }
            catch (Exception ex)
            {
                //Уведомление об ошибке
                MessageBox.Show(ex.ToString());
            }
           
        }
    }
}
