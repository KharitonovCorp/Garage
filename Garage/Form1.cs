﻿// This is an open source non-commercial project. Dear PVS-Studio, please check it.
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
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace Garage
{
    //Главная форма, меню навигации
    public partial class Form1 : Form
    {
        //Инициализация элементов и компонентов формы
        public Form1()
        {
            InitializeComponent();
            InitializeInterface();
        }

        //Метод иницилизации элементов формы
        private void InitializeInterface()
        {
            this.Text = "СУБД Автомастерская";
            label1.Text = "Меню навигации:";
            button1.Text = "Добавление данных";
            button2.Text = "Редактирование и удаление данных";
            button3.Text = "Составление заказа";
            button4.Text = "Просмотр и удаление заказов";
            button5.Text = "Экспорт данных";
            button6.Text = "Информация о разработчике";
            button7.Text = "Выход";
        }
        //Кнопка выхода из программного приложения
        private void button7_Click(object sender, EventArgs e)
        {
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно желаете выйти?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {

                System.Windows.Forms.Application.Exit();
            }
        }

        //Кнопка перехода на форму добавления данных
        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form AddForm = new Form2();
            AddForm.Show();
        }

        //Кнопка перехода на форму редактирования и удаления данных
        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form EditDelForm = new Form3();
            EditDelForm.Show();
        }

        //Кнопка перехода на форму составления заказа
        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form AddDealForm = new Form4();
            AddDealForm.Show();
        }
       
        //Кнопка перехода на форму просмотра, удаления заказов, экспорта данных а word
        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form DelDealForm = new Form5();
            DelDealForm.Show();
        }

        //Кнопка перехода на форму экспорта в excel
        private void button5_Click(object sender, EventArgs e)
        {
            excel_output();
        }

        //Кнопка перехода на форму с информацией о разработчике
        private void button6_Click(object sender, EventArgs e)
        {
            Form CreatorInfoForm = new Form7();
            CreatorInfoForm.ShowDialog();
        }

        //Метод вывода в excel
        private void excel_output()
        {
            try
            {
                //Подключение к бд
                string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Garage.mdb;";
                OleDbConnection myConnection = new OleDbConnection(connectString);
                myConnection.Open();
                //Инициализация компонентов excel
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Workbooks.Add(Type.Missing);

                //Создание первой таблицы в excel
                var car = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.Sheets[1];
                car.Name = "car";
                var cellsD = (Microsoft.Office.Interop.Excel.Range)car.Cells;
                cellsD.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                //Заполнение заголовков столбцов
                car.Cells[1, 1] = "id";
                car.Cells[1, 2] = "Номер машины";
                car.Cells[1, 3] = "Модель машины";
                car.Cells[1, 4] = "Марка машины";
                car.Cells[1, 5] = "Тип машины";
                car.Cells[1, 6] = "Год выпуска";
                //Отправление запроса к бд на вывод данных
                string queryD = "SELECT * FROM car";
                OleDbCommand commandD = new OleDbCommand(queryD, myConnection);
                OleDbDataReader readerD = commandD.ExecuteReader();
                int temp = 2;
                //Заполнение первой таблицы
                while (readerD.Read())
                {
                    car.Cells[temp, 1] = readerD[0].ToString();
                    car.Cells[temp, 2] = readerD[1].ToString();
                    car.Cells[temp, 3] = readerD[2].ToString();
                    car.Cells[temp, 4] = readerD[3].ToString();
                    car.Cells[temp, 5] = readerD[4].ToString();
                    car.Cells[temp, 6] = readerD[5].ToString();
                    temp++;
                }

                //Создение второй таблицы
                var mechanic = (Worksheet)ExcelApp.Sheets.Add(After: ExcelApp.ActiveSheet);
                mechanic.Name = "mechanic";
                var cellsD2 = (Microsoft.Office.Interop.Excel.Range)mechanic.Cells;
                cellsD2.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                //Заполнение заголовков столбцов
                mechanic.Cells[1, 1] = "id";
                mechanic.Cells[1, 2] = "Номер механика";
                mechanic.Cells[1, 3] = "Фамилия механика";
                mechanic.Cells[1, 4] = "Имя механика";
                mechanic.Cells[1, 5] = "Отчество";
                mechanic.Cells[1, 6] = "Стаж";
                mechanic.Cells[1, 7] = "Разряд";
                //Отправление запроса к бд
                string queryD2 = "SELECT * FROM mechanic";
                OleDbCommand commandD2 = new OleDbCommand(queryD2, myConnection);
                OleDbDataReader readerD2 = commandD2.ExecuteReader();
                int temp2 = 2;
                //Заполенение второй таблицы в excel
                while (readerD2.Read())
                {
                    mechanic.Cells[temp2, 1] = readerD2[0].ToString();
                    mechanic.Cells[temp2, 2] = readerD2[1].ToString();
                    mechanic.Cells[temp2, 3] = readerD2[2].ToString();
                    mechanic.Cells[temp2, 4] = readerD2[3].ToString();
                    mechanic.Cells[temp2, 5] = readerD2[4].ToString();
                    mechanic.Cells[temp2, 6] = readerD2[5].ToString();
                    mechanic.Cells[temp2, 7] = readerD2[6].ToString();
                    temp2++;
                }

                //Создание третьей таблицы
                var repair = (Worksheet)ExcelApp.Sheets.Add(After: ExcelApp.ActiveSheet);
                repair.Name = "repair";
                var cellsD3 = (Microsoft.Office.Interop.Excel.Range)mechanic.Cells;
                cellsD3.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                //Заполнение заголовков столбцов
                repair.Cells[1, 1] = "id";
                repair.Cells[1, 2] = "id механика";
                repair.Cells[1, 3] = "id машины";
                repair.Cells[1, 4] = "Дата починки";
                repair.Cells[1, 5] = "Время починки";
                repair.Cells[1, 6] = "Стоимость";
                //Отправление запроса к бд
                string queryD3 = "SELECT * FROM repair";
                OleDbCommand commandD3 = new OleDbCommand(queryD3, myConnection);
                OleDbDataReader readerD3 = commandD3.ExecuteReader();
                int temp3 = 2;
                //Заполнение третьей таблицы excel
                while (readerD3.Read())
                {
                    repair.Cells[temp3, 1] = readerD3[0].ToString();
                    repair.Cells[temp3, 2] = readerD3[1].ToString();
                    repair.Cells[temp3, 3] = readerD3[2].ToString();
                    repair.Cells[temp3, 4] = readerD3[3].ToString();
                    repair.Cells[temp3, 5] = readerD3[4].ToString();
                    repair.Cells[temp3, 6] = readerD3[5].ToString();
                    temp3++;
                }
                //Визуалиция документа excel
                ExcelApp.Visible = true;
                //Прекращение ввода в документ
                readerD.Close();
                car.Columns.AutoFit();
                car.Rows.AutoFit();
                mechanic.Columns.AutoFit();
                mechanic.Rows.AutoFit();
                repair.Columns.AutoFit();
                repair.Rows.AutoFit();
                //Сохранение бд
                ExcelApp.GetSaveAsFilename();
                //Выход из excel
                ExcelApp.Quit();
            }
            catch (Exception ex)
            {
                //собщение об ошибке
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
