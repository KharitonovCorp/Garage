using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Garage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitializeInterface();
        }

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

        private void button7_Click(object sender, EventArgs e)
        {
            var result = new DialogResult();
            result = MessageBox.Show("Вы действительно желаете выйти?", "Внимание!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form AddForm = new Form2();
            AddForm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form EditDelForm = new Form3();
            EditDelForm.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form AddDealForm = new Form4();
            AddDealForm.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form DelDealForm = new Form5();
            DelDealForm.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form OutputInfoForm = new Form6();
            OutputInfoForm.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form CreatorInfoForm = new Form7();
            CreatorInfoForm.ShowDialog();
        }
    }
}
