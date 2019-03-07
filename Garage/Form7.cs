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
    //Форма представления информации об разработчике
    public partial class Form7 : Form
    {
        //инициализация элементов формы и компонентов   
        public Form7()
        {
            InitializeComponent();
            this.Text = "Информация о разработчике";
            label1.Text = "Работу выполнил ученик группы ПС-15\n\nХаритонов Илья\n\n2019 год";
            button1.Text = "Назад";
        }

        //метод перехода на главную форму
        private void GoBackToMainForm()
        {
            this.Close();
            Form MainForm = new Form1();
        }

        //Кнопка перехода на главную форму
        private void button1_Click(object sender, EventArgs e)
        {
            GoBackToMainForm();
        }
    }
}
