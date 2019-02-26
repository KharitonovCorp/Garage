using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Garage
{
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
            InitializeInterface();
        }

        private void InitializeInterface()
        {
            this.Text = "Экспорт данных в Excel";
            button1.Text = "Назад";
            button2.Text = "Отчет";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GoToMainForm();
        }

        private void GoToMainForm()
        {
            this.Hide();
            Form MainForm = new Form1();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            excel_output();
        }

        private void excel_output()
        {
            string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Garage.mdb;";

            OleDbConnection myConnection = new OleDbConnection(connectString);

            myConnection.Open();

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            ExcelApp.Workbooks.Add(Type.Missing);

            ////////////////////////////////////////////////////////////////////////////

            var car = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.Sheets[1];

            car.Name = "car";

            var cellsD = (Microsoft.Office.Interop.Excel.Range)car.Cells;

            cellsD.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            car.Cells[1, 1] = "id";

            car.Cells[1, 2] = "Номер машины";

            car.Cells[1, 3] = "Модель машины";

            car.Cells[1, 4] = "Марка машины";

            car.Cells[1, 5] = "Тип машины";

            car.Cells[1, 6] = "Год выпуска";

            string queryD = "SELECT * FROM car";

            OleDbCommand commandD = new OleDbCommand(queryD, myConnection);

            OleDbDataReader readerD = commandD.ExecuteReader();

            int temp = 2;

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

            ExcelApp.Visible = true;
            readerD.Close();

            car.Columns.AutoFit();

            car.Rows.AutoFit();
            ///////////////////////////////////////////////////////////////////////////

        }
    }
}
