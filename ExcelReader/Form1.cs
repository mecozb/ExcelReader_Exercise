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
using System.IO;

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source=C:/Users/Administrator/Desktop/faturaTest.xlsx;
            Extended Properties='Excel 12.0 xml;HDR=YES;'");

        private void button1_Click(object sender, EventArgs e)
        {
            conn.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Parametreler$]", conn); //Parametreler = excel sheet name
            DataTable dt = new DataTable();
            da.Fill(dt);
            
            dataGridView1.DataSource = dt.DefaultView;
            string a = dt.Rows[2][0].ToString();
            conn.Close();

        }
    }
}
