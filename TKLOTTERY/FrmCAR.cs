using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using FastReport;
using FastReport.Data;

namespace TKLOTTERY
{
    public partial class FrmCAR : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        int PEO = 0;
        int CAR = 0;

        public FrmCAR()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void LOTTERY()
        {
            
           
            if(PEO<= CAR)
            {
                GETCARNO();
            }
            else if(PEO >CAR)
            {
                STARTLOTTERY();
            }

            

        }
        public void STARTLOTTERY()
        {

        }

        public void GETCARNO()
        {
            int[] numbers = { 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34 };
            int CARNO;
            Random rnd = new Random(Guid.NewGuid().GetHashCode());

            CARNO = rnd.Next(1, CAR);
            textBox4.Text = numbers[CARNO].ToString();

            int numToRemove = numbers[CARNO];
            numbers = numbers.Where(val => val != numToRemove).ToArray();


            CAR = CAR - 1;
            
            textBox3.Text = CAR.ToString();

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            PEO = PEO - 1;
            textBox2.Text = PEO.ToString();

            if(PEO>=1)
            {
                LOTTERY();
            }
            else
            {
                MessageBox.Show("抽完了!");
            }
            
        }
        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Text = textBox1.Text;
            PEO = Convert.ToInt16(textBox2.Text);
            CAR = Convert.ToInt16(textBox3.Text);
        }
        #endregion


    }
}
