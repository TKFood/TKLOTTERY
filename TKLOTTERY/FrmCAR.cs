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
using System.Collections;
using TKITDLL;

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
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        int PEO = 0;
        int CAR = 0;
        int PER = 0;
        int NG = 0;
        //int[] CARnumbers = { 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34 };

        int[] CARnumbers = new int[] {  };
        int[] myArray = new int[] { };
        int count = 0;

        public FrmCAR()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void LOTTERY()
        {


            int number;
            string reslut=null;

            //隨機決定那些抽中的號碼，從1到未抽的人數中任選一號，連續抽跟總車位一樣的次數
            //每次抽中後，就檢查是否有重覆抽中，有的話就無限再抽，直到不重覆，且總車位一樣多

            for (int i = 0; i < CAR; i++)
            {
                Random rnd = new Random(Guid.NewGuid().GetHashCode());
                myArray = myArray.Concat(new int[] { rnd.Next(1, PER+1) }).ToArray();

                for (int j = 0; j < i; j++)
                {
                    while (myArray[j] == myArray[i] && i < CAR)
                    {
                        j = 0;
                        myArray[i] = rnd.Next(1, PER);
                    }
                }
            }

            foreach (object i in myArray)
            {
                reslut = reslut + i.ToString() + " ";
            }

            textBox8.Text = textBox8.Text + reslut.ToString();          

          

        }
        public void STARTLOTTERY()
        {
            string check = "N";

            foreach (object i in myArray)
            {
                if (count.ToString().Equals(i.ToString()))
                {
                    check = "Y";
                    break;
                }
               
            }

            if(check.Equals("Y"))
            {
                GETCARNO();
            }
            else
            {
                textBox4.Text = "很抱歉，您未抽中!";
                textBox5.Text = textBox5.Text + Environment.NewLine + count.ToString() + Environment.NewLine + textBox4.Text;
            }

            PEO = PEO - 1;

        }

        public void GETCARNO()
        {           
            int CARNO;
            Random rnd = new Random(Guid.NewGuid().GetHashCode());

            CARNO = rnd.Next(1, CAR)-1;
            textBox4.Text = "您抽中了 " + CARnumbers[CARNO].ToString() + " 號車位!";

            textBox5.Text = textBox5.Text + Environment.NewLine + count.ToString() + Environment.NewLine + textBox4.Text;

            int numToRemove = CARnumbers[CARNO];
            int numIdx = Array.IndexOf(CARnumbers, numToRemove);
            List<int> tmp = new List<int>(CARnumbers);
            tmp.RemoveAt(numIdx);
            CARnumbers = tmp.ToArray();


            CAR = CAR - 1;            
            textBox3.Text = CAR.ToString();

        }

        public void INITAILCARNO()
        {
            ds.Clear();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                
                sbSql.AppendFormat(@" 
                                    SELECT  [ID]  FROM [TKLOTTERY].[dbo].[CARNO] ORDER BY CONVERT(INT,[ID])
                                    ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                   
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        ADDCARNO();
                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

        }

        public void ADDCARNO()
        {
            CARnumbers = new int[] { };

            foreach (DataRow od in ds.Tables["TEMPds1"].Rows)
            {
                CARnumbers=CARnumbers.Concat(new int[] { Convert.ToInt16(od["ID"].ToString()) }).ToArray();
                //CARnumbers = CARnumbers.Concat(new int[] { 2 }).ToArray();
            }
                
        }

        public void SEARCHCARNO()
        {
            ds2.Clear();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();

                sbSql.AppendFormat(@" SELECT  [ID] AS '車位'  FROM [TKLOTTERY].[dbo].[CARNO] ORDER BY CONVERT(INT,[ID])");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView1.AutoResizeColumns();
                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void ADDCARDNO()
        {
            if(!string.IsNullOrEmpty(textBox6.Text)&& !string.IsNullOrEmpty(textBox7.Text))
            {
                try
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(@"  
                                        INSERT INTO [TKLOTTERY].[dbo].[CARNO]  ([ID]) VALUES ('{0}')"
                                        , textBox6.Text);
                    sbSql.AppendFormat(" ");


                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                    }
                    else
                    {
                        tran.Commit();      //執行交易  


                    }

                }
                catch
                {

                }

                finally
                {
                    sqlConn.Close();
                }
            }
           
        }

        public void DECARNO()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  DELETE [TKLOTTERY].[dbo].[CARNO] WHERE [ID] ='{0}'",textBox7.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox7.Text = row.Cells["車位"].Value.ToString();
                    
                }
                else
                {
                    textBox7.Text = null;
                    
                }
            }
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            count = count + 1;

            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                if (PEO >= 1 && CAR >= 1)
                {
                    PER = PEO;
                    STARTLOTTERY();
                }
                else
                {
                    MessageBox.Show("抽完了!");
                }


                textBox2.Text = PEO.ToString();
            }
            else
            {
                MessageBox.Show("請按準備");
            }
           
            
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //先用INITAILCARNO()，初始化可以抽的車位號碼
            //PEO本次要抽的總人數
            //CAR可以被抽的總車位數
            //PER還有多少人未抽


            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                INITAILCARNO();
                textBox2.Text = textBox1.Text;
                textBox3.Text = CARnumbers.Count().ToString();

                PEO = Convert.ToInt16(textBox2.Text);
                CAR = Convert.ToInt16(textBox3.Text);
                PER = Convert.ToInt16(textBox2.Text);
                
                textBox8.Text = null;
                myArray = new int[] { };


                //如果未抽人數多於總車位數，就要抽 LOTTERY()


                if (PEO> CAR)
                {
                    LOTTERY();
                }
                else
                {
                    MessageBox.Show("人數比停車位少，可以直接配停車位!");
                }                
               
            }
            else
            {
                MessageBox.Show("請填本次抽車位人數");
            }

            count = 0;
            textBox4.Text = null;
            textBox5.Text = null;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCHCARNO();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADDCARDNO();
            SEARCHCARNO();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DECARNO();
                SEARCHCARNO();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }


        #endregion

       
    }
}
