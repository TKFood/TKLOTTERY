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
        // 数据集（用于存储查询结果）
        private DataSet _dataSetCarNumbers = new DataSet();
        private DataSet _dataSetSearchResult = new DataSet();

        // 业务逻辑相关
        private int _totalPeople = 0;           // 本次抽签的总人数
        private int _availableParking = 0;      // 可用的停车位数
        private int _remainingPeople = 0;       // 还有多少人未抽
        private int _drawCount = 0;             // 当前抽签次数

        // 停车位数据
        private int[] _parkingNumbers = new int[] { };
        private int[] _winnerNumbers = new int[] { };

        // 线程安全的随机数生成器
        private static readonly Random _random = new Random();
        private readonly object _randomLock = new object();

        public FrmCAR()
        {
            InitializeComponent();
        }

        #region FUNCTION

        /// <summary>
        /// 記錄錯誤日誌
        /// </summary>
        private void LogError(string methodName, Exception ex)
        {
            try
            {
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 方法: {methodName}" + Environment.NewLine +
                                   $"錯誤訊息: {ex.Message}" + Environment.NewLine +
                                   $"堆疊追蹤: {ex.StackTrace}" + Environment.NewLine +
                                   "---" + Environment.NewLine;

                string logPath = Path.Combine(Application.StartupPath, "Logs");
                if (!Directory.Exists(logPath))
                    Directory.CreateDirectory(logPath);

                string logFile = Path.Combine(logPath, $"Error_{DateTime.Now:yyyyMMdd}.log");
                File.AppendAllText(logFile, logMessage);
            }
            catch
            {
                // 日誌記錄失敗時不拋出異常
            }
        }

        /// <summary>
        /// 顯示錯誤訊息並記錄日誌
        /// </summary>
        private void ShowError(string title, Exception ex)
        {
            LogError(title, ex);
            MessageBox.Show($"{title}{Environment.NewLine}{Environment.NewLine}詳情: {ex.Message}", 
                           "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 獲取資料庫連接字符串
        /// </summary>
        private string GetConnectionString()
        {
            try
            {
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);
                return sqlsb.ConnectionString;
            }
            catch (Exception ex)
            {
                ShowError("獲取連接字符串失敗", ex);
                return null;
            }
        }

        /// <summary>
        /// 執行 SELECT 查詢
        /// </summary>
        private DataSet ExecuteQuery(string sql)
        {
            DataSet dataSet = new DataSet();
            try
            {
                if (string.IsNullOrEmpty(sql))
                {
                    LogError("ExecuteQuery", new ArgumentNullException(nameof(sql)));
                    return dataSet;
                }

                string connStr = GetConnectionString();
                if (connStr == null) return dataSet;

                using (SqlConnection conn = new SqlConnection(connStr))
                using (SqlDataAdapter adapter = new SqlDataAdapter(sql, conn))
                {
                    adapter.Fill(dataSet);
                }
            }
            catch (SqlException sqlEx)
            {
                ShowError("資料庫查詢失敗", sqlEx);
            }
            catch (Exception ex)
            {
                ShowError("查詢執行失敗", ex);
            }
            return dataSet;
        }

        /// <summary>
        /// 執行 INSERT/UPDATE/DELETE 命令
        /// </summary>
        private bool ExecuteCommand(string sql, Dictionary<string, object> parameters = null)
        {
            try
            {
                if (string.IsNullOrEmpty(sql))
                {
                    LogError("ExecuteCommand", new ArgumentNullException(nameof(sql)));
                    return false;
                }

                string connStr = GetConnectionString();
                if (connStr == null) return false;

                using (SqlConnection conn = new SqlConnection(connStr))
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    conn.Open();
                    SqlTransaction tran = conn.BeginTransaction();
                    cmd.Transaction = tran;
                    cmd.CommandTimeout = 60;

                    if (parameters != null)
                    {
                        foreach (var param in parameters)
                        {
                            cmd.Parameters.AddWithValue(param.Key, param.Value);
                        }
                    }

                    try
                    {
                        int result = cmd.ExecuteNonQuery();
                        if (result > 0)
                        {
                            tran.Commit();
                            return true;
                        }
                        else
                        {
                            tran.Rollback();
                            return false;
                        }
                    }
                    catch (SqlException sqlEx)
                    {
                        tran.Rollback();
                        ShowError("執行 SQL 命令失敗", sqlEx);
                        return false;
                    }
                    catch (Exception ex)
                    {
                        tran.Rollback();
                        ShowError("執行命令失敗", ex);
                        return false;
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                ShowError("資料庫連接失敗", sqlEx);
                return false;
            }
            catch (Exception ex)
            {
                ShowError("命令執行失敗", ex);
                return false;
            }
        }

        /// <summary>
        /// 获取线程安全的随机数
        /// </summary>
        private int GetRandomNumber(int minValue, int maxValue)
        {
            lock (_randomLock)
            {
                return _random.Next(minValue, maxValue);
            }
        }

        /// <summary>
        /// 使用Fisher-Yates洗牌算法生成不重复的随机序列
        /// </summary>
        public void PerformLottery()
        {
            try
            {
                if (_remainingPeople <= 0 || _availableParking <= 0)
                {
                    MessageBox.Show("請先設定人數和車位數", "驗證", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 创建包含 1 到 _remainingPeople 的数组
                int[] numbers = new int[_remainingPeople];
                for (int i = 0; i < _remainingPeople; i++)
                {
                    numbers[i] = i + 1;
                }

                // Fisher-Yates 洗牌算法
                for (int i = _remainingPeople - 1; i > 0; i--)
                {
                    int randomIndex = GetRandomNumber(0, i + 1);
                    // 交换
                    int temp = numbers[i];
                    numbers[i] = numbers[randomIndex];
                    numbers[randomIndex] = temp;
                }

                // 取前 _availableParking 个数作为抽中的号码
                _winnerNumbers = new int[_availableParking];
                Array.Copy(numbers, 0, _winnerNumbers, 0, _availableParking);

                // 构建结果字符串
                StringBuilder result = new StringBuilder();
                foreach (int number in _winnerNumbers)
                {
                    result.Append(number).Append(" ");
                }

                textBox8.Text = textBox8.Text + result.ToString().Trim() + Environment.NewLine;
            }
            catch (Exception ex)
            {
                ShowError("抽簽失敗", ex);
            }
        }
        public void CheckLotteryResult()
        {
            try
            {
                bool isWinner = false;

                // 检查当前号码是否在中奖号码中
                foreach (int number in _winnerNumbers)
                {
                    if (_drawCount == number)
                    {
                        isWinner = true;
                        break;
                    }
                }

                if (isWinner)
                {
                    AllocateParking();
                }
                else
                {
                    textBox4.Text = "很抱歉，您未抽中!";
                    textBox5.Text = textBox5.Text + Environment.NewLine + _drawCount.ToString() + Environment.NewLine + textBox4.Text;
                }

                _totalPeople = _totalPeople - 1;
            }
            catch (Exception ex)
            {
                ShowError("抽簽檢查失敗", ex);
            }
        }

        public void AllocateParking()
        {
            try
            {
                if (_parkingNumbers.Length == 0)
                {
                    MessageBox.Show("沒有可用的車位", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int randomIndex = GetRandomNumber(0, _parkingNumbers.Length);
                int selectedParkingNumber = _parkingNumbers[randomIndex];

                textBox4.Text = "您抽中了 " + selectedParkingNumber.ToString() + " 號車位!";
                textBox5.Text = textBox5.Text + Environment.NewLine + _drawCount.ToString() + Environment.NewLine + textBox4.Text;

                // 从数组中移除已分配的车位
                List<int> remainingParking = new List<int>(_parkingNumbers);
                remainingParking.RemoveAt(randomIndex);
                _parkingNumbers = remainingParking.ToArray();

                _availableParking = _availableParking - 1;
                textBox3.Text = _availableParking.ToString();
            }
            catch (Exception ex)
            {
                ShowError("分配車位失敗", ex);
            }
        }

        public void INITAILCARNO()
        {
            _dataSetCarNumbers.Clear();

            try
            {
                string sql = @"SELECT [ID] FROM [TKLOTTERY].[dbo].[CARNO] ORDER BY CONVERT(INT,[ID])";
                _dataSetCarNumbers = ExecuteQuery(sql);

                if (_dataSetCarNumbers.Tables.Count > 0 && _dataSetCarNumbers.Tables[0].Rows.Count >= 1)
                {
                    LoadParkingNumbers(_dataSetCarNumbers.Tables[0]);
                }
                else
                {
                    MessageBox.Show("警告: 未找到車位資料", "初始化", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ShowError("初始化車位失敗", ex);
            }
        }

        public void LoadParkingNumbers(DataTable dataTable = null)
        {
            _parkingNumbers = new int[] { };

            try
            {
                DataTable table = dataTable;
                if (table == null && _dataSetCarNumbers.Tables.Count > 0)
                {
                    table = _dataSetCarNumbers.Tables[0];
                }

                if (table != null)
                {
                    foreach (DataRow row in table.Rows)
                    {
                        try
                        {
                            int parkingNumber = Convert.ToInt16(row["ID"].ToString());
                            _parkingNumbers = _parkingNumbers.Concat(new int[] { parkingNumber }).ToArray();
                        }
                        catch (FormatException ex)
                        {
                            LogError("LoadParkingNumbers-DataConversion", new Exception($"車位 ID 轉換失敗: {row["ID"]}", ex));
                        }
                        catch (OverflowException ex)
                        {
                            LogError("LoadParkingNumbers-DataConversion", new Exception($"車位 ID 超出範圍: {row["ID"]}", ex));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("載入車位編號失敗", ex);
            }
        }

        public void SearchParking()
        {
            _dataSetSearchResult.Clear();

            try
            {
                string sql = @"SELECT [ID] AS '車位' FROM [TKLOTTERY].[dbo].[CARNO] ORDER BY CONVERT(INT,[ID])";
                _dataSetSearchResult = ExecuteQuery(sql);

                if (_dataSetSearchResult.Tables.Count > 0 && _dataSetSearchResult.Tables[0].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = _dataSetSearchResult.Tables[0];
                    dataGridView1.AutoResizeColumns();
                }
                else
                {
                    dataGridView1.DataSource = null;
                    MessageBox.Show("未找到車位資料", "搜尋", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                ShowError("搜尋車位失敗", ex);
            }
        }

        public void AddParking()
        {
            if (!string.IsNullOrEmpty(textBox6.Text))
            {
                try
                {
                    string sql = "INSERT INTO [TKLOTTERY].[dbo].[CARNO] ([ID]) VALUES (@ID)";
                    var parameters = new Dictionary<string, object>
                    {
                        { "@ID", textBox6.Text }
                    };

                    if (ExecuteCommand(sql, parameters))
                    {
                        ClearTextBox();
                        SearchParking();
                        //MessageBox.Show("車位新增成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        //MessageBox.Show("車位新增失敗，請檢查車位號是否重複", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    ShowError("新增車位失敗", ex);
                }
            }
            else
            {
                //MessageBox.Show("請輸入車位號", "驗證", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void DeleteParking()
        {
            try
            {
                if (string.IsNullOrEmpty(textBox7.Text))
                {
                    MessageBox.Show("請先選擇要刪除的車位", "驗證", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string sql = "DELETE [TKLOTTERY].[dbo].[CARNO] WHERE [ID] = @ID";
                var parameters = new Dictionary<string, object>
                {
                    { "@ID", textBox7.Text }
                };

                if (ExecuteCommand(sql, parameters))
                {
                    SearchParking();
                    //MessageBox.Show("車位刪除成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    //MessageBox.Show("車位刪除失敗", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ShowError("刪除車位失敗", ex);
            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.CurrentRow != null)
                {
                    int rowindex = dataGridView1.CurrentRow.Index;
                    if (rowindex >= 0)
                    {
                        DataGridViewRow row = dataGridView1.Rows[rowindex];
                        object cellValue = row.Cells["車位"].Value;
                        if (cellValue != null)
                        {
                            textBox7.Text = cellValue.ToString();
                        }
                    }
                    else
                    {
                        textBox7.Text = null;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("dataGridView1_SelectionChanged", ex);
                textBox7.Text = null;
            }
        }

        public void ClearTextBox()
        {
            textBox6.Text = null;
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            _drawCount = _drawCount + 1;

            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                if (_totalPeople >= 1 && _availableParking >= 1)
                {
                    _remainingPeople = _totalPeople;
                    CheckLotteryResult();
                }
                else
                {
                    MessageBox.Show("抽完了!");
                }

                textBox2.Text = _totalPeople.ToString();
            }
            else
            {
                MessageBox.Show("請按準備");
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // 初始化可以抽的車位號碼
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                INITAILCARNO();
                textBox2.Text = textBox1.Text;
                textBox3.Text = _parkingNumbers.Count().ToString();

                _totalPeople = Convert.ToInt16(textBox2.Text);
                _availableParking = Convert.ToInt16(textBox3.Text);
                _remainingPeople = Convert.ToInt16(textBox2.Text);

                textBox8.Text = null;
                _winnerNumbers = new int[] { };

                // 如果人數多於總車位數，就要進行抽籤
                if (_totalPeople > _availableParking)
                {
                    PerformLottery();
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

            _drawCount = 0;
            textBox4.Text = null;
            textBox5.Text = null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SearchParking();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            AddParking();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DeleteParking();
            }
        }


        #endregion

       
    }
}
