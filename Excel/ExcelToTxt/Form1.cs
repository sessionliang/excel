using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToTxt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            bgWorker.DoWork += BgWorker_DoWork;
            bgWorker.RunWorkerCompleted += BgWorker_RunWorkerCompleted;
        }

        //excel文件路径
        private string excelFilePath;
        //保存文件路径
        private string saveFilePath;

        //源表字段
        private const string CH = "串号";
        private const string SPMC = "商品名称";
        private const string CH2 = "串号2";

        //目标表字段
        private const string d_CH = "串号";
        private const string d_CH2 = "串号2";


        /// <summary>
        /// 选择excel文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, EventArgs e)
        {
            tbFile.Text = excelFilePath = OpenDialogAndSelectFile();
        }

        /// <summary>
        /// 选择保存路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSavePath_Click(object sender, EventArgs e)
        {
            tbSavePath.Text = saveFilePath = OpenDialogAndSaveFile();
        }

        /// <summary>
        /// 确定按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnComfirm_Click(object sender, EventArgs e)
        {

            try
            {
                //显示加载动画
                loadingCtl.ShowLoading(this);
                Application.DoEvents();
                bgWorker.RunWorkerAsync();

                #region 打开连接
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=Excel 8.0;";
                OleDbConnection oleDbConn = new OleDbConnection(strConn);
                if (oleDbConn.State == ConnectionState.Closed)
                    oleDbConn.Open();
                #endregion

                int num = 0;

                string selectStr = "SELECT 串号,商品名称,串号2 FROM {0} ";
                string orderStr = " ORDER BY 商品名称";
                //string groupStr = " GROUP BY 商品名称";

                string dbCmdStr = string.Empty;
                dbCmdStr = String.Join(" ", selectStr, orderStr);
                DataTable tmp = GetContentsByExcelFile(dbCmdStr, oleDbConn);

                #region 关闭连接
                oleDbConn.Close();
                #endregion

                if (tmp != null && tmp.Rows.Count > 0)
                {
                    SetRowsToTxt(tmp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误！" + ex.Message);
            }
            loadingCtl.Hide();
        }

        /// <summary>
        /// 把datatable 写入 txt
        /// </summary>
        /// <param name="dt"></param>
        private void SetRowsToTxt(DataTable dt)
        {
            if (!Directory.Exists(Path.GetDirectoryName(tbSavePath.Text)))
                Directory.CreateDirectory(Path.GetDirectoryName(tbSavePath.Text));

            string date = DateTime.Now.ToString("yyyyMMdd");
            string filePath = this.tbSavePath.Text;
            //商品名称  数量  日期
            string fileName = "北京 非定制机 {0}（{1}）{2}.txt";

            int count = 0;
            string beforName = "";
            StringBuilder createBuilder = new StringBuilder();
            foreach (DataRow row in dt.Rows)
            {
                if (row != null)
                {
                    //第一次，赋值beforName
                    if (string.IsNullOrEmpty(beforName))
                    {
                        beforName = row["商品名称"].ToString();
                    }
                    if (beforName.Equals(row["商品名称"].ToString()))
                    {
                        if (row["串号2"] != null && row["串号2"].ToString().Length > 0)
                        {
                            //属于同一个型号
                            createBuilder.AppendFormat("{0},{1}\r\n",
                                row["串号"] != null ? row["串号"].ToString() : "",
                                row["串号2"] != null ? row["串号2"].ToString() : ""
                                );
                            count++;
                        }
                    }
                    else
                    {
                        //不属于同一个型号，创建文件，先清空createBuilder，然后再赋值
                        string file = Path.Combine(filePath, string.Format(fileName, beforName, count, date));
                        WriteTxtToFile(file, createBuilder.ToString());


                        //清空
                        createBuilder.Length = 0;
                        if (row["串号2"] != null && row["串号2"].ToString().Length > 0)
                        {
                            //写入每个型号第一个
                            createBuilder.AppendFormat("{0},{1}\r\n",
                            row["串号"] != null ? row["串号"].ToString() : "",
                            row["串号2"] != null ? row["串号2"].ToString() : ""
                            );
                            count = 1;
                        }
                        else
                        {
                            count = 0;
                        }
                    }
                    //上面操作完成，记录本次的型号，以便下次对比
                    beforName = row["商品名称"].ToString();
                }
            }
        }

        private void WriteTxtToFile(string filePath, string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                filePath = Path.Combine(Path.GetDirectoryName(filePath), "==空==" + Path.GetFileNameWithoutExtension(filePath) + Path.GetExtension(filePath));
            }

            using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                using (StreamWriter writer = new StreamWriter(fs, Encoding.Default))
                {
                    writer.Write(content);
                }
            }
        }

        /// <summary>
        /// 弹出对话框，选择文件
        /// </summary>
        public string OpenDialogAndSelectFile()
        {
            string path = "";
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "文本文件(*.xls)|*.*";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    path = dialog.FileName;
                }
            }
            return path;
        }

        /// <summary>
        /// 弹出对话框，选择文件保存路劲
        /// </summary>
        public string OpenDialogAndSaveFile()
        {
            string path = "";
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择要保存路径";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    path = dialog.SelectedPath;
                }
            }
            return path;
        }

        /// <summary>
        /// 转换excel中的sheet为datatable
        /// </summary>
        /// <param name="oleDbCmdStr">选择sql语句，{0}代表表名</param>
        /// <returns></returns>
        public DataTable GetContentsByExcelFile(string oleDbCmdStr, OleDbConnection oleDbConn)
        {
            OleDbCommand oleDbCmd;
            OleDbDataAdapter oleDbAdp;
            DataTable oleDt = new DataTable();
            DataTable dtTableName = new DataTable();
            if (oleDbConn.State == ConnectionState.Closed)
                oleDbConn.Open();
            dtTableName = oleDbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string[] tableNames = new string[dtTableName.Rows.Count];
            for (int j = 0; j < dtTableName.Rows.Count; j++)
            {
                tableNames[j] = dtTableName.Rows[j]["TABLE_NAME"].ToString();
            }

            foreach (string tableName in tableNames)
            {
                if (!tableName.EndsWith("$"))
                    continue;

                try
                {
                    oleDbCmd = new OleDbCommand(string.Format(oleDbCmdStr, "[" + tableName + "]"), oleDbConn);
                    oleDbAdp = new OleDbDataAdapter(oleDbCmd);
                    oleDbAdp.Fill(oleDt);
                    break;
                }
                catch { }
            }

            return oleDt;
        }

        #region loading window
        BackgroundWorker bgWorker = new BackgroundWorker();
        LoadingCtl loadingCtl = new LoadingCtl(162, true);//定义加载控件
        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {

            System.Threading.Thread.Sleep(3333);

        }

        private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            loadingCtl.Hide();

        }
        #endregion
    }
}
