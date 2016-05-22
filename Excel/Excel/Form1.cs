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

namespace Excel
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
        //模糊查询分组
        private List<string> vagueList;
        //精确查询分组
        private List<string> accurateList;

        //源表字段
        private const string SHCK = "收货仓库";
        private const string YS = "颜色";
        private const string SPMC = "商品名称";
        private const string CH = "串号";
        private const string MEID = "meid";

        //目标表字段
        private const string XH = "序号";
        private const string JX = "机型";
        private const string CM = "串码";
        private const string RKRQ = "入库日期";
        private const string BZ = "备注";


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

                /* * * * * * * * * * * * * * * * *
                 * 序号  机型  串码  入库日期  备注 *
                 * * * * * * * * * * * * * *  * * */
                DataTable dt = new DataTable();
                dt.Columns.Add(new DataColumn("序号"));
                dt.Columns.Add(new DataColumn("机型"));
                dt.Columns.Add(new DataColumn("串码"));
                dt.Columns.Add(new DataColumn("入库日期"));
                dt.Columns.Add(new DataColumn("备注"));

                #region 打开连接
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=Excel 8.0;";
                OleDbConnection oleDbConn = new OleDbConnection(strConn);
                if (oleDbConn.State == ConnectionState.Closed)
                    oleDbConn.Open();
                #endregion

                //获取模糊查询集合
                InitVagueList();
                //获取精确查询集合
                InitAccurateList(oleDbConn);
                int num = 0;

                string selectStr = "SELECT 收货仓库,商品名称,串号,颜色,meid FROM {0} ";
                string orderStr = " ORDER BY 收货仓库,商品名称,颜色 ";
                //string groupStr = string.Format(" GROUP BY {0} ", SHCK);
                foreach (string item in accurateList)
                {
                    string dbCmdStr = string.Empty;
                    string whereStr = " WHERE 1=1 ";
                    whereStr += string.Format(" AND {0} = '{1}' ", SHCK, item);
                    dbCmdStr = String.Join(" ", selectStr, whereStr, orderStr);
                    DataTable tmp = GetContentsByExcelFile(dbCmdStr, oleDbConn);
                    if (tmp.Rows.Count > 0)
                    {
                        foreach (DataRow row in tmp.Rows)
                        {
                            num++;
                            DataRow r = dt.NewRow();
                            r["序号"] = num;
                            r["机型"] = string.Format("{0}-{1}-{2}", tbName.Text, row["商品名称"], row["颜色"]);
                            r["串码"] = row["meid"];
                            r["入库日期"] = DateTime.Now.ToString("yyyy-MM-dd");
                            r["备注"] = row["收货仓库"].ToString().Replace("仓库", string.Empty);
                            dt.Rows.Add(r);
                        }
                    }
                }


                foreach (string item in vagueList)
                {
                    string dbCmdStr = string.Empty;
                    string whereStr = " WHERE 1=1 AND (";
                    whereStr += string.Format(" {0} LIKE '%{1}%' ", SHCK, item);
                    whereStr += string.Format(" OR {0} LIKE '%{1}' ", SHCK, item);
                    whereStr += string.Format(" OR {0} LIKE '{1}%' )", SHCK, item);
                    dbCmdStr = String.Join(" ", selectStr, whereStr, orderStr);
                    DataTable tmp = GetContentsByExcelFile(dbCmdStr, oleDbConn);
                    if (tmp.Rows.Count > 0)
                    {
                        foreach (DataRow row in tmp.Rows)
                        {
                            num++;
                            DataRow r = dt.NewRow();
                            r["序号"] = num;
                            r["机型"] = string.Format("{0}-{1}-{2}", tbName.Text, row["商品名称"], row["颜色"]);
                            r["串码"] = row["meid"];
                            r["入库日期"] = DateTime.Now.ToString("yyyy-MM-dd");
                            r["备注"] = item;
                            dt.Rows.Add(r);
                        }
                    }
                }

                #region 关闭连接
                oleDbConn.Close();
                #endregion

                if (dt != null && dt.Rows.Count > 0)
                {
                    SetRowsToExcel(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误！" + ex.Message);
            }
            loadingCtl.Hide();
        }

        /// <summary>
        /// 把datatable 写入 excel
        /// </summary>
        /// <param name="dt"></param>
        private void SetRowsToExcel(DataTable dt)
        {
            if (!Directory.Exists(Path.GetDirectoryName(tbSavePath.Text)))
                Directory.CreateDirectory(Path.GetDirectoryName(tbSavePath.Text));
            //if (!File.Exists(tbSavePath.Text))
            //    File.Create(tbSavePath.Text).Dispose();
            string fileName = Path.GetFileNameWithoutExtension(tbSavePath.Text);
            //string OLEDBConnStr = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};", tbSavePath.Text);
            string OLEDBConnStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};", tbSavePath.Text);
            OLEDBConnStr += " Extended Properties=Excel 8.0;";
            StringBuilder createBuilder = new StringBuilder();

            createBuilder.AppendFormat("CREATE TABLE {0} ( ", fileName);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                createBuilder.AppendFormat(" [{0}] NTEXT, ", dt.Columns[i]);
            }
            createBuilder.Length = createBuilder.Length - 2;
            createBuilder.Append(" )");

            StringBuilder preInsertBuilder = new StringBuilder();
            preInsertBuilder.AppendFormat("INSERT INTO {0} (", fileName);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                preInsertBuilder.AppendFormat("[{0}], ", dt.Columns[i]);
            }


            preInsertBuilder.Length = preInsertBuilder.Length - 2;
            preInsertBuilder.Append(") VALUES (");

            ArrayList insertSqlArrayList = new ArrayList();
            foreach (DataRow row in dt.Rows)
            {
                if (row != null)
                {
                    StringBuilder insertBuilder = new StringBuilder();
                    insertBuilder.Append(preInsertBuilder.ToString());
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        insertBuilder.AppendFormat("'{0}', ", row[dt.Columns[i]]);
                    }


                    insertBuilder.Length = insertBuilder.Length - 2;
                    insertBuilder.Append(") ");

                    insertSqlArrayList.Add(insertBuilder.ToString());
                }
            }

            OleDbConnection oConn = new OleDbConnection();

            oConn.ConnectionString = OLEDBConnStr;
            OleDbCommand oCreateComm = new OleDbCommand();
            oCreateComm.Connection = oConn;
            oCreateComm.CommandText = createBuilder.ToString();

            oConn.Open();
            oCreateComm.ExecuteNonQuery();
            foreach (string insertSql in insertSqlArrayList)
            {
                OleDbCommand oInsertComm = new OleDbCommand();
                oInsertComm.Connection = oConn;
                oInsertComm.CommandText = insertSql;
                oInsertComm.ExecuteNonQuery();
            }
            oConn.Close();
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
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "文本文件(*.xls)|*.xls";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    path = dialog.FileName;
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

        /// <summary>
        /// 初始化模糊集合
        /// </summary>
        public void InitVagueList()
        {
            var keywordArr = tbKeywords.Text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            vagueList = new List<string>(keywordArr);
        }

        /// <summary>
        /// 初始化精确集合
        /// </summary>
        public void InitAccurateList(OleDbConnection oleDbConn)
        {
            accurateList = new List<string>();
            string dbCmdStr = string.Empty;
            string selectStr = "SELECT " + SHCK + " FROM {0} ";
            string groupStr = string.Format(" GROUP BY {0} ", SHCK);
            string whereStr = " WHERE 1=1 ";
            foreach (string itme in vagueList)
            {
                whereStr += string.Format(" AND {0} NOT LIKE '%{1}%' ", SHCK, itme);
                whereStr += string.Format(" AND {0} NOT LIKE '%{1}' ", SHCK, itme);
                whereStr += string.Format(" AND {0} NOT LIKE '%{1}' ", SHCK, itme);
            }
            dbCmdStr = String.Join(" ", selectStr, whereStr, groupStr);
            DataTable oleDt = GetContentsByExcelFile(dbCmdStr, oleDbConn);
            if (oleDt.Rows.Count > 0)
            {
                //ArrayList attributeNames = new ArrayList();
                //for (int i = 0; i < oleDt.Columns.Count; i++)
                //{
                //    string columnName = oleDt.Columns[i].ColumnName;
                //    if (!attributeNames.Contains(columnName))
                //        attributeNames.Add(columnName);
                //}

                foreach (DataRow row in oleDt.Rows)
                {
                    for (int i = 0; i < oleDt.Columns.Count; i++)
                    {
                        string value = row[i].ToString();
                        if (!accurateList.Contains(value))
                        {
                            accurateList.Add(value);
                        }
                    }
                }
            }
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
