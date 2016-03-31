using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
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
        }

        //excel文件路径
        private string excelFilePath;
        //保存文件路径
        private string saveFilePath;
        //关键字集合，关键字用来做模糊查询分组
        private List<string> keywordList;

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
            /* * * * * * * * * * * * * * * * *
             * 序号  机型  串码  入库日期  备注 *
             * * * * * * * * * * * * * *  * * */
            InitKeywordList();
            string dbCmdStr = string.Empty;
            string selectStr = "SELECT * FROM {0} ";
            string whereStr = " WHERE 1=1 ";
            string groupStr = string.Format(" GROUP BY {0} ", SHCK);
            foreach (string keyword in keywordList)
            {
                whereStr += string.Format(" AND {0} NOT LIKE '%{1}%' ", SHCK, keyword);
                whereStr += string.Format(" AND {0} NOT LIKE '%{1}' ", SHCK, keyword);
                whereStr += string.Format(" AND {0} NOT LIKE '{1}%' ", SHCK, keyword);
            }
            dbCmdStr = String.Join(" ", selectStr, whereStr, groupStr);//排除掉通过关键字帅选的数据，然后分组

            whereStr = " WHERE 1=1 ";
            foreach (string keyword in keywordList)
            {
                whereStr += string.Format(" AND {0} LIKE '%{1}%' ", SHCK, keyword);
                whereStr += string.Format(" AND {0} LIKE '%{1}' ", SHCK, keyword);
                whereStr += string.Format(" AND {0} LIKE '{1}%' ", SHCK, keyword);
                dbCmdStr += " UNION " + String.Join(" ", selectStr, whereStr, groupStr);
            }
            DataTable dt = GetContentsByExcelFile(dbCmdStr);
        }

        /// <summary>
        /// 弹出对话框，选择文件
        /// </summary>
        public string OpenDialogAndSelectFile()
        {
            string path = "";
            using (OpenFileDialog dialog = new OpenFileDialog())
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
        public DataTable GetContentsByExcelFile(string oleDbCmdStr)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFilePath + ";Extended Properties=Excel 8.0;";
            OleDbConnection oleDbConn = new OleDbConnection(strConn);
            OleDbCommand oleDbCmd;
            OleDbDataAdapter oleDbAdp;
            DataTable oleDt = new DataTable();
            DataTable dtTableName = new DataTable();

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
                    oleDbCmd = new OleDbCommand(string.Format(oleDbCmdStr, tableName), oleDbConn);
                    oleDbAdp = new OleDbDataAdapter(oleDbCmd);
                    oleDbAdp.Fill(oleDt);
                    oleDbConn.Close();

                    //if (oleDt.Rows.Count > 0)
                    //{

                    //    ArrayList attributeNames = new ArrayList();
                    //    for (int i = 0; i < oleDt.Columns.Count; i++)
                    //    {
                    //        string columnName = oleDt.Columns[i].ColumnName;
                    //        if (!attributeNames.Contains(columnName))
                    //            attributeNames.Add(columnName);
                    //    }

                    //    foreach (DataRow row in oleDt.Rows)
                    //    {
                    //        for (int i = 0; i < oleDt.Columns.Count; i++)
                    //        {
                    //            string attributeName = attributeNames[i] as string;
                    //            if (!string.IsNullOrEmpty(attributeName))
                    //            {
                    //                string value = row[i].ToString();

                    //            }
                    //        }
                    //    }
                    //}
                }
                catch { }
            }

            return oleDt;
        }

        /// <summary>
        /// 初始化关键字集合
        /// </summary>
        public void InitKeywordList()
        {
            var keywordArr = tbKeywords.Text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            keywordList = new List<string>(keywordArr);
        }
    }
}
