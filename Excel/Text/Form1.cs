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

namespace Text
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
        private const string src_imei = "imei";
        private const string src_jx = "机型";
        private const string src_imei2 = "imei2";

        //目标表字段
        private const string dst_imei = "imei";
        private const string dst_imei2 = "imei2";


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


                //读取txt
                Dictionary<string, List<string>> dbSource = ReadTxtToHt();
                //写入txt
                WriteTxtToFile(dbSource);
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误！" + ex.Message);
            }
            loadingCtl.Hide();
        }

        private void WriteTxtToFile(Dictionary<string, List<string>> dbSource)
        {
            foreach (string key in dbSource.Keys)
            {
                List<string> db = dbSource[key];
                string sourFileName = Path.GetFileNameWithoutExtension(this.tbFile.Text);
                string fileName = string.Format("{0}{1} {2}.txt", sourFileName, key, DateTime.Now.ToString("yyyyMMdd"));
                string filePath = Path.Combine(this.tbSavePath.Text, fileName);
                using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate))
                {
                    using (StreamWriter writer = new StreamWriter(fs, Encoding.Default))
                    {
                        writer.Write(string.Join("\r\n", db));
                    }
                }
            }
        }

        private Dictionary<string, List<string>> ReadTxtToHt()
        {
            /* * * * * * * * * * * * * * * * *
             * imei                      imei2 *
             * * * * * * * * * * * * * *  * * */
            string dst_Format = "{0},{1}";
            Dictionary<string, List<string>> dbSource = new Dictionary<string, List<string>>();
            using (FileStream fs = new FileStream(this.tbFile.Text, FileMode.OpenOrCreate))
            {
                using (StreamReader reader = new StreamReader(fs, Encoding.UTF8))
                {
                    string dbLine;
                    //处理每一行数据
                    while ((dbLine = reader.ReadLine()) != null)
                    {
                        if (!string.IsNullOrEmpty(dbLine))
                        {
                            var cols = dbLine.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            if (cols[0] == "imei")
                                continue;
                            if (dbSource.Keys.Contains(cols[1]))
                            {
                                dbSource[cols[1]].Add(string.Format(dst_Format, cols[0], cols[2]));
                            }
                            else
                            {
                                dbSource[cols[1]] = new List<string>();
                                dbSource[cols[1]].Add(string.Format(dst_Format, cols[0], cols[2]));
                            }
                        }
                    }
                }
            }
            return dbSource;
        }

        /// <summary>
        /// 弹出对话框，选择文件
        /// </summary>
        public string OpenDialogAndSelectFile()
        {
            string path = "";
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "文本文件(*.txt)|*.txt";
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
