using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Configuration;

namespace Excel2Txt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //加载分类选项
            LoadOptions();
        }

        private Dictionary<string, string> dicOptionsTmp = new Dictionary<string, string>();
        private DataTable dt;
        private string path;

        /// <summary>
        /// 加载分类选项
        /// </summary>
        private void LoadOptions()
        {
            //得到分类选项
            List<string> options = CommonHelper.ReadAppSettings();
            foreach (var item in options)
            {
                listBoxOptions.Items.Add(item);
            }
        }

        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn1_Click(object sender, EventArgs e)
        {
            this.cbSheets.Items.Clear();
            path = "";
            try
            {
                //1.弹窗,选择要转换文件
                path = CommonHelper.OpenDialogAndSelectFile();
                List<string> sheets = CommonHelper.GetSheets(path);
                foreach (var item in sheets)
                {
                    cbSheets.Items.Add(item);
                }
                this.cbSheets.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        /// <summary>
        /// 数据分类
        /// </summary>
        private void FilterData()
        {
            dicOptionsTmp.Clear();
            //得到用户选择分类项
            for (int i = 0; i < listBoxOptions.Items.Count; i++)
            {
                if (listBoxOptions.GetItemChecked(i) && CommonHelper.Options.Keys.Contains(listBoxOptions.Items[i].ToString()))
                {
                    dicOptionsTmp[listBoxOptions.Items[i].ToString()] = CommonHelper.Options[listBoxOptions.Items[i].ToString()];
                    //dicOptionsTmp.Add(listBoxOptions.Items[i].ToString(), CommonHelper.Options[listBoxOptions.Items[i].ToString()]);
                }

            }
            //var options = listBoxOptions.SelectedItems;
            ////用户的选择
            //dicOptionsTmp = new Dictionary<string, string>();
            //foreach (var item in options)
            //{
            //    if (CommonHelper.Options.Keys.Contains(item.ToString()))
            //    {
            //        dicOptionsTmp.Add(item.ToString(), CommonHelper.Options[item.ToString()]);
            //    }
            //}

            //解析数据,生成文件
            TakeData2TxtFile();
        }

        /// <summary>
        /// 按要求生成文件
        /// </summary>
        private void TakeData2TxtFile()
        {
            //文件名
            List<string> listFileName = new List<string>();
            //文件夹名称
            var folderName = "";
            //总路径
            var path = ConfigurationManager.AppSettings["path"];
            //数据个数
            var count = 0;
            //文件的内容
            List<string> listTxt = new List<string>();

            List<string> listKeys = new List<string>(dicOptionsTmp.Keys);
            for (int i = 0; i < listKeys.Count; i++)
            {
                //文件夹名称
                folderName += listKeys[i] + "_";
            }

            if (dt == null || dt.Rows.Count <= 0)
            {
                MessageBox.Show("Excel中没有数据！");
                return;
            }

            if (dicOptionsTmp.Keys.Count == 1)
            {
                //单个分组
                var result = dt.AsEnumerable().GroupBy(x => new { GroupName = x.Field<string>(listKeys[0]) });//dt.Rows.Cast<DataRow>().GroupBy<DataRow, string>(dr => dr[item].ToString());
                int j = 0;
                foreach (var r in result)
                {
                    if (j == 0)
                    {
                        j++;
                        continue;
                    }
                    //每一个分组,生成一个文件
                    count = r.Count();
                    listFileName.Add(r.Key.GroupName + count.ToString() + "台.txt");
                    string fileTxt = "";
                    foreach (var row in r)
                    {
                        //一组中的每一行,只取barcode
                        fileTxt += row["串码"].ToString().Trim() + "\r\n";
                    }
                    listTxt.Add(fileTxt.TrimEnd('\r', '\n'));
                }

            }
            else if (dicOptionsTmp.Keys.Count == 2)
            {
                //双分组
                //var result = dt.AsEnumerable().GroupBy(x => new { GroupName = x.Field<string>(listKeys[0]), GroupName2 = x.Field<string>(listKeys[1]) });

                //var result1 = from d in dt.Rows.Cast<DataRow>().AsEnumerable()
                //              group d by d.Field<string>(listKeys[0])
                //                  into grouped
                //                  select grouped;
                //var result2 = from d in result1.ToList()
                //              from dd in d.AsEnumerable()
                //              group dd by dd.Field<string>(listKeys[1])
                //                  into grouped
                //                  select grouped;

                ////dt.Rows.Cast<DataRow>().GroupBy<DataRow, string>(dr => dr[item].ToString());
                //int j = 0;
                //foreach (var r in result2)
                //{
                //    if (j == 0)
                //    {
                //        j++;
                //        continue;
                //    }
                //    //每一个分组,生成一个文件
                //    count = r.Count();
                //    //listFileName.Add("忻州定制机" + r.Key.GroupName + r.Key.GroupName2 + "(" + count.ToString() + "台)" + DateTime.Now.Year + "." + DateTime.Now.Month + "." + DateTime.Now.Day + ".txt");
                //    listFileName.Add("忻州定制机" + r.Key + "(" + count.ToString() + "台)" + DateTime.Now.Year + "." + DateTime.Now.Month + "." + DateTime.Now.Day + ".txt");
                //    string fileTxt = "";
                //    foreach (var row in r)
                //    {
                //        //一组中的每一行,只取barcode
                //        fileTxt += row["串码"].ToString().Trim() + "\r\n";
                //    }
                //    listTxt.Add(fileTxt.TrimEnd('\r', '\n'));
                //}

                List<DataTable> destination = new List<DataTable>();
                GroupDataRows(dt.Rows.Cast<DataRow>(), destination, new string[] { listKeys[0], listKeys[1] }, 0, dt);
                DataSet dsSuNing = new DataSet();
                DataSet dsZhongDa = new DataSet();
                foreach (DataTable item in destination)
                {
                    if (item.Rows.Count > 0)
                    {
                        //if ((item.Rows[0][listKeys[0]] != null ? item.Rows[0][listKeys[0]].ToString().IndexOf("苏宁") : 0) > 0)
                        //{
                        //    dsSuNing.Merge(item.Copy());
                        //    continue;
                        //}
                        //else if ((item.Rows[0][listKeys[0]] != null ? item.Rows[0][listKeys[0]].ToString().IndexOf("大中") : 0) > 0)
                        //{
                        //    //dtZhongDa = CombineTheSameDatatable(dtZhongDa, item);
                        //    dsZhongDa.Merge(item.Copy());
                        //    continue;
                        //}


                        listFileName.Add(((item.Rows[0] != null ? item.Rows[0][listKeys[0]] : "").ToString().IndexOf("苏宁") > 0 ? "苏宁_" : "") + ((item.Rows[0] != null ? item.Rows[0][listKeys[0]] : "").ToString().IndexOf("大中") > 0 ? "大中_" : "") + DateTime.Now.ToString("yyyyMMdd") + "_步步高vivo_" + (item.Rows[0] != null ? item.Rows[0][listKeys[1]] : "") + "（" + (item.Rows[0] != null ? item.Rows[0][listKeys[0]] : "") + "）_" + item.Rows.Count + (((item.Rows[0] != null ? item.Rows[0][listKeys[0]] : "").ToString().IndexOf("苏宁") + (item.Rows[0] != null ? item.Rows[0][listKeys[0]] : "").ToString().IndexOf("大中")) > 0 ? "" : "_四五星级店面") + ".txt");
                        //每一个表示分好的一个组
                        string fileTxt = "";
                        foreach (DataRow row in item.Rows)
                        {
                            //一组中的每一行,只取barcode
                            fileTxt += row["串码"].ToString().Trim() + "\r\n";
                        }
                        listTxt.Add(fileTxt.TrimEnd('\r', '\n'));
                    }
                }


                //if (dsSuNing.Tables.Count > 0)
                //{
                //    DataTable dtSuNing = dsSuNing.Tables[0];
                //    listFileName.Add(DateTime.Now.ToString("yyyyMMdd") + "_步步高vivo_" + (dtSuNing.Rows[0] != null ? dtSuNing.Rows[0][listKeys[1]] : "") + "（" + (dtSuNing.Rows[0] != null ? dtSuNing.Rows[0][listKeys[0]] : "") + "）_" + dtSuNing.Rows.Count + (((dtSuNing.Rows[0] != null ? dtSuNing.Rows[0][listKeys[0]] : "").ToString().IndexOf("苏宁") + (dtSuNing.Rows[0] != null ? dtSuNing.Rows[0][listKeys[0]] : "").ToString().IndexOf("大中")) > 0 ? "" : "_四五星级店面") + ".txt");
                //    //每一个表示分好的一个组
                //    string fileTxt = "";
                //    foreach (DataRow row in dtSuNing.Rows)
                //    {
                //        //一组中的每一行,只取barcode
                //        fileTxt += row["串码"].ToString().Trim() + "\r\n";
                //    }
                //    listTxt.Add(fileTxt.TrimEnd('\r', '\n'));
                //}

                //if (dsZhongDa.Tables.Count > 0)
                //{
                //    DataTable dtZhongDa = dsZhongDa.Tables[0];
                //    listFileName.Add(DateTime.Now.ToString("yyyyMMdd") + "_步步高vivo_" + (dtZhongDa.Rows[0] != null ? dtZhongDa.Rows[0][listKeys[1]] : "") + "（" + (dtZhongDa.Rows[0] != null ? dtZhongDa.Rows[0][listKeys[0]] : "") + "）_" + dtZhongDa.Rows.Count + (((dtZhongDa.Rows[0] != null ? dtZhongDa.Rows[0][listKeys[0]] : "").ToString().IndexOf("苏宁") + (dtZhongDa.Rows[0] != null ? dtZhongDa.Rows[0][listKeys[0]] : "").ToString().IndexOf("大中")) > 0 ? "" : "_四五星级店面") + ".txt");
                //    //每一个表示分好的一个组
                //    string fileTxt = "";
                //    foreach (DataRow row in dtZhongDa.Rows)
                //    {
                //        //一组中的每一行,只取barcode
                //        fileTxt += row["串码"].ToString().Trim() + "\r\n";
                //    }
                //    listTxt.Add(fileTxt.TrimEnd('\r', '\n'));
                //}


            }
            else
            {
                MessageBox.Show("按照两个以上列名进行分组,正在开发中...");
                return;
            }

            //文件夹的名称
            folderName = folderName.TrimEnd('_');

            //检查文件夹是否存在
            if (!Directory.Exists(path + folderName))
            {
                Directory.CreateDirectory(path + folderName);
            }

            for (int i = 0; i < listFileName.Count; i++)
            {
                //创建文件
                using (StreamWriter writer = new StreamWriter(path + folderName + "/" + listFileName[i]))
                {
                    writer.Write(listTxt[i]);
                }
            }
            MessageBox.Show("条码文件已经创建成功,并存放在：" + ConfigurationManager.AppSettings["path"]);
        }

        private void btn2_Click(object sender, EventArgs e)
        {
            //try
            //{
            if (this.cbSheets.SelectedIndex < 0)
                MessageBox.Show("请选择需要转换的表名！");

            //1.按照用户选择分类,过滤数据
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                //2.读取文件NPOI
                dt = CommonHelper.ReadFileByNPOI(path, this.cbSheets.SelectedItem.ToString());
            }
            FilterData();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

        }

        /// <summary>
        /// DataTable分组
        /// </summary>
        /// <param name="source">数据源(要拆分的数据)类型转换</param>
        /// <param name="destination">分组结果</param>
        /// <param name="groupByFields">分组条件</param>
        /// <param name="fieldIndex">条件个数</param>
        /// <param name="schema">数据源(要拆分的数据)</param>
        public static void GroupDataRows(IEnumerable<DataRow> source, List<DataTable> destination, string[] groupByFields, int fieldIndex, DataTable schema)
        {
            if (fieldIndex >= groupByFields.Length || fieldIndex < 0)
            {
                DataTable dt = schema.Clone();
                foreach (DataRow row in source)
                {
                    DataRow dr = dt.NewRow();
                    dr.ItemArray = row.ItemArray;
                    dt.Rows.Add(dr);
                }
                destination.Add(dt);
                return;
            }

            var results = source.GroupBy(o => o[groupByFields[fieldIndex]]);
            foreach (var rows in results)
            {
                GroupDataRows(rows, destination, groupByFields, fieldIndex + 1, schema);
            }
            fieldIndex++;
        }

        /// <summary>
        /// 合并两个相同的DataTable,返回合并后的结果
        /// </summary>
        /// <param name="dt1"></param>
        /// <param name="dt2"></param>
        /// <returns></returns>
        public DataTable CombineTheSameDatatable(DataTable dt1, DataTable dt2)
        {
            if (dt1.Rows.Count == 0 && dt2.Rows.Count == 0)
            {
                return new DataTable();
            }
            if (dt1.Rows.Count == 0)
            {
                return dt2;
            }
            if (dt2.Rows.Count == 0)
            {
                return dt1;
            }
            DataSet ds = new DataSet();
            ds.Tables.Add(dt1.Copy());
            ds.Merge(dt2.Copy());
            return ds.Tables[0];
        }

    }
}
