using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Windows.Forms;
using System.Data;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;

namespace Excel2Txt
{
    public class CommonHelper
    {
        public static Dictionary<string, string> Options = new Dictionary<string, string>();

        /// <summary>
        /// 读取Excel文件
        /// </summary>
        public static DataTable ReadFileByNPOI(string path, string sheetName)
        {
            DataTable dt = new DataTable();
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                IWorkbook hwb = new HSSFWorkbook(fs);
                //string sheetName = ConfigurationManager.AppSettings["sheetName"];
                var sheet = hwb.GetSheet(sheetName);

                for (int i = 0; i < sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    var row = sheet.GetRow(i);
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        var cell = row.GetCell(j);
                        if (i == 0)
                        {
                            //列名
                            dt.Columns.Add(cell.ToString());
                        }
                        else
                        {
                            if (cell == null)
                            {
                                dr[j] = null;
                            }
                            else
                            {
                                dr[j] = new DataColumn(cell.ToString());
                            }
                        }
                    }
                    dt.Rows.Add(dr);
                }

            }
            return dt;
        }

        public static List<string> GetSheets(string path)
        {
            List<string> list = new List<string>();
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                IWorkbook hwb = new HSSFWorkbook(fs);
                int sheetCount = hwb.NumberOfSheets;
                for (int i = 0; i < sheetCount; i++)
                {
                    list.Add(hwb.GetSheetAt(i).SheetName);
                }
            }
            return list;
        }

        /// <summary>
        /// 弹出对话框，选择文件
        /// </summary>
        public static string OpenDialogAndSelectFile()
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
        /// 从配置中读取分类选项
        /// </summary>
        public static List<string> ReadAppSettings()
        {
            List<string> result = new List<string>();
            int i = 0;
            var list = ConfigurationManager.AppSettings;
            string[] keys = list.AllKeys;
            foreach (var item in keys)
            {
                if (item.StartsWith("Op_"))
                {
                    result.Add(item.TrimStart('O', 'p', '_'));
                    //先存起来，不必下次再读取
                    Options.Add(item.TrimStart('O', 'p', '_'), list[item]);
                    i++;
                }
            }
            return result;
        }
    }
}
