using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;

namespace WPFExportExcelDemo
{
    public class ExportHelper
    {
        /// <summary>
        /// 利用反射导出指定的Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="lst"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool ExportToExcel<T>(List<T> lst, string fileName) where T : ExcelModel
        {
            try
            {
                FileStream fs;
                StreamWriter sw;
                string? data = null;
                var type = lst?.FirstOrDefault()?.GetType();
                Dictionary<string, ExcelExportAttribute> dic = new Dictionary<string, ExcelExportAttribute>();
                foreach (PropertyInfo item in type.GetProperties())
                {
                    var attrBase = item.GetCustomAttributes(typeof(ExcelExportAttribute), true).FirstOrDefault();
                    var attrIgnore = item.GetCustomAttributes(typeof(ExcelExportIgnoreAttribute), true).FirstOrDefault();
                    if (attrBase == null || (attrIgnore != null && (attrIgnore as ExcelExportIgnoreAttribute).Ignore)) continue;
                    dic.Add(item.Name, (attrBase as ExcelExportAttribute));
                }

                //判断文件是否存在,存在就不再次写入列名
                if (!File.Exists(fileName))
                {
                    fs = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                    sw = new StreamWriter(fs, Encoding.UTF8);

                    //添加表头
                    for (int col = 0; col < dic.Count; col++)
                    {
                        data += dic.ElementAt(col).Value.Caption;
                        if (col < dic.Count - 1)
                        {
                            data += ",";//中间用，隔开
                        }
                    }
                    sw.WriteLine(data);
                }
                else
                {
                    fs = new FileStream(fileName, FileMode.Append, FileAccess.Write);
                    sw = new StreamWriter(fs, Encoding.UTF8);
                }

                //写出各行数据
                for (var index = 0; index < lst?.Count; index++)
                {
                    data = null;
                    var obj = lst[index];
                    for (var col = 0; col < dic.Count; col++)
                    {
                        var item = dic.ElementAt(col);
                        var property = type.GetProperty(item.Key);
                        var value = property?.GetValue(obj, null);
                        data += value?.ToString();
                        if (col < dic.Count - 1)
                        {
                            data += ",";//中间用，隔开
                        }
                    }
                    sw.WriteLine(data);
                }
                sw.Close();
                fs.Close();
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }
    }
}
