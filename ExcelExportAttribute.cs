using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFExportExcelDemo
{
    public class ExcelExportAttribute:Attribute
    {
        public string Caption { get; set; }
        public ExcelExportAttribute(string _Caption)
        {
            Caption = _Caption;
        }
    }

    /// <summary>
    /// 导出自定义特性
    /// </summary>
    public class ExcelExportIgnoreAttribute : Attribute
    {
        public bool Ignore { get; set; }

        public ExcelExportIgnoreAttribute(bool ignore)
        {
            this.Ignore = ignore;
        }
    }
}
