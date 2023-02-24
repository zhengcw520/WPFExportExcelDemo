using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFExportExcelDemo.Model
{
    public class Customer:ExcelModel
    {
		[ExcelExport("客户中文名")]
        public string? Cust_Name { get; set; }

        [ExcelExport("客户英文名")]
        public string? Cust_EName { get; set; }

        [ExcelExport("电话")]
        public string? Cust_Tel { get; set; }

        [ExcelExport("国家")]
        public string? Country { get; set; }

        [ExcelExport("邮件")]
        public string? Fax { get; set; }

        [ExcelExport("地址")]
        public string? Address { get; set; }
    }
}
