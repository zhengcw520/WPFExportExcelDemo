using Microsoft.Win32;
using WPFExportExcelDemo.Model;
using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace WPFExportExcelDemo.ViewModels
{
    public class MainWindowViewModel:BindableBase
    {
        private string title="sdsaaddsdf";
        public ObservableCollection<Customer> CustomerInfo { get; set; }

        public string Title
        {
            get { return title; }
            set { title = value; RaisePropertyChanged(); }
        }

        public DelegateCommand ExportCommond { get; set; }

        public MainWindowViewModel()
        {
            CustomerInfo = new ObservableCollection<Customer>();    
            ExportCommond = new DelegateCommand(ExportData);
            for (int i = 0; i < 10; i++)
            {
                CustomerInfo.Add(new Customer { Cust_Name = i + "客户中文", Cust_EName = i + "客户英文",Cust_Tel=i.ToString() });
            }
        }

        private void ExportData()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            //saveFileDialog.Filter = "Excel files (*.csv)|*.csv";
            saveFileDialog.Filter = "Excel files (*.xls)|*.xls";
            saveFileDialog.InitialDirectory = "C:\\Users\\Administrator\\Desktop"; //设置初始目录
            if ((bool)saveFileDialog.ShowDialog())
            {
                string filePath = saveFileDialog.FileName; //获取选择的文件，或者自定义的文件名的全路径。

                List<Customer> customerList = new List<Customer>();
                foreach (var item in CustomerInfo)
                {
                    customerList.Add(item);
                }
                if (customerList != null)
                {
                    ExportHelper.ExportToExcel(customerList, filePath);
                }
            }
        }
    }
}
