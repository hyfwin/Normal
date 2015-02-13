using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace StockEv
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        public System.Data.DataTable StockDt,ExportDt = new System.Data.DataTable();


        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.ShowDialog();
            if (ofd.FileName != null)
            {
                if (ofd.FileName.Substring(ofd.FileName.Length - 3, 3) != "xls" && ofd.FileName.Substring(ofd.FileName.Length - 4, 4) != "xlsx")
                {
                    MessageBox.Show("后缀名不对！");
                    return;
                }
                StockDt = ImportExcel.doImport(ofd.FileName, "选股结果").Tables[0];
                CreateDataTable();

                for (int i = 0; i < StockDt.Rows.Count; i++)
                {
                    StockDt.Rows[i]["ProfitProperty"] = Convert.ToDecimal(StockDt.Rows[i]["营业利润"])
                        / Convert.ToDecimal(StockDt.Rows[i]["资产总计"]);
                    StockDt.Rows[i]["ProfitPropertyRank"] = i+1;
                    StockDt.Rows[i]["ProfitValue"] = Convert.ToDecimal(StockDt.Rows[i]["营业利润"])
                        / (Convert.ToDecimal(StockDt.Rows[i]["总市值"])+Convert.ToDecimal(Convert.ToDecimal(StockDt.Rows[i]["负债合计"])));
                    StockDt.Rows[i]["ProfitValueRank"] = i+1;
                }

                //for (int i = 0; i < StockDt.Rows.Count-1; i++)
                //{
                //    for (int j = 0; j < StockDt.Rows.Count - 1 - i; j++)
                //    {
                //        if (Convert.ToDecimal(StockDt.Rows[j]["ProfitProperty"]) < Convert.ToDecimal(StockDt.Rows[j + 1]["ProfitProperty"]))
                //        {
                //            int iChange = Convert.ToInt32(StockDt.Rows[j]["ProfitPropertyRank"]);
                //            StockDt.Rows[j]["ProfitPropertyRank"] = StockDt.Rows[j + 1]["ProfitPropertyRank"];
                //            StockDt.Rows[j + 1]["ProfitPropertyRank"] = iChange;
                //        }
                //    }
                //}

                var q1 = from dt1 in StockDt.AsEnumerable()
                         orderby Convert.ToDecimal(dt1.Field<object>("ProfitProperty")) descending 
                         select dt1;
                int index = 0;
                foreach (var item in q1)//显示查询结果
                {
                    item["ProfitPropertyRank"] = index+1;             
                    index++;
                }

                var q2 = from dt1 in q1 orderby Convert.ToDecimal(dt1.Field<object>("ProfitValue")) descending select dt1;
                index = 0;
                foreach (var item in q2)//显示查询结果
                {
                    item["ProfitValueRank"] = index + 1;
                    index++;
                }

                foreach (var item in q2)//显示查询结果
                {
                    item["TotalRank"] = Convert.ToInt32(item["ProfitValueRank"])
                       + Convert.ToInt32(item["ProfitPropertyRank"]);
                }


                //for (int i = 0; i < StockDt.Rows.Count - 1; i++)
                //{
                //    for (int j = 0; j < StockDt.Rows.Count-1-i; j++)
                //    {
                //        if (Convert.ToDecimal(StockDt.Rows[j]["ProfitValue"]) < Convert.ToDecimal(StockDt.Rows[j+1]["ProfitValue"]))
                //        {
                //            int iChange = Convert.ToInt32(StockDt.Rows[j]["ProfitValueRank"]);
                //            StockDt.Rows[j]["ProfitValueRank"] = StockDt.Rows[j+1]["ProfitValueRank"];
                //            StockDt.Rows[j+1]["ProfitValueRank"] = iChange;
                //        }
                //    }
                //}

                //for (int i = 0; i < StockDt.Rows.Count; i++)
                //{
                //    StockDt.Rows[i]["TotalRank"] = Convert.ToInt32(StockDt.Rows[i]["ProfitValueRank"])
                //        + Convert.ToInt32(StockDt.Rows[i]["ProfitPropertyRank"]);
                //}

                dgStock.ItemsSource = q2.CopyToDataTable().DefaultView;

                ExportDt = q2.CopyToDataTable();

            }
        }


        void CreateDataTable()
        {
            StockDt.Columns.Add("ProfitProperty");
            StockDt.Columns.Add("ProfitPropertyRank");
            StockDt.Columns.Add("ProfitValue");
            StockDt.Columns.Add("ProfitValueRank");
            StockDt.Columns.Add("TotalRank");
        }

        private void btnExport1_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel (*.xlsx)|*.xls"; ;
            if ((bool)(saveFileDialog.ShowDialog()))
            {
                try
                {
                    ImportExcel IE = new ImportExcel();
                    IE.SaveToExcel(saveFileDialog.FileName, ExportDt);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出失败：" + ex.Message);
                }
                MessageBox.Show("导出成功");

            }
        }
    }
}
