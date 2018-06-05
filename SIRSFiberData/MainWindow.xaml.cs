using ExcelDataReader;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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

namespace SIRSFiberData
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public String Text { get; set; }
        public SIRS_fibers_fiber_details SIRS = new SIRS_fibers_fiber_details();

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;

            string filePath = "C:\\Users\\robsmith\\Desktop\\PE14-125-2.2LI_18221062265569.xls";
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    var table = result.Tables[0];
                    var headers = table.Rows[0];
                    var data = table.Rows[1];
                    int idx = 0;
                    foreach (DataColumn column in table.Columns)
                    {
                        switch (headers[column].ToString())
                        {
                            case "Product Name":
                                this.SIRS.ProductName = data[column].ToString();
                                break;
                            case "Serial Number":
                                this.SIRS.SerialNumber = data[column].ToString();
                                break;
                            default:
                                break;
                        }
                        idx++;
                    }
                    //this.Text = result.GetXml();
                }
            }
            string json = JsonConvert.SerializeObject(SIRS, Formatting.Indented);
            this.Text = json;

        }
    }

    public class SIRS_fibers_fiber_details
    {
        public String ProductName { get; set; }
        public String SerialNumber { get; set; }
    }
}
