using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelCompare
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public IList<dynamic> People { get; }
        public MainWindow()
        {
            People = new List<dynamic>();
            for (int i = 0; i < 100; i++)
            {
                People.Add(new Person { ID = i, FirstName = $"FirstName{i}", LastName = $"FirstName{i}", DOB = Guid.NewGuid().ToString() });
            }
            InitializeComponent();

            //DataGrid1.Columns[0].Header = "新标题";
        }


    }

    public class Person
    {
        public int ID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string DOB { get; set; }
    }
}
