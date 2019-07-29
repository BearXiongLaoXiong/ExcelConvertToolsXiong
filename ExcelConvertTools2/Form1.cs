using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelConvertTools2
{
    public partial class Form1 : Form
    {
        private string _fileName = "";
        public Form1()
        {
            InitializeComponent();
            dataGridView1.RowHeadersWidth = 70;
            dataGridView2.RowHeadersWidth = 70;
            dataGridView3.RowHeadersWidth = 70;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            _fileName = "";
            textBox1.Text = "";


            var openFileDialog = new OpenFileDialog
            {
                Filter = @"All files (*.*)|*.*|xlsx(*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = false
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK) _fileName = openFileDialog.FileName;

            if (_fileName.Length == 0) return;
            textBox1.Text = _fileName;

            string configFile = Environment.CurrentDirectory + "\\Config.xlsx";
            if (!File.Exists(configFile))
            {
                MessageBox.Show(configFile + @"出现配置文件不存在的致命错误,请恢复配置文件后再操作!\r\n");
                return;
            }

            DataTable configTable = ExcelOpenXml.GetSheet(configFile, "Sheet1");
            if (configTable == null || configTable.Rows.Count < 1)
            {
                MessageBox.Show(@"列转换配置表[Config][Sheet1]表数据不完整,请给出正确格式的配置文件");
                return;
            }

            //1.选择excel
            //2.验证列,不存在则报错
            //3.读取配置文件,不存在则报错
        }

        private void Button2_Click(object sender, EventArgs e)
        {

        }
    }
}
