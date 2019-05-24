using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Linq;

namespace ExcelConvertToolsXiongSetup
{
    public partial class Form1 : Form
    {

        private string _fileName = "";
        private DataTable _targetTable;
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            _fileName = "";
            _targetTable = null;
            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = @"All files (*.*)|*.*|txt files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = false
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK) _fileName = openFileDialog.FileName;

            string configFile = Environment.CurrentDirectory + "\\列转换配置.xlsx";
            if (!File.Exists(configFile))
            {
                MessageBox.Show(configFile + @"出现配置文件不存在的致命错误,请恢复配置文件后再操作!\r\n");
                return;
            }

            DataTable configTable = ExcelOpenXml.GetSheet(configFile, "Sheet1");
            DataTable configMappingTable = ExcelOpenXml.GetSheet(configFile, "Sheet2");

            //配置文件Sheet1 列
            var cofingList = configTable.Columns.Cast<DataColumn>().Select(x => x.ColumnName.Trim().ToLower()).ToList();
            //配置文件Sheet2 数据
            var columnsList = configMappingTable.Rows.Cast<DataRow>().Select(x => new { ChargeCurrency = x[0].ToString().Trim().ToLower(), ChargeCode = x[1].ToString().Trim().ToLower(), Columns = x[2].ToString().Trim().ToLower() }).ToList();
            var errConfig = columnsList.Where(x => !cofingList.Contains(x.Columns)).Select(x => x.Columns);

            if (errConfig.Any())
            {
                MessageBox.Show($"配置文件出现致命错误!!! \r\n 以下表[Sheet2].列[Columns]中的数据未在表[Sheet1]中查找到 \r\n {string.Join("\r\n", errConfig)}");
                return;
            }


            DataTable dataTable = ExcelOpenXml.GetSheet(_fileName, "Sheet0");
            if (dataTable.Rows.Count < 3)
            {
                MessageBox.Show(@"未读取到数据");
                return;
            }

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                string row0 = dataTable.Rows[0][i].ToString().Trim();
                string row1 = dataTable.Rows[1][i].ToString().Trim();
                if (row0.Length > 0 && row1.Length == 0)
                    dataTable.Columns[i].ColumnName = row0;
                if (row0.Length > 0 && row1.Length > 0)
                    dataTable.Columns[i].ColumnName = row1;
                if (row0.Length == 0 && row1.Length > 0)
                    dataTable.Columns[i].ColumnName = row1;
                dataTable.Columns[i].ColumnName = dataTable.Columns[i].ColumnName.Replace("\r", " ").Replace("\n", " ");
            }
            dataTable.Rows.RemoveAt(0);
            dataTable.Rows.RemoveAt(0);
            _targetTable = new DataTable("Sheet1");
            _targetTable = configTable.Clone();
            _targetTable.TableName = "Sheet1";
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                DataRow row = _targetTable.NewRow();
                for (int j = 0; j < configTable.Columns.Count; j++)
                {
                    for (int k = 0; k < dataTable.Columns.Count; k++)
                    {
                        if (configTable.Rows[0][j].ToString().Trim().ToLower() == dataTable.Columns[k].ColumnName.Trim().ToLower())
                        {
                            Console.WriteLine(dataTable.Columns[k].ColumnName.Trim().ToLower());
                            row[j] = dataTable.Rows[i][k];
                        }
                    }
                }
                var targetColumnName = columnsList.FirstOrDefault(x => x.ChargeCurrency == dataTable.Rows[i]["Charge Currency"].ToString().Trim() && x.ChargeCode == dataTable.Rows[i]["Charge Code"].ToString())?.Columns;
                if (cofingList.Any(x => x.Contains(targetColumnName ?? "abcdefghigk123")))
                    dataTable.Rows[i][targetColumnName ?? ""] = dataTable.Rows[i]["Charge Amount"];
                _targetTable.Rows.Add(row);
            }
            dataGridView1.DataSource = configTable;
            dataGridView2.DataSource = dataTable;
            dataGridView3.DataSource = _targetTable;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            string path = Environment.CurrentDirectory + "\\转换结果";
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            ExcelOpenXml.Create($@"{path}\{Path.GetFileNameWithoutExtension(_fileName)}.{DateTime.Now:yyyy.MM.dd.HH.mm.ss}.xlsx", new DataSet()
            {
                Tables = { _targetTable }
            });
            MessageBox.Show(@"转换完成!");
        }
    }
}
