using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

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
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = @"All files (*.*)|*.*|txt files (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = false;
            if (openFileDialog.ShowDialog((IWin32Window)this) == DialogResult.OK)
                _fileName = openFileDialog.FileName;
            string str1 = Environment.CurrentDirectory + "\\列转换配置.xlsx";
            if (!File.Exists(str1))
            {
                int num1 = (int)MessageBox.Show(str1 + @"出现配置文件不存在的致命错误,请恢复配置文件后再操作!\r\n");
            }
            else
            {
                DataTable sheet1 = ExcelOpenXml.GetSheet(str1, "Sheet1");
                DataTable sheet2 = ExcelOpenXml.GetSheet(_fileName, "Sheet1");
                if (sheet2.Rows.Count < 3)
                {
                    int num2 = (int)MessageBox.Show(@"未读取到数据");
                }
                else
                {
                    for (int index = 0; index < sheet2.Columns.Count; ++index)
                    {
                        string str2 = sheet2.Rows[0][index].ToString().Trim();
                        string str3 = sheet2.Rows[1][index].ToString().Trim();
                        if (str2.Length > 0 && str3.Length == 0)
                            sheet2.Columns[index].ColumnName = str2;
                        if (str2.Length > 0 && str3.Length > 0)
                            sheet2.Columns[index].ColumnName = str3;
                        if (str2.Length == 0 && str3.Length > 0)
                            sheet2.Columns[index].ColumnName = str3;
                    }
                    sheet2.Rows.RemoveAt(0);
                    sheet2.Rows.RemoveAt(0);
                    _targetTable = new DataTable("Sheet1");
                    _targetTable = sheet1.Clone();
                    _targetTable.TableName = "Sheet1";
                    for (int index1 = 0; index1 < sheet2.Rows.Count; ++index1)
                    {
                        DataRow row = _targetTable.NewRow();
                        for (int index2 = 0; index2 < sheet1.Columns.Count; ++index2)
                        {
                            for (int index3 = 0; index3 < sheet2.Columns.Count; ++index3)
                            {
                                sheet1.Rows[0][index2].ToString();
                                string columnName = sheet2.Columns[index3].ColumnName;
                                if (sheet1.Rows[0][index2].ToString().Trim().ToLower() == sheet2.Columns[index3].ColumnName.Trim().ToLower())
                                {
                                    Console.WriteLine(sheet2.Columns[index3].ColumnName.Trim().ToLower());
                                    row[index2] = sheet2.Rows[index1][index3];
                                }
                            }
                        }
                        _targetTable.Rows.Add(row);
                    }
                    dataGridView1.DataSource = sheet1;
                    dataGridView2.DataSource = sheet2;
                    dataGridView3.DataSource = _targetTable;
                }
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            string path = Environment.CurrentDirectory + "\\转换结果";
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            ExcelOpenXml.Create($@"{path}\{Path.GetFileNameWithoutExtension(_fileName)}.{DateTime.Now:yyyy.MM.dd.HH.mm.ss}.xlsx", new DataSet()
            {
                Tables = {_targetTable}
            });
            int num = (int)MessageBox.Show(@"转换完成!");
        }
    }
}
