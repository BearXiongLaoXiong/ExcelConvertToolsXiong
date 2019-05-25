﻿using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using System.Collections.Generic;

namespace ExcelConvertToolsXiongSetup
{
    public partial class Form1 : Form
    {

        private string _fileName = "";

        //配置文件Sheet
        private readonly List<Config> _configList = new List<Config>();

        private DataTable _targetTable;
        private DataTable _dataTable;
        public Form1()
        {
            InitializeComponent();
            dataGridView1.RowHeadersWidth = 70;
            dataGridView2.RowHeadersWidth = 70;
            dataGridView3.RowHeadersWidth = 70;
            dataGridView1.RowStateChanged += RowStateChanged;
            dataGridView2.RowStateChanged += RowStateChanged;
            dataGridView3.RowStateChanged += RowStateChanged;
            dataGridView1.CellClick += DataGrid1CellClick;
            dataGridView2.CellClick += DataGrid2CellClick;
            dataGridView3.CellClick += DataGrid3CellClick;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            _configList.Clear();
            _fileName = "";
            textBox1.Text = "";
            _targetTable = null;
            this._dataTable = null;
            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = @"All files (*.*)|*.*|xlsx(*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = false
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK) _fileName = openFileDialog.FileName;

            if (_fileName.Length == 0) return;
            textBox1.Text = _fileName;

            string configFile = Environment.CurrentDirectory + "\\列转换配置.xlsx";
            if (!File.Exists(configFile))
            {
                MessageBox.Show(configFile + @"出现配置文件不存在的致命错误,请恢复配置文件后再操作!\r\n");
                return;
            }
            //try
            //{
            DataTable configTable = ExcelOpenXml.GetSheet(configFile, "Sheet1");
            if (configTable == null || configTable.Rows.Count < 1)
            {
                MessageBox.Show("列转换配置表[Sheet1]表数据不完整,请给出正确格式的配置文件");
                return;
            }
            for (int i = 0; i < configTable.Columns.Count; i++)
                _configList.Add(new Config { Dt1 = configTable.Columns[i].ColumnName.Trim(), Dt2 = configTable.Rows[0][i].ToString().Trim() });

            DataTable configMappingTable = ExcelOpenXml.GetSheet(configFile, "Sheet2");

            //配置文件Sheet2 数据
            var columnsList = configMappingTable.Rows.Cast<DataRow>().Select(x => new { ChargeCurrency = x[0].ToString().Trim(), ChargeCode = x[1].ToString().Trim(), Columns = x[2].ToString().Trim() }).ToList();
            var cofingListToLower = _configList.Select(x => x.Dt1.ToLower());
            var errConfig = columnsList.Where(x => !cofingListToLower.Contains(x.Columns.ToLower())).Select(x => x.Columns);

            if (errConfig.Any())
            {
                MessageBox.Show($"配置文件出现致命错误!!! \r\n 以下表[Sheet2].列[Columns]中的数据未在表[Sheet1]中查找到 \r\n {string.Join("\r\n", errConfig)}");
                return;
            }


            _dataTable = ExcelOpenXml.GetSheet(_fileName, "Sheet0");
            if (_dataTable.Rows.Count < 3)
            {
                MessageBox.Show(@"未读取到数据");
                return;
            }

            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                string row0 = _dataTable.Rows[0][i].ToString().Trim();
                string row1 = _dataTable.Rows[1][i].ToString().Trim();
                if (row0.Length > 0 && row1.Length == 0)
                    _dataTable.Columns[i].ColumnName = row0;
                if (row0.Length > 0 && row1.Length > 0)
                    _dataTable.Columns[i].ColumnName = row1;
                if (row0.Length == 0 && row1.Length > 0)
                    _dataTable.Columns[i].ColumnName = row1;
                _dataTable.Columns[i].ColumnName = _dataTable.Columns[i].ColumnName.Replace("\r", " ").Replace("\n", " ");
            }
            _dataTable.Rows.RemoveAt(0);
            _dataTable.Rows.RemoveAt(0);

            _targetTable = new DataTable("Sheet1");
            _targetTable = configTable.Clone();
            _targetTable.TableName = "Sheet1";
            for (int i = 0; i < _dataTable.Rows.Count; i++)
            {
                DataRow row = _targetTable.NewRow();
                for (int j = 0; j < configTable.Columns.Count; j++)
                {
                    for (int k = 0; k < _dataTable.Columns.Count; k++)
                    {
                        if (_configList[j].Dt2.ToLower() == _dataTable.Columns[k].ColumnName.Trim().ToLower())
                        {
                            //Console.WriteLine(dataTable.Columns[k].ColumnName.Trim().ToLower());
                            row[j] = _dataTable.Rows[i][k];
                        }
                    }
                }
                //  Charge Currency	Charge Code 行列转换
                var targetColumnName = columnsList.FirstOrDefault(x => x.ChargeCurrency.ToLower() == _dataTable.Rows[i]["Charge Currency"].ToString().Trim().ToLower() && x.ChargeCode.ToLower() == _dataTable.Rows[i]["Charge Code"].ToString().ToLower())?.Columns;
                if (_configList.Select(x => x.Dt1).Any(x => x.Contains(targetColumnName ?? "abcdefghigk123")))
                    row[targetColumnName ?? ""] = _dataTable.Rows[i]["Charge Amount"];
                _targetTable.Rows.Add(row);

                //通过BL nr. 列判断 blvposno的计数,从1开始计数，每行++1 直到[BL nr. 列]和上一行不一致后,重新从1开始计数
                if (cofingListToLower.Contains(@"blvposno") && cofingListToLower.Contains("BL nr.".ToLower()))
                {
                    int blvposno = 1;
                    if (i > 0 && i < _dataTable.Rows.Count)
                        blvposno = _targetTable.Rows[i]["BL nr."].ToString().ToLower() == _targetTable.Rows[i - 1]["BL nr."].ToString().ToLower()
                                    ? int.Parse(_targetTable.Rows[i - 1]["blvposno"].ToString()) + 1
                                    : 1;
                    _targetTable.Rows[i]["blvposno"] = blvposno;
                }
            }
            dataGridView1.DataSource = configTable;
            dataGridView2.DataSource = _dataTable;
            dataGridView3.DataSource = _targetTable;

            //}
            //catch (Exception err)
            //{
            //    MessageBox.Show(err.Message);
            //}
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


        private void RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = $"{e.Row.Index + 1}";
        }

        private void DataGrid1CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var columnName = _configList[e.ColumnIndex].Dt2.Trim().ToLower();
            var index = dataGridView2.SelectedCells.Count > 0 ? dataGridView2.SelectedCells[0].RowIndex :
                        dataGridView3.SelectedCells.Count > 0 ? dataGridView3.SelectedCells[0].RowIndex : 0;

            dataGridView3.CurrentCell = dataGridView3.Rows[index].Cells[e.ColumnIndex];
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                if (columnName == _dataTable.Columns[i].ColumnName.Trim().ToLower())
                {
                    dataGridView2.CurrentCell = dataGridView2.Rows[index].Cells[i];
                    return;
                }
                else
                {
                    dataGridView2.CurrentCell = null;
                }
            }
        }

        private void DataGrid2CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var columnName = _dataTable.Columns[e.ColumnIndex].ColumnName;

            for (int i = 0; i < _configList.Count; i++)
            {
                if (columnName.Trim().ToLower() == _configList[i].Dt2.Trim().ToLower())
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[i];
                    dataGridView3.CurrentCell = dataGridView3.Rows[e.RowIndex].Cells[i];
                    return;
                }
                else
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView3.CurrentCell = null;
                }
            }
            //dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[3];
            //var list = _configList.Select(x => x.Dt2);

            //dataGridView3.CurrentCell = dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex];
        }


        private void DataGrid3CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var columnName = _configList[e.ColumnIndex].Dt2.Trim().ToLower();

            dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[e.ColumnIndex];
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                if (columnName == _dataTable.Columns[i].ColumnName.Trim().ToLower())
                {
                    dataGridView2.CurrentCell = dataGridView2.Rows[e.RowIndex].Cells[i];
                    return;
                }
                else
                {
                    dataGridView2.CurrentCell = null;
                }
            }
        }
    }

    public class Config
    {
        public string Dt1 { get; set; }
        public string Dt2 { get; set; }
    }
}
