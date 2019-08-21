using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelConvertTools2
{
    public partial class Form1 : Form
    {


        //      1.表1，仅保留黄色表头部分
        //        chargeCode的内容变为列
        //        BLLD放在人民币的第一行
        //        根据collection office的勾选的内容来生成表
        //        只选择Incl.OFT.列的内容为N的数据，其他的数据不要
        //        单独加2列。美金合计 人民币合计

        //增加币种选择
        //BLLD不再单独作为一行,合并到其他的行中, LDF 为特殊标记
        //增加HK1,以及和其组合的代码的选项(配置文件)





        /// <summary>
        /// Excel文件路径
        /// </summary>
        private string _fileName = "";
        /// <summary>
        /// 源表格
        /// </summary>
        private DataTable _dataTable;
        /// <summary>
        /// 目标表格,要导出的表格
        /// </summary>
        private DataTable _targetTable;
        private DataTable _configCheckBoxTable;
        private List<ConfigSheet1> _configSheet1List = new List<ConfigSheet1>();
        /// <summary>
        /// 币种配置表格
        /// </summary>
        private List<string> _configCurrencyCheckedList = new List<string>();
        /// <summary>
        /// CollectionOffice配置表格
        /// </summary>
        private List<string> _configCollectionOfficeCheckedList = new List<string>();
        /// <summary>
        /// POD Code配置表格
        /// </summary>
        private List<string> _configPolCodeCheckedList = new List<string>();
        public Form1()
        {
            InitializeComponent();
            CheckedListBoxCurrency1.Items.Clear();
            checkedListBox1.Items.Clear();
            checkedListBoxPolCode1.Items.Clear();

            dataGridView1.ColumnHeadersHeight = 22;
            dataGridView1.RowHeadersWidth = 70;
            dataGridView2.RowHeadersWidth = 70;
            dataGridView3.RowHeadersWidth = 70;
            dataGridView1.RowStateChanged += RowStateChanged;
            dataGridView2.RowStateChanged += RowStateChanged;
            dataGridView3.RowStateChanged += RowStateChanged;


            //根据collection office的勾选的内容来生成表(需要做出配置功能，只显示常用的几个(大约5个),其他的统统不要)
            string configFile = Environment.CurrentDirectory + "\\Config.xlsx";

            var configChargeCurrency = ExcelOpenXml.GetSheet(configFile, "货币种类");
            if (configChargeCurrency == null)
            {
                MessageBox.Show(@"列转换配置表[Config][货币种类]表数据不完整,请给出正确格式的配置文件");
                return;
            }
            foreach (var item in configChargeCurrency.Columns.Cast<DataColumn>().Select(x => x.ColumnName))
                CheckedListBoxCurrency1.Items.Add(item, true);


            _configCheckBoxTable = ExcelOpenXml.GetSheet(configFile, "显示collectionoffice集合");
            if (_configCheckBoxTable == null)
            {
                MessageBox.Show(@"列转换配置表[Config][显示collectionoffice集合]表数据不完整,请给出正确格式的配置文件");
                return;
            }

            foreach (var item in _configCheckBoxTable.Columns.Cast<DataColumn>().Select(x => x.ColumnName))
            {
                checkedListBox1.Items.Add(item, true);
                for (int i = 0; i < _configCheckBoxTable.Rows.Count; i++)
                {
                    var value = _configCheckBoxTable.Rows[i][item]?.ToString()?.Trim() ?? "";
                    if (value.Length > 0) checkedListBoxPolCode1.Items.Add($"{item}-{value}", true);
                }
            }
            //todo:event是在改变前，必须选择其他event
            checkedListBox1.ItemCheck += CheckedListBox1_ItemCheck;
            //checkedListBox1.Items.AddRange(configCheckBoxTable.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray());
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            _fileName = "";
            textBox1.Text = "";
            _configSheet1List.Clear();
            _configCurrencyCheckedList.Clear();
            _configCollectionOfficeCheckedList.Clear();
            _configPolCodeCheckedList.Clear();
            _targetTable = null;
            _dataTable = null;
            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;


            for (var i = 0; i < CheckedListBoxCurrency1.Items.Count; i++)
            {
                var item = CheckedListBoxCurrency1.GetItemChecked(i);
                if (item) _configCurrencyCheckedList.Add(CheckedListBoxCurrency1.GetItemText(CheckedListBoxCurrency1.Items[i]));
            }
            if (_configCurrencyCheckedList.Count == 0)
            {
                MessageBox.Show(@"请至少选择一种[货币]");
                return;
            }

            for (var i = 0; i < checkedListBox1.Items.Count; i++)
            {
                var item = checkedListBox1.GetItemChecked(i);
                if (item) _configCollectionOfficeCheckedList.Add(checkedListBox1.GetItemText(checkedListBox1.Items[i]));
            }

            for (var i = 0; i < checkedListBoxPolCode1.Items.Count; i++)
            {
                var item = checkedListBoxPolCode1.GetItemChecked(i);
                if (item) _configPolCodeCheckedList.Add(checkedListBoxPolCode1.GetItemText(checkedListBoxPolCode1.Items[i]));
            }

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

            DataTable configTable = ExcelOpenXml.GetSheet(configFile, "显示列集合");
            if (configTable == null || configTable.Rows.Count < 1)
            {
                MessageBox.Show(@"列转换配置表[Config][显示列集合]表数据不完整,请给出正确格式的配置文件");
                return;
            }



            //读取配置文件
            for (int i = 0; i < configTable.Columns.Count; i++)
                _configSheet1List.Add(new ConfigSheet1 { ColumnName = configTable.Columns[i].ColumnName.Replace("\r", " ").Replace("\n", " ").Trim(), IsHaveY = configTable.Rows[0][i].ToString().Trim().ToUpper() == "Y" });


            //
            //_dataTable = ExcelOpenXml.GetSheet(_fileName, "Sheet0", 3);
            _dataTable = ExcelOpenXml.GetSheet(_fileName, "Sheet0", 3);
            if (_dataTable.Rows.Count < 3)
            {
                MessageBox.Show(@"未读取到[Sheet0]数据");
                return;
            }
            //Incl. OFT.
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
                _dataTable.Columns[i].ColumnName = _dataTable.Columns[i].ColumnName.Replace("\r", " ").Replace("\n", " ").Trim();
            }
            _dataTable.Rows.RemoveAt(0);
            _dataTable.Rows.RemoveAt(0);


            var collectionOfficeDict = _configCollectionOfficeCheckedList?.ToDictionary(x => x, x => new List<string>());
            foreach (var item in _configPolCodeCheckedList)
            {
                var polCodes = item.Split('-');
                if (polCodes.Length >= 2 && collectionOfficeDict.ContainsKey(polCodes[0]))
                    collectionOfficeDict[polCodes[0]].Add(polCodes[1]);
            }

            //只有[Incl. OFT.]值为N才计算,其他值的行,则忽略,仅显示collection office勾选的列
            var table = new DataTable("Sheet0");
            table = _dataTable.Clone();
            for (int i = 0; i < _dataTable.Rows.Count; i++)
                if (_dataTable.Rows[i]["Incl. OFT."].ToString().Trim().ToUpper() == "N" && _configCollectionOfficeCheckedList.Contains(_dataTable.Rows[i]["Collection Office"].ToString().Trim()) && _configCurrencyCheckedList.Contains(_dataTable.Rows[i]["Charge Currency"].ToString().Trim()))
                {
                    var key = _dataTable.Rows[i]["Collection Office"].ToString().Trim();
                    var polCode = _dataTable.Rows[i]["POL Code"].ToString().Trim();
                    //if (_dataTable.Rows[i]["B/L No."].ToString().Trim() == "NG11900810")
                    //{
                    //    Console.WriteLine(key);
                    //}
                    if (!collectionOfficeDict.ContainsKey(key)) continue;
                    if (collectionOfficeDict[key].Count > 0)
                    {
                        if (collectionOfficeDict[key].Contains(polCode))
                            table.Rows.Add(_dataTable.Rows[i].ItemArray);
                    }
                    else
                        table.Rows.Add(_dataTable.Rows[i].ItemArray);
                }
            _dataTable = table;

            if (_dataTable.Rows.Count < 1)
            {
                MessageBox.Show(@"当前条件下无数据");
                return;
            }

            //表1，仅保留黄色表头部分,仅保留配置Sheet0中值为Y的列
            var list = _configSheet1List.Where(x => x.IsHaveY).Select(x => x.ColumnName.ToUpper()).ToList();
            var table1 = new DataTable("Sheet0");
            table1 = table.Clone();

            for (int i = 0; i < table.Rows.Count; i++)
                table1.Rows.Add(table.Rows[i].ItemArray);

            for (int i = table.Columns.Count - 1; i >= 0; i--)
                if (!list.Contains(table1.Columns[i].ColumnName.ToUpper())) table1.Columns.RemoveAt(i);



            _targetTable = table1;



            _targetTable = new DataTable("Sheet1");
            //_targetTable = table1.Clone();
            _targetTable.TableName = "Sheet1";

            //chargeCode的内容变为列
            var chargeCodeUsList = table.Select("[Charge Currency] = 'USD'").Select(x => x.Field<string>("Charge Code"))?.Distinct()?.ToList();
            var chargeCodeCnList = table.Select("[Charge Currency] = 'CNY'").Select(x => x.Field<string>("Charge Code"))?.Distinct()?.ToList();

            var targetTableColumnsNamesList = table1.Columns.Cast<DataColumn>().Select(x => x.ColumnName)?.ToList();
            var targetOldTable = targetTableColumnsNamesList.ToArray();
            if (chargeCodeUsList?.Count + chargeCodeCnList?.Count < 1 || targetTableColumnsNamesList?.Count < 1)
            {
                MessageBox.Show(@"未找到关键列[Charge Code],注:大小写,空字符敏感");
                return;
            }
            //将chargeCode的内容添加至tableColumns
            for (int i = 0; i < targetTableColumnsNamesList.Count; i++)
                if (targetTableColumnsNamesList[i] == "Charge Code")
                {
                    //添加美元列、合计
                    if (chargeCodeUsList.Count > 0) targetTableColumnsNamesList.InsertRange(i + 1, chargeCodeUsList);
                    targetTableColumnsNamesList.Insert(i + 1 + chargeCodeUsList.Count, "Us Sum");

                    //添加RMB列、合计
                    if (chargeCodeCnList.Count > 0) targetTableColumnsNamesList.InsertRange(i + 1 + chargeCodeUsList.Count + 1, chargeCodeCnList);
                    targetTableColumnsNamesList.Insert(i + 1 + chargeCodeUsList.Count + 1 + chargeCodeCnList.Count, "Cn Sum");
                }


            //单独加2列。美金合计 人民币合计
            //targetTableColumnsNamesList.AddRange(new[] { "USD Sum", "CNY Sum" });
            _targetTable.Columns.AddRange(targetTableColumnsNamesList.Select(x => new DataColumn(x)).ToArray());

            //按照BLNo 分组
            var blNoList = table1.Rows.Cast<DataRow>().Select(x => x.Field<string>("B/L No.").Trim())?.Distinct()?.ToList();
            if (blNoList?.Count < 1)
            {
                MessageBox.Show(@"未找到关键数据[B/L No.],请检查数据源!");
                return;
            }


            for (int i = 0; i < blNoList.Count; i++)
            {
                //检索唯一[BL/No]
                var rows = table1.Select($"[B/L No.] = '{blNoList[i]}'");

                var datas = rows.Select(x => new RowModel
                {
                    Code = x.Field<string>("Charge Code"),
                    Currency = x.Field<string>("Charge Currency"),
                    Rated = x.Field<string>("Rated As"),
                    Unit = x.Field<string>("Unit"),
                    Amount = x.Field<string>("Charge Amount"),
                    Us = "",
                    Cn = ""
                })?.ToList();

                //检索唯一[BL/No]下唯一[Unit]
                var units = datas.Select(x => x.Unit).Distinct().ToList();
                //将BLLD移至第一行
                if (units.Contains("BLLD"))
                {
                    units.Remove("BLLD");
                    //units.Insert(0, "BLLD");
                    foreach (var item in datas.Where(x => x.Unit.Contains("BLLD")))
                        item.Unit = units[0];
                }


                foreach (var unit in units)
                {
                    var row = _targetTable.NewRow();
                    foreach (var columnName in targetOldTable)
                    {
                        row[columnName] = rows[0][columnName];
                    }

                    var unitEqualsDatas = datas.Where(x => x.Unit == unit).ToList();

                    foreach (var data in unitEqualsDatas)
                        row[data.Code] = data.Amount;
                    row["Us Sum"] = unitEqualsDatas.Where(x => x.Currency == "USD").Sum(x => Convert.ToDecimal(x.Amount)).ToString();
                    row["Cn Sum"] = unitEqualsDatas.Where(x => x.Currency == "CNY").Sum(x => Convert.ToDecimal(x.Amount)).ToString();
                    row["Rated As"] = unitEqualsDatas.FirstOrDefault()?.Rated;
                    row["Unit"] = unit;
                    row["Charge Amount"] = unitEqualsDatas.Sum(x => Convert.ToDecimal(x.Amount)).ToString();
                    _targetTable.Rows.Add(row);
                }

            }

            dataGridView1.DataSource = configTable;
            dataGridView2.DataSource = _dataTable;
            dataGridView3.DataSource = _targetTable;
            //1.选择excel
            //2.验证列,不存在则报错
            //3.读取配置文件,不存在则报错
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

        private void CheckedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            checkedListBoxPolCode1.Items.Clear();
            _configCollectionOfficeCheckedList.Clear();
            for (var i = 0; i < checkedListBox1.Items.Count; i++)
            {
                var item = checkedListBox1.GetItemChecked(i);
                if (item) _configCollectionOfficeCheckedList.Add(checkedListBox1.GetItemText(checkedListBox1.Items[i]));
            }

            var text = checkedListBox1.GetItemText(checkedListBox1.Items[e.Index]);
            if (e.NewValue == CheckState.Checked)
            {
                if (!_configCollectionOfficeCheckedList.Contains(text)) _configCollectionOfficeCheckedList.Add(text);
            }
            else if (_configCollectionOfficeCheckedList.Contains(text)) _configCollectionOfficeCheckedList.Remove(text);

            foreach (var item in _configCollectionOfficeCheckedList)
            {
                for (int i = 0; i < _configCheckBoxTable.Rows.Count; i++)
                {
                    var value = _configCheckBoxTable.Rows[i][item]?.ToString()?.Trim() ?? "";
                    if (value.Length > 0) checkedListBoxPolCode1.Items.Add($"{item}-{value}", true);
                }
            }
        }
    }

    /// <summary>
    /// Sheet1,配置要保存的列,存在Y则是最后需要保存的列
    /// </summary>
    public class ConfigSheet1
    {
        public string ColumnName { get; set; } = string.Empty;
        public bool IsHaveY { get; set; } = false;
    }

    public class RowModel
    {
        public string Code { get; set; }
        public string Currency { get; set; }
        public string Rated { get; set; }
        public string Unit { get; set; }
        public string Amount { get; set; }
        public string Us { get; set; }
        public string Cn { get; set; }
    }
}