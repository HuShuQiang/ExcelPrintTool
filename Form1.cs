using Seagull.BarTender.Print;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel打印辅助工具
{
    public partial class Form1 : Form
    {
        // 选择的文件路径
        private string _ExcelPath;
        private string _BtwPath;

        private int TotalNum = 0;
        private Engine btEngine;
        private LabelFormatDocument btFormat;
        public Form1()
        {
            InitializeComponent();
            Text = "Excel打印辅助工具 V1.3";
            PrintBtn1.Click += PrintAction1;
            PrintBtn2.Click += PrintAction2;
        }

        private void btnChooseExcel_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "选择 Excel 文件";
                dlg.Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*";
                dlg.Multiselect = false;
                dlg.CheckFileExists = true;
                dlg.CheckPathExists = true;
                if (!string.IsNullOrEmpty(_ExcelPath) && File.Exists(_ExcelPath))
                {
                    dlg.InitialDirectory = Path.GetDirectoryName(_ExcelPath);
                }

                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    _ExcelPath = dlg.FileName;
                    ExcelPath.Text = _ExcelPath;
                    Properties.Settings.Default.ExcelPath = _ExcelPath;
                    Properties.Settings.Default.Save();

                    Tips.Text = "";
                    TotalNum = 0;
                    Properties.Settings.Default.TotalNum = TotalNum;
                    Properties.Settings.Default.Save();
                    PrintedLineNum.Text = "0";
                    PrintedNum.Text = "0";
                }
            }
        }

        private void btnChooseBartender_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "选择 Bartender 模板";
                dlg.Filter = "Bartender 模板 (*.btw)|*.btw|所有文件 (*.*)|*.*";
                dlg.Multiselect = false;
                dlg.CheckFileExists = true;
                dlg.CheckPathExists = true;
                if (!string.IsNullOrEmpty(_BtwPath) && File.Exists(_BtwPath))
                {
                    dlg.InitialDirectory = Path.GetDirectoryName(_BtwPath);
                }

                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    _BtwPath = dlg.FileName;
                    BartenderPath.Text = _BtwPath;
                    Properties.Settings.Default.BtwPath = _BtwPath;
                    Properties.Settings.Default.Save();
                }
            }
        }

        private async Task<bool> BartenderPrint(List<string> Datas)
        {
            bool res = false;
            try
            {

                if (btEngine == null)
                {
                    btEngine = new Engine();
                }

                bool isAlive = btEngine.IsAlive;
                if (!isAlive)
                {
                    btEngine.Start();
                }

                await Task.Run(() =>
                {
                    btFormat = btEngine.Documents.Open(_BtwPath);

                    btFormat.PrintSetup.PrinterName = Printer.Text;//调用打印名称
                    btFormat.PrintSetup.IdenticalCopiesOfLabel = 1; //打印份数
                    for (int x = 0; x < Datas.Count(); x++)
                    {
                        int num = x + 1;
                        string name = PrintField.Text + num.ToString();
                        btFormat.SubStrings[name].Value = Datas[x];//修改条码数据源信息
                    }

                    Result result = btFormat.Print();
                    if (result == Result.Success)
                    {
                        res = true;
                    }

                    btFormat.PrintSetup.Cache.FlushInterval = CacheFlushInterval.PerSession;
                    btFormat.Close(SaveOptions.DoNotSaveChanges);//不保存对打开模板的修改
                    btEngine.Stop();
                });






            }
            catch (Exception ex)
            {
                ShowTip(ex.Message);
            }
            return res;
        }

        private async void PrintAction1(object sender, EventArgs e)
        {
            Tips.Text = "";
            BtnEnable(false);
            int printnum = 0;
            for (int x = 0; x < int.Parse(PrintTagNum.Text); x++)
            {
                int skip = int.Parse(PrintedLineNum.Text);
                int take = int.Parse(DataNum.Text);
                List<string> data = MiniExcelColumnReader.ReadColumnSliceAsList(_ExcelPath, PrintField.Text, skip, take);

                bool printResult = await BartenderPrint(data);
                if (printResult)
                {
                    ChangeNum();
                    printnum++;
                }
            }
            ShowTip($"打印结束！{printnum}张成功，{int.Parse(PrintTagNum.Text) - printnum}张失败");
            BtnEnable(true);
        }

        private async void PrintAction2(object sender, EventArgs e)
        {
            Tips.Text = "";
            BtnEnable(false);
            int printnum = 0;
            int start = int.Parse(StartNum.Text) - 1;
            int end = int.Parse(EndNum.Text);
            int num = int.Parse(DataNum.Text);
            for (int x = start; x < end; x++)
            {
                List<string> data = MiniExcelColumnReader.ReadColumnSliceAsList(_ExcelPath, PrintField.Text, x * num, num);
                bool printResult = await BartenderPrint(data);
                if (printResult)
                {
                    ChangeNum();
                    printnum++;
                }
            }
            int totalPrintNum = (end - start);
            ShowTip($"打印结束！{printnum}张成功，{int.Parse(PrintTagNum.Text) - printnum}张失败");

            BtnEnable(true);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _ExcelPath = Properties.Settings.Default.ExcelPath;
            ExcelPath.Text = _ExcelPath;
            _BtwPath = Properties.Settings.Default.BtwPath;
            BartenderPath.Text = _BtwPath;
            foreach (var printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                Printer.Items.Add(printer);
            }
            Printer.Text = Properties.Settings.Default.Printer;
            TotalNum = Properties.Settings.Default.TotalNum;
            PrintedLineNum.Text = "0";
            PrintedNum.Text = "0";
            if (TotalNum > 0)
            {
                ChangeNum();
            }

            PrintField.Text = Properties.Settings.Default.PrintField;
            DataNum.Text = Properties.Settings.Default.DataNum;
            PrintTagNum.Text = Properties.Settings.Default.PrintTagNum;
            StartNum.Text = Properties.Settings.Default.StartNum;
            EndNum.Text = Properties.Settings.Default.EndNum;

        }

        private void ChangeNum()
        {
            TotalNum++;
            Properties.Settings.Default.TotalNum = TotalNum;
            Properties.Settings.Default.Save();
            int num = int.Parse(DataNum.Text);

            PrintedLineNum.Text = ((TotalNum - 1) * num + num).ToString();
            PrintedNum.Text = (TotalNum).ToString();
        }

        private void ChangeLastNum()
        {
            TotalNum++;
            Properties.Settings.Default.TotalNum = TotalNum;
            Properties.Settings.Default.Save();
            int num = int.Parse(DataNum.Text);
            int lastnum = int.Parse(LastNum.Text);

            PrintedLineNum.Text = ((TotalNum - 1) * num + lastnum).ToString();
            PrintedNum.Text = (TotalNum).ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("当前打印总数为：" + TotalNum.ToString() + "，是否重置？", "提示", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                Tips.Text = "";
                TotalNum = 0;
                Properties.Settings.Default.TotalNum = TotalNum;
                Properties.Settings.Default.Save();
                PrintedLineNum.Text = "0";
                PrintedNum.Text = "0";
            }
        }

        private void ShowTip(string tip)
        {
            Invoke(new Action(() =>
            {
                Tips.Text = DateTime.Now.ToString("HH:mm:ss") + " " + tip;
            }));
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (btEngine != null)
            {
                btEngine.Stop();
                btEngine.Dispose();
            }
        }

        private void BtnEnable(bool state)
        {
            Invoke(new Action(() =>
            {
                if (state)
                {
                    PrintBtn1.Enabled = true;
                    PrintBtn2.Enabled = true;
                    PrintBtn3.Enabled = true;
                    ClearBtn.Enabled = true;
                }
                else
                {
                    PrintBtn1.Enabled = false;
                    PrintBtn2.Enabled = false;
                    PrintBtn3.Enabled = false;
                    ClearBtn.Enabled = false;
                }
            }));
        }

        private void Printer_SelectedValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Printer = Printer.Text;
            Properties.Settings.Default.Save();
        }

        private void PrintField_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Properties.Settings.Default.PrintField = PrintField.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void DataNum_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Properties.Settings.Default.DataNum = DataNum.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void PrintTagNum_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Properties.Settings.Default.PrintTagNum = PrintTagNum.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void StartNum_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Properties.Settings.Default.StartNum = StartNum.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void EndNum_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Properties.Settings.Default.EndNum = EndNum.Text;
                Properties.Settings.Default.Save();
            }
        }

        private async void PrintBtn3_Click(object sender, EventArgs e)
        {
            Tips.Text = "";
            BtnEnable(false);
            int lastnum = int.Parse(LastNum.Text);
            int num = int.Parse(DataNum.Text);
            int printednum = int.Parse(PrintedNum.Text);
            List<string> data = MiniExcelColumnReader.ReadColumnSliceAsList(_ExcelPath, PrintField.Text, printednum * num, lastnum);
            bool printResult = await BartenderPrint(data);
            if (printResult)
            {
                ChangeLastNum();
                ShowTip($"尾数贴纸打印成功！");

            }
            else
            {
                ShowTip($"尾数贴纸打印失败！");
            }

            

            BtnEnable(true);
        }
    }
}
