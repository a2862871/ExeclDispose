using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ExeclDispose.BLL;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace ExeclDispose
{
    public partial class Form1 : Form
    {
        private bool isStart = false;
        private string pathologyRegisterFormPath = System.Environment.CurrentDirectory + @"\筛选合并\汇总表.xls";
        private string resultFormPath = System.Environment.CurrentDirectory+@"\内镜登记表.xlsx";

        private IWorkbook pathologyworkbook = null;
        private FileStream pathologyfs = null;

        private IWorkbook resultworkbook = null;
        private FileStream resultfs = null;

        public Form1()
        {
            InitializeComponent();
            this.Text = "内镜自动病理" + Application.ProductVersion;
            pathologyRegisterFormPath =  IniHelper.Read("表路径设置", "病理登记表", System.Environment.CurrentDirectory + @"\筛选合并\汇总表.xls");
            resultFormPath = IniHelper.Read("表路径设置", "结果报告表", System.Environment.CurrentDirectory + @"\内镜登记表.xlsx");
            if (File.Exists(pathologyRegisterFormPath)) 
            {
                textBox1.Text = pathologyRegisterFormPath;
            }
            if (File.Exists(resultFormPath))
            {
                textBox2.Text = resultFormPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //StartExeclDispose();

            if (isStart) 
            {
                return;
            }
            if (!OpenExeclFile())
            {
                return;
            }
            isStart = true;
            richTextBox1.Clear();
            Thread thread = new Thread(new ThreadStart(StartExeclDispose));
            thread.IsBackground = true;
            thread.Start();
        }
        private void ShowLog(string msg) 
        {
            this.BeginInvoke(new Action(() =>
            {
                richTextBox1.AppendText(System.Environment.NewLine + msg);
                richTextBox1.ScrollToCaret();
            })
            );
        }
        private bool OpenExeclFile() 
        {
            if(!File.Exists(pathologyRegisterFormPath)|| !File.Exists(resultFormPath)) 
            {
                MessageBox.Show("没有找到Execl文件,请选择正确的文件路径");
                return false;
            }
            try 
            {
                resultfs = new FileStream(resultFormPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                if (resultFormPath.IndexOf(".xlsx") > 0) // 2007版本
                {
                    resultworkbook = new XSSFWorkbook(resultfs);
                }
                else if (resultFormPath.IndexOf(".xls") > 0) // 2003版本
                {

                    resultworkbook = new HSSFWorkbook(resultfs);
                }
                resultfs.Close();

                pathologyfs = new FileStream(pathologyRegisterFormPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                if (pathologyRegisterFormPath.IndexOf(".xlsx") > 0) // 2007版本
                {
                    pathologyworkbook = new XSSFWorkbook(pathologyfs);
                }
                else if (pathologyRegisterFormPath.IndexOf(".xls") > 0) // 2003版本
                {

                    pathologyworkbook = new HSSFWorkbook(pathologyfs);
                }
                pathologyfs.Close();

            } 
            catch (Exception ex) 
            {
                MessageBox.Show("文件正在被使用，无法打开。请关闭后再试。");
                if(resultfs != null) 
                    resultfs.Close();
                if (pathologyfs != null)
                    pathologyfs.Close();
                if (resultworkbook != null)
                    resultworkbook.Close();
                if (pathologyworkbook != null)
                    pathologyworkbook.Close();
                return false;
            }
            return true;
        }
        private void StartExeclDispose()
        {
            try 
            {
                int Count = 0;
                var sheet = resultworkbook.GetSheetAt(0);
                int RowCount = sheet.LastRowNum;
                for (int i = 1; i < RowCount; i++)
                {
                    var Row = sheet.GetRow(i);
                    if (Row == null)
                    {
                        continue;
                    }
                    var cellname = Row.GetCell(2);
                    if (cellname == null)
                    {
                        continue;
                    }
                    string name = cellname.StringCellValue;
                    if (name.Length == 0)
                    {
                        continue;
                    }
                    var cell = Row.GetCell(10);
                    if (cell == null)
                    {
                        Row.CreateCell(10);
                        cell = Row.GetCell(10);
                    }
                    string str = cell.StringCellValue;
                    if (str.Length == 0)
                    {
                        string ret = FindNameToReport(name.Replace(" ",string.Empty));
                        cell.SetCellValue(ret);
                        if (ret.Length == 0)
                        {
                            ShowLog(name + "未找到相关数据");
                        }
                        else
                        {
                            ShowLog(name + "匹配到数据，匹配结果" + ret);
                            Count++;
                        }
                    }
                }
                FileStream fs2 = System.IO.File.Create(resultFormPath);
                resultworkbook.Write(fs2);
                resultworkbook.Close();
                pathologyworkbook.Close();
                fs2.Close();
                isStart = false;
                ShowLog("本次整理结束");
                ShowLog("本次共匹配" + Count + "条病例结果。");
            } 
            catch (Exception ex) 
            {
                ShowLog("程序发生异常，异常原因:"+ex.ToString());
                MessageBox.Show("程序发生异常，请联系管理员处理");
            }
            
        }

        private string FindNameToReport(string tname) 
        {
            try 
            {
                string currentReport = "";
                string retReport = "";
                var sheet = pathologyworkbook.GetSheetAt(0);
                int RowCount = sheet.LastRowNum;
                for (int i = 0; i < RowCount; i++)
                {
                    var Row = sheet.GetRow(i);
                    if (Row == null)
                    {
                        continue;
                    }
                    var cellname = Row.GetCell(2);
                    string name = cellname.StringCellValue;
                    if (name.Equals(tname))
                    {
                        var cell = Row.GetCell(13);
                        if (!cell.StringCellValue.Equals(currentReport)) 
                        {
                            retReport += cell.StringCellValue+"\n";
                            currentReport = cell.StringCellValue;
                        }
                    }
                }
                return retReport;
            }
            catch (Exception ex)
            {
                ShowLog("程序发生异常，异常原因:" + ex.ToString());
                MessageBox.Show("程序发生异常，请联系管理员处理");
                return "";
            }
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdFile;
            ofdFile = new OpenFileDialog();
            ofdFile.Filter = "工作表|*.xlsx;*.xlsm;*.xlsb;*.xls";
            //ofdFile.Filter = "所有文件|*.*|照片文件|*.jpg|语音文件|*.wav";
            if (ofdFile.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofdFile.FileName;
                pathologyRegisterFormPath = ofdFile.FileName;
                IniHelper.Write("表路径设置", "病理登记表", ofdFile.FileName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdFile;
            ofdFile = new OpenFileDialog();
            ofdFile.Filter = "工作表|*.xlsx;*.xlsm;*.xlsb;*.xls";
            //ofdFile.Filter = "所有文件|*.*|照片文件|*.jpg|语音文件|*.wav";
            if (ofdFile.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = ofdFile.FileName;
                resultFormPath = ofdFile.FileName;
                IniHelper.Write("表路径设置", "结果报告表", ofdFile.FileName);
            }
        }
    }
}
