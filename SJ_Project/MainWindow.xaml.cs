using HslCommunication;
using HslCommunication.Profinet.Melsec;
using log4net;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using SJ_Project.Entity;
using SJ_Project.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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
using System.Windows.Threading;

namespace SJ_Project
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

        private ConfigData config;
        private MelsecMcNet plc;
        private DispatcherTimer ShowTimer;
        private DispatcherTimer timer;
        private DispatcherTimer timer1;
        private DataService dataService;
        private OperateResult connect;
        private List<float> flist = new List<float>();
        private List<float> flist1 = new List<float>();
        private bool remark = false;
        private bool sPort = false;
        private int row = 0;
        private int sheetSum = 0;
        private string fileName = null;
        private IWorkbook workbook = null;
        private string Path = "C:\\Datas\\";
        private GwType gwType = GwType.小径;
        private static BitmapImage IFalse = new BitmapImage(new Uri("/Static/01.png", UriKind.Relative));
        private static BitmapImage ITrue = new BitmapImage(new Uri("/Static/02.png", UriKind.Relative));
        private ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public MainWindow()
        {
            InitializeComponent();

            #region 启动时串口最大化显示
            this.WindowState = WindowState.Maximized;
            Rect rc = SystemParameters.WorkArea; //获取工作区大小
            //this.Topmost = true;
            this.Left = 0; //设置位置
            this.Top = 0;
            this.Width = rc.Width;
            this.Height = rc.Height;
            #endregion

            Init();

            plc = new MelsecMcNet(config.PlcIpAddress, config.PlcPort);
            plc.ConnectTimeOut = 10000; //超时时间

            connect = plc.ConnectServer();

            #region PLC连接定时器
            //timer1 = new System.Windows.Threading.DispatcherTimer();
            //timer1.Tick += new EventHandler(ThreadCheck);
            //timer1.Interval = new TimeSpan(0, 0, 0, 5);
            //timer1.Start();
            #endregion

            //CycleDataRead();

            //try
            //{
            //    dataService.ReadData(GwType.小径);
            //    flist.Add(dataService.TransformData("DT10000-0000114.782M"));
            //    flist.Add(dataService.TransformData("DT10000-0000113.282M"));
            //    flist.Add(dataService.TransformData("DT10000+0000234.782M"));
            //    flist.Add(dataService.TransformData("DT10000+0000414.782M"));
            //    xiaolist.ItemsSource = flist;
            //}
            //catch (Exception ex)
            //{
            //    log.Error(ex.Message);
            //}
            //try
            //{
            //    if (row == 0 || row > 40000)
            //    {
            //        if (!System.IO.Directory.Exists(Path))
            //            System.IO.Directory.CreateDirectory(Path);
            //        if (sheetSum == 0 || sheetSum > 250)
            //        {
            //            fileName = DateTime.Now.ToString("yyyyMMdd-HH") + ".xls";
            //            workbook = new HSSFWorkbook();
            //            var sheet = workbook.CreateSheet("Sheet" + sheetSum);
            //            var sheeqt = workbook.GetSheetAt(sheetSum - 1);
            //            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
            //            {
            //                workbook.Write(fs);  //创建test.xls文件。
            //            }
            //            sheetSum = 1;
            //        }
            //        else
            //        {
            //            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\2020-11-03-16.xls"));
            //            var sheet = workbook.CreateSheet("Sheet" + sheetSum);
            //            using (var fs = new FileStream(Path + "\\2020-11-03-16.xls", FileMode.OpenOrCreate))
            //            {
            //                workbook.Write(fs);  //创建test.xls文件。
            //            }
            //            sheetSum += 1;
            //        }
            //    }
            //}catch(Exception ex)
            //{
            //    log.Error(ex.Message);
            //}
        }

        private void Init()
        {
            LoadJsonData();

            #region 时间定时器
            ShowTimer = new System.Windows.Threading.DispatcherTimer();
            ShowTimer.Tick += new EventHandler(ShowTimer1);
            ShowTimer.Interval = new TimeSpan(0, 0, 0, 1);
            ShowTimer.Start();
            #endregion

            dataService = new DataService(config);

            sPort = dataService.GetConnection();

            PortImage.Source = (sPort ? ITrue : IFalse);
        }

        private void CycleDataRead()
        {

            timer = new DispatcherTimer();
            timer.Tick += (s, e) =>
            {
                try
                {
                    // 创建文件
                    if (row == 0 || row > 40000)
                    {
                        if (!System.IO.Directory.Exists(Path))
                            System.IO.Directory.CreateDirectory(Path);
                        if (sheetSum == 0 || sheetSum > 250)
                        {
                            fileName = DateTime.Now.ToString("yyyyMMdd-HHmm") + ".xls";
                            workbook = new HSSFWorkbook();
                            var sheet = workbook.CreateSheet("Sheet" + sheetSum);
                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                            {
                                workbook.Write(fs);  //创建test.xls文件。
                            }
                            sheetSum = 1;
                            row = 0;
                        }
                        else
                        {
                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                            var sheet = workbook.CreateSheet("Sheet" + sheetSum);
                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                            {
                                workbook.Write(fs);  //创建test.xls文件。
                            }
                            sheetSum += 1;
                            row = 0;
                        }
                    }

                    #region 小径读取信号
                    var xiaor = plc.ReadBool("M100");
                    
                    if (xiaor.IsSuccess && xiaor.Content)
                    {
                        if (gwType == GwType.小径)
                        {
                            ErrorInfo.Text = "测量开始！";
                            flist.Clear();
                            //读取数据
                            for (int i = 0; i < 4; i++)
                            {
                                var re = dataService.ReadData(gwType);
                                flist.Add(re);
                                if (i != 3)
                                    Thread.Sleep(config.XiaoJingTime);
                            }
                            log.Info(flist.Count()+"  "+flist.First());
                            // write to excel
                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                            var sheet = workbook.GetSheetAt(sheetSum - 1);
                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                            {
                                sheet.CreateRow(row).CreateCell(0).SetCellValue("小径数据");
                                row += 1;
                                var crow = sheet.CreateRow(row);
                                for (int i = 0; i < flist.Count(); i++)
                                {
                                    crow.CreateCell(i).SetCellValue(flist[i]);
                                }
                                row += 1;
                                workbook.Write(fs);
                            }

                            xiaolist.ItemsSource = null;
                            xiaolist.ItemsSource = flist;
                            xiaolist.Items.Refresh();
                            bool mark = true;
                            flist.ForEach(f =>
                            {
                                if (mark)
                                {
                                    mark = (config.XiaoJingMin <= f && f <= config.XiaoJingMax);
                                }
                            });

                            // 回写PLC
                            if (mark)
                            {
                                plc.Write("M110", true);
                                gwType = GwType.大径活塞高度;
                                xiaoResult.Text = "OK";
                            }
                            else
                            {
                                plc.Write("M120", true);
                                xiaoResult.Text = "NG";
                            }
                        }
                        else
                        {
                            ErrorInfo.Text = "当前测量不应在此量测位置！";
                        }
                    }

                    var xiaoReset = plc.ReadBool("M130");
                    if (xiaoReset.IsSuccess && xiaoReset.Content)
                    {
                        // clear
                        xiaolist.ItemsSource = null;
                        xiaolist.Items.Refresh();
                        xiaoResult.Text = "";
                        ErrorInfo.Text = "";
                    }

                    #endregion

                    #region 大径活塞高度
                    var dar = plc.ReadBool("M101");
                    if (dar.IsSuccess && dar.Content)
                    {
                        if (gwType == GwType.大径活塞高度)
                        {
                            ErrorInfo.Text = "测量开始！";
                            flist.Clear();
                            flist1.Clear();
                            //读取数据
                            for (int i = 0; i < 4; i++)
                            {
                                var re = dataService.ReadData(gwType);
                                flist.Add(re);
                                var re1 = dataService.ReadData(GwType.活塞高度);
                                flist1.Add(re1);

                                if (i != 3)
                                    Thread.Sleep(config.DaJingHuoSaiTime);
                            }
                            // write to excel
                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                            var sheet = workbook.GetSheetAt(sheetSum - 1);
                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                            {
                                sheet.CreateRow(row).CreateCell(0).SetCellValue("大径数据");
                                row += 1;
                                var crow = sheet.CreateRow(row);
                                for (int i = 0; i < flist.Count(); i++)
                                {
                                    crow.CreateCell(i).SetCellValue(flist[i]);
                                }
                                row += 1;

                                sheet.CreateRow(row).CreateCell(0).SetCellValue("活塞高度数据");
                                row += 1;
                                var crow1 = sheet.CreateRow(row);
                                for (int i = 0; i < flist1.Count(); i++)
                                {
                                    crow.CreateCell(i).SetCellValue(flist1[i]);
                                }
                                row += 1;

                                workbook.Write(fs);
                            }

                            dalist.ItemsSource = null;
                            huolist.ItemsSource = null;
                            dalist.ItemsSource = flist;
                            huolist.ItemsSource = flist1;
                            dalist.Items.Refresh();
                            huolist.Items.Refresh();
                            bool mark = true;
                            flist.ForEach(f =>
                            {
                                if (mark)
                                {
                                    mark = (config.DaJingMin <= f && f <= config.DaJingMax);
                                }
                            });
                            bool mark1 = true;
                            flist1.ForEach(f =>
                            {
                                if (mark1)
                                {
                                    mark1 = (config.HuoSaiMin <= f && f <= config.HuoSaiMax);
                                }
                            });

                            // 回写PLC
                            if (mark && mark1)
                            {
                                plc.Write("M111", true);
                                gwType = GwType.槽径;
                                daResult.Text = "OK";
                                huoResult.Text = "OK";
                            }
                            else
                            {
                                plc.Write("M121", true);
                                daResult.Text = "NG";
                                huoResult.Text = "NG";
                            }
                        }
                        else
                        {
                            ErrorInfo.Text = "当前测量不应在此量测位置！";
                        }
                    }

                    var daReset = plc.ReadBool("M131");
                    if (daReset.IsSuccess && daReset.Content)
                    {
                        // clear
                        dalist.ItemsSource = null;
                        dalist.Items.Refresh();
                        daResult.Text = "";
                        huolist.ItemsSource = null;
                        huolist.Items.Refresh();
                        huoResult.Text = "";
                        ErrorInfo.Text = "";
                    }
                    #endregion

                    #region 槽径
                    var caojr = plc.ReadBool("M102");
                    if (caojr.IsSuccess && caojr.Content)
                    {
                        if (gwType == GwType.槽径)
                        {
                            ErrorInfo.Text = "测量开始！";
                            flist.Clear();
                            //读取数据
                            for (int i = 0; i < 4; i++)
                            {
                                var re = dataService.ReadData(gwType);
                                flist.Add(re);
                                if (i != 3)
                                    Thread.Sleep(config.CaoJingTime);
                            }
                            // write to excel
                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                            var sheet = workbook.GetSheetAt(sheetSum - 1);
                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                            {
                                sheet.CreateRow(row).CreateCell(0).SetCellValue("槽径数据");
                                row += 1;
                                var crow = sheet.CreateRow(row);
                                for (int i = 0; i < flist.Count(); i++)
                                {
                                    crow.CreateCell(i).SetCellValue(flist[i]);
                                }
                                row += 1;
                                workbook.Write(fs);
                            }

                            caojlist.ItemsSource = null;
                            caojlist.ItemsSource = flist;
                            caojlist.Items.Refresh();
                            bool mark = true;
                            flist.ForEach(f =>
                            {
                                if (mark)
                                {
                                    mark = (config.CaoJingMin <= f && f <= config.CaoJingMax);
                                }
                            });

                            // 回写PLC
                            if (mark)
                            {
                                plc.Write("M112", true);
                                gwType = GwType.BUSH;
                                caojResult.Text = "OK";
                            }
                            else
                            {
                                plc.Write("M122", true);
                                caojResult.Text = "NG";
                            }
                        }
                        else
                        {
                            ErrorInfo.Text = "当前测量不应在此量测位置！";
                        }
                    }

                    var caojReset = plc.ReadBool("M132");
                    if (caojReset.IsSuccess && caojReset.Content)
                    {
                        // clear
                        caojlist.ItemsSource = null;
                        caojlist.Items.Refresh();
                        caojResult.Text = "";
                        ErrorInfo.Text = "";
                    }
                    #endregion

                    #region BUSH
                    var bushr = plc.ReadBool("M103");
                    if (bushr.IsSuccess && bushr.Content)
                    {
                        if (gwType == GwType.BUSH)
                        {
                            ErrorInfo.Text = "测量开始！";
                            flist.Clear();
                            //读取数据
                            for (int i = 0; i < 8; i++)
                            {
                                var re = dataService.ReadData(gwType);
                                flist.Add(re);
                                if (i != 7)
                                    Thread.Sleep(config.BushTime);
                            }
                            // write to excel
                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                            var sheet = workbook.GetSheetAt(sheetSum - 1);
                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                            {
                                sheet.CreateRow(row).CreateCell(0).SetCellValue("BUSH数据");
                                row += 1;
                                var crow = sheet.CreateRow(row);
                                for (int i = 0; i < flist.Count(); i++)
                                {
                                    crow.CreateCell(i).SetCellValue(flist[i]);
                                }
                                row += 1;
                                workbook.Write(fs);
                            }

                            bushlist.ItemsSource = null;
                            bushlist.ItemsSource = flist;
                            bushlist.Items.Refresh();
                            bool mark = true;
                            flist.ForEach(f =>
                            {
                                if (mark)
                                {
                                    mark = (config.BushMin <= f && f <= config.BushMax);
                                }
                            });

                            // 回写PLC
                            if (mark)
                            {
                                plc.Write("M113", true);
                                gwType = GwType.槽高;
                                bushResult.Text = "OK";
                            }
                            else
                            {
                                plc.Write("M123", true);
                                bushResult.Text = "NG";
                            }
                        }
                        else
                        {
                            ErrorInfo.Text = "当前测量不应在此量测位置！";
                        }
                    }

                    var bushReset = plc.ReadBool("M133");
                    if (bushReset.IsSuccess && bushReset.Content)
                    {
                        // clear
                        bushlist.ItemsSource = null;
                        bushlist.Items.Refresh();
                        bushResult.Text = "";
                        ErrorInfo.Text = "";
                    }
                    #endregion

                    #region 槽高
                    var caogr = plc.ReadBool("M104");
                    if (caogr.IsSuccess && caogr.Content)
                    {
                        if (gwType == GwType.槽高)
                        {
                            ErrorInfo.Text = "测量开始！";
                            flist.Clear();
                            //读取数据
                            for (int i = 0; i < 4; i++)
                            {
                                var re = dataService.ReadData(gwType);
                                flist.Add(re);
                                if (i != 3)
                                    Thread.Sleep(config.CaoGaoTime);
                            }
                            // write to excel
                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                            var sheet = workbook.GetSheetAt(sheetSum - 1);
                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                            {
                                sheet.CreateRow(row).CreateCell(0).SetCellValue("槽高数据");
                                row += 1;
                                var crow = sheet.CreateRow(row);
                                for (int i = 0; i < flist.Count(); i++)
                                {
                                    crow.CreateCell(i).SetCellValue(flist[i]);
                                }
                                row += 1;
                                workbook.Write(fs);
                            }

                            caogaolist.ItemsSource = null;
                            caogaolist.ItemsSource = flist;
                            caogaolist.Items.Refresh();
                            bool mark = true;
                            flist.ForEach(f =>
                            {
                                if (mark)
                                {
                                    mark = (config.CaoGaoMin <= f && f <= config.CaoGaoMax);
                                }
                            });

                            // 回写PLC
                            if (mark)
                            {
                                plc.Write("M114", true);
                                gwType = GwType.小径;
                                caogaoResult.Text = "OK";
                            }
                            else
                            {
                                plc.Write("M124", true);
                                caogaoResult.Text = "NG";
                            }
                        }
                        else
                        {
                            ErrorInfo.Text = "当前测量不应在此量测位置！";
                        }
                    }

                    var caogReset = plc.ReadBool("M134");
                    if (caogReset.IsSuccess && caogReset.Content)
                    {
                        // clear
                        caogaolist.ItemsSource = null;
                        caogaolist.Items.Refresh();
                        caogaoResult.Text = "";
                        ErrorInfo.Text = "";
                    }
                    #endregion

                    // 移除光电信号
                    var NGSingal = plc.ReadBool("M140");
                    if (NGSingal.IsSuccess && NGSingal.Content)
                    {
                        gwType = GwType.小径;
                        xiaolist.ItemsSource = null;
                        xiaolist.Items.Refresh();
                        xiaoResult.Text = "";
                        dalist.ItemsSource = null;
                        dalist.Items.Refresh();
                        daResult.Text = "";
                        huolist.ItemsSource = null;
                        huolist.Items.Refresh();
                        huoResult.Text = "";
                        caojlist.ItemsSource = null;
                        caojlist.Items.Refresh();
                        caojResult.Text = "";
                        bushlist.ItemsSource = null;
                        bushlist.Items.Refresh();
                        bushResult.Text = "";
                        caogaolist.ItemsSource = null;
                        caogaolist.Items.Refresh();
                        caogaoResult.Text = "";
                        ErrorInfo.Text = "";
                    }

                    remark = true;
                }
                catch (Exception exc)
                {
                    log.Error("------PLC访问出错------");
                    log.Error(gwType.ToString()+"  "+row + "  " + exc.Message);
                    timer.Stop();
                    remark = false;
                }
            };
            timer.Interval = TimeSpan.FromMilliseconds(50);
            timer.Start();
        }

        /// <summary>
        /// 读取本地配置文件
        /// </summary>
        private void LoadJsonData()
        {
            try
            {
                using (var sr = File.OpenText("C:\\config\\SJConfig.json"))
                {
                    string JsonStr = sr.ReadToEnd();
                    if (config == null)
                    {
                        config = JsonConvert.DeserializeObject<ConfigData>(JsonStr);
                    }
                }
            }
            catch (Exception e)
            {
                log.Error(e.Message);
            }

        }

        private void ShowTimer1(object sender, EventArgs e)
        {
            this.TM.Text = " ";
            //获得年月日 
            this.TM.Text += DateTime.Now.ToString("yyyy年MM月dd日");   //yyyy年MM月dd日 
            this.TM.Text += "                  \r\n         ";
            //获得时分秒 
            this.TM.Text += DateTime.Now.ToString("HH:mm:ss");
            this.TM.Text += "              ";
            this.TM.Text += DateTime.Now.ToString("dddd", new System.Globalization.CultureInfo("zh-cn"));
            this.TM.Text += " ";
        }

        public void ThreadCheck(object sender, EventArgs e)
        {
            var check = plc.ReadBool("M100");

            //log.Info("log: "+check.IsSuccess + "   " + check.Content + check.Message);
            if (check.IsSuccess)
            {
                QPLCImage.Source = ITrue;

                if (!remark)
                {
                    CycleDataRead();
                }
            }
            else
            {
                QPLCImage.Source = IFalse;
            }
        }

        /// <summary>
        /// 窗口关闭事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (connect.IsSuccess)
            {
                plc.ConnectClose();
            }
            log.Info("PLC Disconnected!");
            if (timer != null && timer.IsEnabled)
                timer.Stop();
            if (timer1 != null && timer1.IsEnabled)
                timer1.Stop();
            if (ShowTimer != null && ShowTimer.IsEnabled)
                ShowTimer.Stop();

            dataService.Close();
        }

    }
}
