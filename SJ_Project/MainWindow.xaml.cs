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
        private List<float> NgList = new List<float>();
        private bool remark = false;
        private bool sPort = false;
        private int row = -1;
        private int sheetSum = 0;
        private int Number = 0;
        private int xiaoNo = 0;
        private bool xiaoMark = true;
        private int daNo = 0;
        private bool daMark = true;
        private bool huoMark = true;
        private int caojNo = 0;
        private bool caojMark = true;
        private int bushNo = 0;
        private bool bushMark = true;
        private int caogNo = 0;
        private bool caogMark = true;
        private string fileName = null;
        private IWorkbook workbook = null;
        private string Path = null;
        private GwType gwType = GwType.小径;
        private static BitmapImage IFalse = new BitmapImage(new Uri("/Static/01.png", UriKind.Relative));
        private static BitmapImage ITrue = new BitmapImage(new Uri("/Static/02.png", UriKind.Relative));
        private ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private bool xiaomodel = false; // false 为正常模式，true 为标准件模式
        private bool damodel = false;
        private bool huomodel = false;
        private bool bushmodel = false;
        private bool cjmodel = false;
        private bool cgmodel = false;

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
            timer1 = new System.Windows.Threading.DispatcherTimer();
            timer1.Tick += new EventHandler(ThreadCheck);
            timer1.Interval = new TimeSpan(0, 0, 0, 5);
            timer1.Start();
            #endregion

            //CycleDataRead();

            //List<float> ss = new List<float>();
            //for(int i = 0;i<8;i++)
            //{

            //ss.Add(1f);
            //}
            //bushlist.ItemsSource = ss;

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

            xiaomax.Text = "上限:" + config.XiaoJingMax;
            xiaomin.Text = "下限:" + config.XiaoJingMin;
            damax.Text = "上限:" + config.DaJingMax;
            damin.Text = "下限:" + config.DaJingMin;
            huomax.Text = "上限:" + config.HuoSaiMax;
            huomin.Text = "下限:" + config.HuoSaiMin;
            bushmax.Text = "上限:" + config.BushMax;
            bushmin.Text = "下限:" + config.BushMin;
            cjmax.Text = "上限:" + config.CaoJingMax;
            cjmin.Text = "下限:" + config.CaoJingMin;
            cgmax.Text = "上限:" + config.CaoGaoMax;
            cgmin.Text = "下限:" + config.CaoGaoMin;

        }

        private void CycleDataRead()
        {

            timer = new DispatcherTimer();
            timer.Tick += (s, e) =>
            {
                try
                {
                    // 创建文件
                    var date = DateTime.Now.ToString("yyyyMMdd");
                    Path = "C:\\Datas\\" + date;
                    if (!System.IO.Directory.Exists(Path))
                        System.IO.Directory.CreateDirectory(Path);
                    if (row == -1 || row > 40000)
                    {

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
                        if (xiaomodel) // true 为标准件模式 
                        {
                            Thread.Sleep(config.FirstTime);
                            ErrorInfo.Text = "标准件测量开始！";

                            var re = dataService.ReadData(GwType.小径);
                            bool m = (config.XiaoJingMin <= re && re <= config.XiaoJingMax);

                            ShowModelInfo(re, GwType.小径, m);
                        }
                        else
                        {
                            if (gwType == GwType.小径)
                            {
                                //log.Info(xiaoNo);

                                xiaoNo += 1;
                                #region clear
                                if (xiaoNo == 1)
                                {

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
                                    xiaolist.Background = Brushes.SteelBlue;
                                    dalist.Background = Brushes.SteelBlue;
                                    huolist.Background = Brushes.SteelBlue;
                                    caogaolist.Background = Brushes.SteelBlue;
                                    caojlist.Background = Brushes.SteelBlue;
                                    bushlist.Background = Brushes.SteelBlue;

                                    flist.Clear();
                                    flist1.Clear();
                                }
                                #endregion

                                Thread.Sleep(config.FirstTime);
                                ErrorInfo.Text = "测量开始！";
                                if (xiaoMark)
                                {
                                    //读取数据
                                    var re = dataService.ReadData(gwType);
                                    flist.Add(re);

                                    // 每次
                                    xiaoMark = (config.XiaoJingMin <= re && re <= config.XiaoJingMax);
                                    //log.Info(xiaoMark);
                                    if (!xiaoMark)
                                    {
                                        NgList.Add(re);

                                        xiaolist.ItemsSource = null;
                                        xiaolist.ItemsSource = flist;
                                        xiaolist.Background = Brushes.Red;
                                        xiaolist.Items.Refresh();

                                        plc.Write("M120", true);
                                        xiaoResult.Text = "NG";

                                        workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                        var sheet = workbook.GetSheetAt(sheetSum - 1);
                                        using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                        {
                                            Number += 1;
                                            var srow = sheet.CreateRow(row);
                                            srow.CreateCell(0).SetCellValue(Number);
                                            srow.CreateCell(1).SetCellValue("小径数据");
                                            //row += 1;
                                            //var crow = sheet.CreateRow(row);

                                            for (int i = 0; i < flist.Count(); i++)
                                            {
                                                srow.CreateCell(i + 2).SetCellValue(flist[i]);
                                            }
                                            row += 1;
                                            workbook.Write(fs);
                                        }
                                    }
                                    else
                                    {
                                        //log.Info(re + "  " + xiaoNo + "  " + config.XiaoJingCount);

                                        xiaolist.ItemsSource = null;
                                        xiaolist.ItemsSource = flist;
                                        xiaolist.Items.Refresh();
                                        xiaoResult.Text = "OK";

                                        if (xiaoNo == config.XiaoJingCount)
                                        {
                                            plc.Write("M110", true);
                                            gwType = GwType.大径活塞高度;

                                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                            var sheet = workbook.GetSheetAt(sheetSum - 1);
                                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                            {
                                                Number += 1;
                                                var srow = sheet.CreateRow(row);
                                                srow.CreateCell(0).SetCellValue(Number);
                                                srow.CreateCell(1).SetCellValue("小径数据");
                                                //row += 1;
                                                //var crow = sheet.CreateRow(row);
                                                for (int i = 0; i < flist.Count(); i++)
                                                {
                                                    srow.CreateCell(i + 2).SetCellValue(flist[i]);
                                                }
                                                row += 1;
                                                workbook.Write(fs);
                                            }
                                            NgList.Clear();
                                            xiaoNo = 0;
                                        }
                                    }
                                }
                                else
                                {
                                    ErrorInfo.Text = "前次NG!";
                                }

                                #region cancel
                                //for (int i = 0; i < 4; i++)
                                //{
                                //    var re = dataService.ReadData(gwType);
                                //    flist.Add(re);
                                //    if (i != 3)
                                //        Thread.Sleep(config.XiaoJingTime);
                                //}
                                // write to excel
                                //if (xiaoNo == config.XiaoJingCount)
                                //{
                                //    workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                //    var sheet = workbook.GetSheetAt(sheetSum - 1);
                                //    using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                //    {
                                //        sheet.CreateRow(row).CreateCell(0).SetCellValue("小径数据");
                                //        row += 1;
                                //        var crow = sheet.CreateRow(row);
                                //        for (int i = 0; i < flist.Count(); i++)
                                //        {
                                //            crow.CreateCell(i).SetCellValue(flist[i]);
                                //        }
                                //        row += 1;
                                //        workbook.Write(fs);
                                //    }

                                //    xiaolist.ItemsSource = null;
                                //    xiaolist.ItemsSource = flist;
                                //    xiaolist.Items.Refresh();
                                //    bool mark = true;
                                //    flist.ForEach(f =>
                                //    {
                                //        if (mark)
                                //        {
                                //            mark = (config.XiaoJingMin <= f && f <= config.XiaoJingMax);
                                //        }
                                //    });

                                //    // 回写PLC
                                //    if (mark)
                                //    {
                                //        plc.Write("M110", true);
                                //        gwType = GwType.大径活塞高度;
                                //        xiaoResult.Text = "OK";
                                //    }
                                //    else
                                //    {
                                //        plc.Write("M120", true);
                                //        xiaoResult.Text = "NG";
                                //        xiaolist.Background = Brushes.Red;
                                //    }
                                //}
                                #endregion
                            }
                            else
                            {
                                ErrorInfo.Text = $"当前测量应在 {gwType.ToString()} 量测位置！";
                            }
                        }
                    }

                    var xiaoReset = plc.ReadBool("M130");
                    if (xiaoReset.IsSuccess && xiaoReset.Content)
                    {
                        // clear
                        if (gwType == GwType.小径)
                        {
                            xiaolist.ItemsSource = null;
                            xiaolist.Items.Refresh();
                            xiaoResult.Text = "";
                            ErrorInfo.Text = "";
                            xiaolist.Background = Brushes.SteelBlue;
                            xiaoNo = 0;
                            xiaoMark = true;
                        }
                    }

                    #endregion

                    #region 大径活塞高度
                    var dar = plc.ReadBool("M101");
                    if (dar.IsSuccess && dar.Content)
                    {
                        if (damodel || huomodel) // true 为标准件模式 
                        {
                            Thread.Sleep(config.FirstTime);
                            ErrorInfo.Text = "标准件测量开始！";

                            var re = dataService.ReadData(GwType.大径活塞高度);
                            bool m = (config.DaJingMin <= re && re <= config.DaJingMax);
                            var re1 = dataService.ReadData(GwType.活塞高度);
                            bool m1 = (config.HuoSaiMin <= re1 && re1 <= config.HuoSaiMax);

                            ShowModelInfo1(re, re1, GwType.大径活塞高度, m, m1);
                        }
                        else
                        {
                            if (gwType == GwType.大径活塞高度)
                            {
                                daNo += 1;
                                if (daNo == 1)
                                {
                                    flist.Clear();
                                    flist1.Clear();

                                }

                                Thread.Sleep(config.FirstTime);
                                ErrorInfo.Text = "测量开始！";

                                #region 一起
                                if (daMark && huoMark)
                                {
                                    //读取数据
                                    var re = dataService.ReadData(gwType);
                                    flist.Add(re);
                                    var re1 = dataService.ReadData(GwType.活塞高度);
                                    flist1.Add(re1);

                                    daMark = (config.DaJingMin <= re && re <= config.DaJingMax);
                                    huoMark = (config.HuoSaiMin <= re1 && re1 <= config.HuoSaiMax);

                                    if (!daMark || !huoMark)
                                    {
                                        #region mark
                                        if (daMark)
                                        {
                                            daResult.Text = "OK";
                                            dalist.Background = Brushes.SteelBlue;
                                        }
                                        else
                                        {
                                            daResult.Text = "NG";
                                            dalist.Background = Brushes.Red;
                                            NgList.Add(re);

                                        }

                                        if (huoMark)
                                        {
                                            huoResult.Text = "OK";
                                            huolist.Background = Brushes.SteelBlue;
                                        }
                                        else
                                        {
                                            huoResult.Text = "NG";
                                            huolist.Background = Brushes.Red;
                                            NgList.Add(re1);
                                        }

                                        dalist.ItemsSource = null;
                                        dalist.ItemsSource = flist;
                                        //dalist.Background = daMark ? Brushes.SteelBlue : Brushes.Red;
                                        dalist.Items.Refresh();

                                        huolist.ItemsSource = null;
                                        huolist.ItemsSource = flist1;
                                        //huolist.Background = huoMark ? Brushes.SteelBlue : Brushes.Red;
                                        huolist.Items.Refresh();
                                        #endregion

                                        plc.Write("M121", true);
                                        //daResult.Text = daMark ? "OK" : "NG";
                                        //huoResult.Text = huoMark ? "OK" : "NG";


                                        workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                        var sheet = workbook.GetSheetAt(sheetSum - 1);
                                        using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                        {
                                            var crow = sheet.CreateRow(row);
                                            crow.CreateCell(1).SetCellValue("大径数据");
                                            for (int i = 0; i < flist.Count(); i++)
                                            {
                                                crow.CreateCell(i + 2).SetCellValue(flist[i]);
                                            }
                                            row += 1;

                                            var crow1 = sheet.CreateRow(row);
                                            crow1.CreateCell(1).SetCellValue("活塞高度数据");
                                            for (int i = 0; i < flist1.Count(); i++)
                                            {
                                                crow.CreateCell(i + 2).SetCellValue(flist1[i]);
                                            }
                                            row += 1;

                                            workbook.Write(fs);

                                        }

                                    }
                                    else
                                    {
                                        dalist.ItemsSource = null;
                                        huolist.ItemsSource = null;
                                        dalist.ItemsSource = flist;
                                        huolist.ItemsSource = flist1;
                                        dalist.Items.Refresh();
                                        huolist.Items.Refresh();
                                        daResult.Text = "OK";
                                        huoResult.Text = "OK";
                                        if (daNo == config.DaJingHuoSaiCount)
                                        {


                                            plc.Write("M111", true);
                                            gwType = GwType.槽径;

                                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                            var sheet = workbook.GetSheetAt(sheetSum - 1);
                                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                            {
                                                var crow = sheet.CreateRow(row);
                                                crow.CreateCell(1).SetCellValue("大径数据");
                                                for (int i = 0; i < flist.Count(); i++)
                                                {
                                                    crow.CreateCell(i + 2).SetCellValue(flist[i]);
                                                }
                                                row += 1;

                                                var crow1 = sheet.CreateRow(row);
                                                crow1.CreateCell(1).SetCellValue("活塞高度数据");
                                                for (int i = 0; i < flist1.Count(); i++)
                                                {
                                                    crow.CreateCell(i + 2).SetCellValue(flist1[i]);
                                                }
                                                row += 1;

                                                workbook.Write(fs);
                                            }
                                            NgList.Clear();
                                            daNo = 0;
                                        }
                                    }

                                }
                                else
                                {
                                    ErrorInfo.Text = "前次NG!";
                                }
                                #endregion

                                #region DAJING
                                //if (daMark)
                                //{
                                //    //读取数据
                                //    var re = dataService.ReadData(gwType);
                                //    flist.Add(re);
                                //    //var re1 = dataService.ReadData(GwType.活塞高度);
                                //    //flist1.Add(re1);

                                //    daMark = (config.DaJingMin <= re && re <= config.DaJingMax);
                                //    //huoMark = (config.HuoSaiMin <= re1 && re1 <= config.HuoSaiMax);
                                //    if (!daMark)
                                //    {
                                //        #region mark
                                //        if (daMark)
                                //        {
                                //            daResult.Text = "OK";
                                //            dalist.Background = Brushes.SteelBlue;
                                //        }
                                //        else
                                //        {
                                //            daResult.Text = "NG";
                                //            dalist.Background = Brushes.Red;
                                //            NgList.Add(re);

                                //        }

                                //        //if (huoMark)
                                //        //{
                                //        //    huoResult.Text = "OK";
                                //        //    huolist.Background = Brushes.SteelBlue;
                                //        //}
                                //        //else
                                //        //{
                                //        //    huoResult.Text = "NG";
                                //        //    huolist.Background = Brushes.Red;
                                //        //    NgList.Add(re1);
                                //        //}

                                //        dalist.ItemsSource = null;
                                //        dalist.ItemsSource = flist;
                                //        //dalist.Background = daMark ? Brushes.SteelBlue : Brushes.Red;
                                //        dalist.Items.Refresh();

                                //        //huolist.ItemsSource = null;
                                //        //huolist.ItemsSource = flist1;
                                //        ////huolist.Background = huoMark ? Brushes.SteelBlue : Brushes.Red;
                                //        //huolist.Items.Refresh();
                                //        #endregion

                                //        plc.Write("M121", true);
                                //        //daResult.Text = daMark ? "OK" : "NG";
                                //        //huoResult.Text = huoMark ? "OK" : "NG";


                                //        workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                //        var sheet = workbook.GetSheetAt(sheetSum - 1);
                                //        using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                //        {
                                //            var crow = sheet.CreateRow(row);
                                //            crow.CreateCell(1).SetCellValue("大径数据");
                                //            for (int i = 0; i < flist.Count(); i++)
                                //            {
                                //                crow.CreateCell(i + 2).SetCellValue(flist[i]);
                                //            }
                                //            row += 1;

                                //            //var crow1 = sheet.CreateRow(row);
                                //            //crow1.CreateCell(1).SetCellValue("活塞高度数据");
                                //            //for (int i = 0; i < flist1.Count(); i++)
                                //            //{
                                //            //    crow.CreateCell(i + 2).SetCellValue(flist1[i]);
                                //            //}
                                //            //row += 1;

                                //            workbook.Write(fs);

                                //        }

                                //    }
                                //    else
                                //    {
                                //        if (daNo == config.DaJingHuoSaiCount)
                                //        {

                                //            dalist.ItemsSource = null;
                                //            huolist.ItemsSource = null;
                                //            dalist.ItemsSource = flist;
                                //            huolist.ItemsSource = flist1;
                                //            dalist.Items.Refresh();
                                //            huolist.Items.Refresh();

                                //            plc.Write("M111", true);
                                //            gwType = GwType.槽径;
                                //            daResult.Text = "OK";
                                //            //huoResult.Text = "OK";

                                //            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                //            var sheet = workbook.GetSheetAt(sheetSum - 1);
                                //            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                //            {
                                //                var crow = sheet.CreateRow(row);
                                //                crow.CreateCell(1).SetCellValue("大径数据");
                                //                for (int i = 0; i < flist.Count(); i++)
                                //                {
                                //                    crow.CreateCell(i + 2).SetCellValue(flist[i]);
                                //                }
                                //                row += 1;

                                //                //var crow1 = sheet.CreateRow(row);
                                //                //crow1.CreateCell(1).SetCellValue("活塞高度数据");
                                //                //for (int i = 0; i < flist1.Count(); i++)
                                //                //{
                                //                //    crow.CreateCell(i + 2).SetCellValue(flist1[i]);
                                //                //}
                                //                //row += 1;

                                //                workbook.Write(fs);
                                //            }
                                //            NgList.Clear();
                                //        }
                                //    }

                                //}
                                //else
                                //{
                                //    ErrorInfo.Text = "前次NG!";
                                //}
                                #endregion

                                #region cancel
                                //flist.Clear();
                                //flist1.Clear();
                                ////读取数据
                                //for (int i = 0; i < 4; i++)
                                //{
                                //    var re = dataService.ReadData(gwType);
                                //    flist.Add(re);
                                //    var re1 = dataService.ReadData(GwType.活塞高度);
                                //    flist1.Add(re1);

                                //    if (i != 3)
                                //        Thread.Sleep(config.DaJingHuoSaiTime);
                                //}
                                // write to excel
                                //workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                //var sheet = workbook.GetSheetAt(sheetSum - 1);
                                //using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                //{
                                //    sheet.CreateRow(row).CreateCell(0).SetCellValue("大径数据");
                                //    row += 1;
                                //    var crow = sheet.CreateRow(row);
                                //    for (int i = 0; i < flist.Count(); i++)
                                //    {
                                //        crow.CreateCell(i).SetCellValue(flist[i]);
                                //    }
                                //    row += 1;

                                //    sheet.CreateRow(row).CreateCell(0).SetCellValue("活塞高度数据");
                                //    row += 1;
                                //    var crow1 = sheet.CreateRow(row);
                                //    for (int i = 0; i < flist1.Count(); i++)
                                //    {
                                //        crow.CreateCell(i).SetCellValue(flist1[i]);
                                //    }
                                //    row += 1;

                                //    workbook.Write(fs);
                                //}

                                //dalist.ItemsSource = null;
                                //huolist.ItemsSource = null;
                                //dalist.ItemsSource = flist;
                                //huolist.ItemsSource = flist1;
                                //dalist.Items.Refresh();
                                //huolist.Items.Refresh();
                                //bool mark = true;
                                //flist.ForEach(f =>
                                //{
                                //    if (mark)
                                //    {
                                //        mark = (config.DaJingMin <= f && f <= config.DaJingMax);
                                //    }
                                //});
                                //bool mark1 = true;
                                //flist1.ForEach(f =>
                                //{
                                //    if (mark1)
                                //    {
                                //        mark1 = (config.HuoSaiMin <= f && f <= config.HuoSaiMax);
                                //    }
                                //});

                                //// 回写PLC
                                //if (mark && mark1)
                                //{
                                //    plc.Write("M111", true);
                                //    gwType = GwType.槽径;
                                //    daResult.Text = "OK";
                                //    huoResult.Text = "OK";
                                //}
                                //else
                                //{
                                //    plc.Write("M121", true);
                                //    daResult.Text = "NG";
                                //    huoResult.Text = "NG";
                                //    dalist.Background = Brushes.Red;
                                //    huolist.Background = Brushes.Red;
                                //}
                                #endregion

                            }
                            else
                            {
                                ErrorInfo.Text = $"当前测量应在 {gwType.ToString()} 量测位置！";
                            }
                        }
                    }

                    var daReset = plc.ReadBool("M131");
                    if (daReset.IsSuccess && daReset.Content)
                    {
                        // clear
                        if (gwType == GwType.大径活塞高度)
                        {
                            dalist.ItemsSource = null;
                            dalist.Items.Refresh();
                            daResult.Text = "";
                            huolist.ItemsSource = null;
                            huolist.Items.Refresh();
                            huoResult.Text = "";
                            ErrorInfo.Text = "";
                            dalist.Background = Brushes.SteelBlue;
                            huolist.Background = Brushes.SteelBlue;
                            daNo = 0;
                            daMark = true;
                            huoMark = true;
                        }
                    }
                    #endregion

                    #region 槽径
                    var caojr = plc.ReadBool("M102");
                    if (caojr.IsSuccess && caojr.Content)
                    {
                        if (cjmodel) // true 为标准件模式 
                        {
                            Thread.Sleep(config.FirstTime);
                            ErrorInfo.Text = "标准件测量开始！";

                            var re = dataService.ReadData(GwType.槽径);
                            bool m = (config.CaoJingMin <= re && re <= config.CaoJingMax);

                            ShowModelInfo(re, GwType.槽径, m);
                        }
                        else
                        {
                            if (gwType == GwType.槽径)
                            {
                                caojNo += 1;
                                if (caojNo == 1)
                                {
                                    flist.Clear();
                                }

                                Thread.Sleep(config.FirstTime);
                                ErrorInfo.Text = "测量开始！";

                                if (caojMark)
                                {
                                    //读取数据
                                    var re = dataService.ReadData(gwType);
                                    flist.Add(re);

                                    caojMark = (config.CaoJingMin <= re && re <= config.CaoJingMax);
                                    if (!caojMark)
                                    {
                                        NgList.Add(re);
                                        caojlist.ItemsSource = null;
                                        caojlist.ItemsSource = flist;
                                        caojlist.Background = Brushes.Red;
                                        caojlist.Items.Refresh();

                                        plc.Write("M122", true);
                                        caojResult.Text = "NG";

                                        workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                        var sheet = workbook.GetSheetAt(sheetSum - 1);
                                        using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                        {
                                            //sheet.CreateRow(row).CreateCell(1).SetCellValue("槽径数据");
                                            //row += 1;
                                            var crow = sheet.CreateRow(row);
                                            crow.CreateCell(1).SetCellValue("槽径数据");
                                            for (int i = 0; i < flist.Count(); i++)
                                            {
                                                crow.CreateCell(i + 2).SetCellValue(flist[i]);
                                            }
                                            row += 1;
                                            workbook.Write(fs);

                                        }

                                    }
                                    else
                                    {
                                        caojlist.ItemsSource = null;
                                        caojlist.ItemsSource = flist;
                                        caojlist.Items.Refresh();
                                        caojResult.Text = "OK";
                                        if (caojNo == config.CaoJingCount)
                                        {

                                            plc.Write("M112", true);
                                            gwType = GwType.BUSH;

                                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                            var sheet = workbook.GetSheetAt(sheetSum - 1);
                                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                            {
                                                //sheet.CreateRow(row).CreateCell(1).SetCellValue("槽径数据");
                                                //row += 1;
                                                var crow = sheet.CreateRow(row);
                                                crow.CreateCell(1).SetCellValue("槽径数据");
                                                for (int i = 0; i < flist.Count(); i++)
                                                {
                                                    crow.CreateCell(i + 2).SetCellValue(flist[i]);
                                                }
                                                row += 1;
                                                workbook.Write(fs);
                                            }
                                            NgList.Clear();
                                            caojNo = 0;
                                        }
                                    }
                                }
                                else
                                {
                                    ErrorInfo.Text = "前次NG!";
                                }

                                #region cancel
                                //flist.Clear();
                                //读取数据
                                //for (int i = 0; i < 4; i++)
                                //{
                                //    var re = dataService.ReadData(gwType);
                                //    flist.Add(re);
                                //    if (i != 3)
                                //        Thread.Sleep(config.CaoJingTime);
                                //}
                                // write to excel
                                //workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                //var sheet = workbook.GetSheetAt(sheetSum - 1);
                                //using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                //{
                                //    sheet.CreateRow(row).CreateCell(0).SetCellValue("槽径数据");
                                //    row += 1;
                                //    var crow = sheet.CreateRow(row);
                                //    for (int i = 0; i < flist.Count(); i++)
                                //    {
                                //        crow.CreateCell(i).SetCellValue(flist[i]);
                                //    }
                                //    row += 1;
                                //    workbook.Write(fs);
                                //}

                                //caojlist.ItemsSource = null;
                                //caojlist.ItemsSource = flist;
                                //caojlist.Items.Refresh();
                                //bool mark = true;
                                //flist.ForEach(f =>
                                //{
                                //    if (mark)
                                //    {
                                //        mark = (config.CaoJingMin <= f && f <= config.CaoJingMax);
                                //    }
                                //});

                                //// 回写PLC
                                //if (mark)
                                //{
                                //    plc.Write("M112", true);
                                //    gwType = GwType.BUSH;
                                //    caojResult.Text = "OK";
                                //}
                                //else
                                //{
                                //    plc.Write("M122", true);
                                //    caojResult.Text = "NG";
                                //    caojlist.Background = Brushes.Red;
                                //}
                                #endregion
                            }
                            else
                            {
                                ErrorInfo.Text = $"当前测量应在 {gwType.ToString()} 量测位置！";
                            }
                        }
                    }

                    var caojReset = plc.ReadBool("M132");
                    if (caojReset.IsSuccess && caojReset.Content)
                    {
                        // clear
                        if (gwType == GwType.槽径)
                        {
                            caojlist.ItemsSource = null;
                            caojlist.Items.Refresh();
                            caojResult.Text = "";
                            ErrorInfo.Text = "";
                            caojlist.Background = Brushes.SteelBlue;
                            caojNo = 0;
                            caojMark = true;
                        }
                    }
                    #endregion

                    #region BUSH
                    var bushr = plc.ReadBool("M103");
                    if (bushr.IsSuccess && bushr.Content)
                    {
                        if (bushmodel) // true 为标准件模式 
                        {
                            Thread.Sleep(config.FirstTime);
                            ErrorInfo.Text = "标准件测量开始！";

                            var re = dataService.ReadData(GwType.BUSH);
                            bool m = (config.BushMin <= re && re <= config.BushMax);

                            ShowModelInfo(re, GwType.BUSH, m);
                        }
                        else
                        {
                            if (gwType == GwType.BUSH)
                            {
                                bushNo += 1;
                                if (bushNo == 1)
                                {
                                    flist.Clear();
                                }

                                Thread.Sleep(config.FirstTime);
                                ErrorInfo.Text = "测量开始！";
                                if (bushMark)
                                {
                                    //读取数据
                                    var re = dataService.ReadData(gwType);
                                    flist.Add(re);

                                    bushMark = (config.BushMin <= re && re <= config.BushMax);
                                    if (!bushMark)
                                    {
                                        NgList.Add(re);
                                        bushlist.ItemsSource = null;
                                        bushlist.ItemsSource = flist;
                                        bushlist.Background = Brushes.Red;
                                        bushlist.Items.Refresh();

                                        plc.Write("M123", true);
                                        bushResult.Text = "NG";

                                        workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                        var sheet = workbook.GetSheetAt(sheetSum - 1);
                                        using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                        {
                                            //sheet.CreateRow(row).CreateCell(0).SetCellValue("BUSH数据");
                                            //row += 1;
                                            var crow = sheet.CreateRow(row);
                                            crow.CreateCell(1).SetCellValue("BUSH数据");
                                            for (int i = 0; i < flist.Count(); i++)
                                            {
                                                crow.CreateCell(i + 1).SetCellValue(flist[i]);
                                            }
                                            row += 1;
                                            workbook.Write(fs);

                                        }

                                    }
                                    else
                                    {
                                        bushlist.ItemsSource = null;
                                        bushlist.ItemsSource = flist;
                                        bushlist.Items.Refresh();
                                        bushResult.Text = "OK";
                                        if (bushNo == config.BushCount)
                                        {

                                            plc.Write("M113", true);
                                            gwType = GwType.槽高;

                                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                            var sheet = workbook.GetSheetAt(sheetSum - 1);
                                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                            {
                                                //sheet.CreateRow(row).CreateCell(0).SetCellValue("BUSH数据");
                                                //row += 1;
                                                var crow = sheet.CreateRow(row);
                                                crow.CreateCell(1).SetCellValue("BUSH数据");
                                                for (int i = 0; i < flist.Count(); i++)
                                                {
                                                    crow.CreateCell(i + 1).SetCellValue(flist[i]);
                                                }
                                                row += 1;
                                                workbook.Write(fs);
                                            }
                                            NgList.Clear();
                                            bushNo = 0;
                                        }
                                    }
                                }
                                else
                                {
                                    ErrorInfo.Text = "前次NG!";
                                }

                                #region cancel
                                //flist.Clear();
                                //读取数据
                                //for (int i = 0; i < 8; i++)
                                //{
                                //    var re = dataService.ReadData(gwType);
                                //    flist.Add(re);
                                //    if (i != 7)
                                //        Thread.Sleep(config.BushTime);
                                //}
                                // write to excel
                                //workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                //var sheet = workbook.GetSheetAt(sheetSum - 1);
                                //using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                //{
                                //    sheet.CreateRow(row).CreateCell(0).SetCellValue("BUSH数据");
                                //    row += 1;
                                //    var crow = sheet.CreateRow(row);
                                //    for (int i = 0; i < flist.Count(); i++)
                                //    {
                                //        crow.CreateCell(i).SetCellValue(flist[i]);
                                //    }
                                //    row += 1;
                                //    workbook.Write(fs);
                                //}

                                //bushlist.ItemsSource = null;
                                //bushlist.ItemsSource = flist;
                                //bushlist.Items.Refresh();
                                //bool mark = true;
                                //flist.ForEach(f =>
                                //{
                                //    if (mark)
                                //    {
                                //        mark = (config.BushMin <= f && f <= config.BushMax);
                                //    }
                                //});

                                //// 回写PLC
                                //if (mark)
                                //{
                                //    plc.Write("M113", true);
                                //    gwType = GwType.槽高;
                                //    bushResult.Text = "OK";
                                //}
                                //else
                                //{
                                //    plc.Write("M123", true);
                                //    bushResult.Text = "NG";
                                //    bushlist.Background = Brushes.Red;
                                //}
                                #endregion
                            }
                            else
                            {
                                ErrorInfo.Text = $"当前测量应在 {gwType.ToString()} 量测位置！";
                            }
                        }
                    }

                    var bushReset = plc.ReadBool("M133");
                    if (bushReset.IsSuccess && bushReset.Content)
                    {
                        // clear
                        if (gwType == GwType.BUSH)
                        {
                            bushlist.ItemsSource = null;
                            bushlist.Items.Refresh();
                            bushResult.Text = "";
                            ErrorInfo.Text = "";
                            bushlist.Background = Brushes.SteelBlue;
                            bushNo = 0;
                            bushMark = true;
                        }
                    }
                    #endregion

                    #region 槽高
                    var caogr = plc.ReadBool("M104");
                    if (caogr.IsSuccess && caogr.Content)
                    {
                        if (cgmodel) // true 为标准件模式 
                        {
                            Thread.Sleep(config.FirstTime);
                            ErrorInfo.Text = "标准件测量开始！";

                            var re = dataService.ReadData(GwType.槽高);
                            bool m = (config.CaoGaoMin <= re && re <= config.CaoGaoMax);

                            ShowModelInfo(re, GwType.槽高, m);
                        }
                        else
                        {
                            if (gwType == GwType.槽高)
                            {
                                caogNo += 1;
                                if (caogNo == 1)
                                {
                                    flist.Clear();
                                }

                                Thread.Sleep(config.FirstTime);
                                ErrorInfo.Text = "测量开始！";
                                if (caogMark)
                                {
                                    //读取数据
                                    var re = dataService.ReadData(gwType);
                                    flist.Add(re);

                                    caogMark = (config.CaoGaoMin <= re && re <= config.CaoGaoMax);
                                    if (!caogMark)
                                    {
                                        NgList.Add(re);
                                        caogaolist.ItemsSource = null;
                                        caogaolist.ItemsSource = flist;
                                        caogaolist.Background = Brushes.Red;
                                        caogaolist.Items.Refresh();

                                        plc.Write("M124", true);
                                        caogaoResult.Text = "NG";

                                        workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                        var sheet = workbook.GetSheetAt(sheetSum - 1);
                                        using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                        {
                                            //sheet.CreateRow(row).CreateCell(0).SetCellValue("槽高数据");
                                            //row += 1;
                                            var crow = sheet.CreateRow(row);
                                            crow.CreateCell(1).SetCellValue("槽高数据");
                                            for (int i = 0; i < flist.Count(); i++)
                                            {
                                                crow.CreateCell(i + 1).SetCellValue(flist[i]);
                                            }
                                            row += 1;
                                            workbook.Write(fs);

                                        }

                                    }
                                    else
                                    {
                                        caogaolist.ItemsSource = null;
                                        caogaolist.ItemsSource = flist;
                                        caogaolist.Items.Refresh();
                                        caogaoResult.Text = "OK";
                                        if (caogNo == config.CaoGaoCount)
                                        {

                                            plc.Write("M114", true);
                                            gwType = GwType.小径;

                                            workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                            var sheet = workbook.GetSheetAt(sheetSum - 1);
                                            using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                            {
                                                //sheet.CreateRow(row).CreateCell(0).SetCellValue("槽高数据");
                                                //row += 1;
                                                var crow = sheet.CreateRow(row);
                                                crow.CreateCell(1).SetCellValue("槽高数据");
                                                for (int i = 0; i < flist.Count(); i++)
                                                {
                                                    crow.CreateCell(i + 1).SetCellValue(flist[i]);
                                                }
                                                row += 1;
                                                workbook.Write(fs);
                                            }
                                            NgList.Clear();
                                            caogNo = 0;
                                        }
                                    }
                                }
                                else
                                {
                                    ErrorInfo.Text = "前次NG!";
                                }

                                #region cancel
                                //flist.Clear();
                                //读取数据
                                //for (int i = 0; i < 4; i++)
                                //{
                                //    var re = dataService.ReadData(gwType);
                                //    flist.Add(re);
                                //    if (i != 3)
                                //        Thread.Sleep(config.CaoGaoTime);
                                //}
                                // write to excel
                                //workbook = new HSSFWorkbook(File.OpenRead(Path + "\\" + fileName));
                                //var sheet = workbook.GetSheetAt(sheetSum - 1);
                                //using (var fs = new FileStream(Path + "\\" + fileName, FileMode.OpenOrCreate))
                                //{
                                //    sheet.CreateRow(row).CreateCell(0).SetCellValue("槽高数据");
                                //    row += 1;
                                //    var crow = sheet.CreateRow(row);
                                //    for (int i = 0; i < flist.Count(); i++)
                                //    {
                                //        crow.CreateCell(i).SetCellValue(flist[i]);
                                //    }
                                //    row += 1;
                                //    workbook.Write(fs);
                                //}

                                //caogaolist.ItemsSource = null;
                                //caogaolist.ItemsSource = flist;
                                //caogaolist.Items.Refresh();
                                //bool mark = true;
                                //flist.ForEach(f =>
                                //{
                                //    if (mark)
                                //    {
                                //        mark = (config.CaoGaoMin <= f && f <= config.CaoGaoMax);
                                //    }
                                //});

                                //// 回写PLC
                                //if (mark)
                                //{
                                //    plc.Write("M114", true);
                                //    gwType = GwType.小径;
                                //    caogaoResult.Text = "OK";
                                //}
                                //else
                                //{
                                //    plc.Write("M124", true);
                                //    caogaoResult.Text = "NG";
                                //    caogaolist.Background = Brushes.Red;
                                //}
                                #endregion
                            }
                            else
                            {
                                ErrorInfo.Text = $"当前测量应在 {gwType.ToString()} 量测位置！";
                            }
                        }
                    }

                    var caogReset = plc.ReadBool("M134");
                    if (caogReset.IsSuccess && caogReset.Content)
                    {
                        // clear
                        if (gwType == GwType.槽高)
                        {
                            caogaolist.ItemsSource = null;
                            caogaolist.Items.Refresh();
                            caogaoResult.Text = "";
                            ErrorInfo.Text = "";
                            caogaolist.Background = Brushes.SteelBlue;
                            caogNo = 0;
                            caogMark = true;
                        }
                    }
                    #endregion

                    // 移除光电信号
                    var NGSingal = plc.ReadBool("M140");
                    if (NGSingal.IsSuccess && NGSingal.Content)
                    {
                        // 提示框
                        if (NgList.Any())
                        {
                            ShowInfo(NgList, gwType);
                        }

                        #region 复位
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
                        xiaolist.Background = Brushes.SteelBlue;
                        dalist.Background = Brushes.SteelBlue;
                        huolist.Background = Brushes.SteelBlue;
                        caogaolist.Background = Brushes.SteelBlue;
                        caojlist.Background = Brushes.SteelBlue;
                        bushlist.Background = Brushes.SteelBlue;

                        NgList.Clear();
                        xiaoNo = 0;
                        xiaoMark = true;
                        daNo = 0;
                        daMark = true;
                        huoMark = true;
                        caojNo = 0;
                        caojMark = true;
                        bushNo = 0;
                        bushMark = true;
                        caogNo = 0;
                        caogMark = true;
                        #endregion
                    }

                    remark = true;
                }
                catch (Exception exc)
                {
                    log.Error("------PLC访问出错------");
                    log.Error(gwType.ToString() + "  " + row + "  " + exc.Message);
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //var da = dataService.Readtest();

            //flist.Add(-235.51f);
            //xiaolist.ItemsSource = flist;
            //ShowInfo(flist, gwType);

            ShowModelInfo(23.2f, GwType.小径, false);
        }

        private void ShowInfo(List<float> list, GwType type)
        {

            InfoWindow iw = new InfoWindow(list, type);
            iw.WindowStartupLocation = WindowStartupLocation.Manual;
            Rect rc = SystemParameters.WorkArea; //获取工作区大小
            //this.Topmost = true;
            iw.Left = 0; //设置位置
            iw.Top = rc.Height * 0.3;
            iw.Width = rc.Width;
            iw.Height = rc.Height * 0.7;
            iw.ShowDialog();
        }

        private void ShowModelInfo(float data, GwType type, bool mark)
        {

            ModelInfoWindow iw = new ModelInfoWindow(data, type, mark);
            iw.WindowStartupLocation = WindowStartupLocation.Manual;
            Rect rc = SystemParameters.WorkArea; //获取工作区大小
            //this.Topmost = true;
            iw.Left = 0; //设置位置
            iw.Top = rc.Height * 0.3;
            iw.Width = rc.Width;
            iw.Height = rc.Height * 0.7;
            iw.ShowDialog();
        }

        private void ShowModelInfo1(float data1, float data2, GwType type, bool mark1, bool mark2)
        {

            ModelInfoWindow iw = new ModelInfoWindow(data1, data2, type, mark1, mark2);
            iw.WindowStartupLocation = WindowStartupLocation.Manual;
            Rect rc = SystemParameters.WorkArea; //获取工作区大小
            //this.Topmost = true;
            iw.Left = 0; //设置位置
            iw.Top = rc.Height * 0.3;
            iw.Width = rc.Width;
            iw.Height = rc.Height * 0.7;
            iw.ShowDialog();
        }

        #region model change
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ShowInfo(NgList, gwType);
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            xiaomodel = true;
            button1.Visibility = Visibility.Hidden;
            cancel1.Visibility = Visibility.Visible;
        }

        private void Button11_Click(object sender, RoutedEventArgs e)
        {
            xiaomodel = false;
            button1.Visibility = Visibility.Visible;
            cancel1.Visibility = Visibility.Hidden;
        }

        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            damodel = true;
            button2.Visibility = Visibility.Hidden;
            cancel2.Visibility = Visibility.Visible;
        }

        private void Button12_Click(object sender, RoutedEventArgs e)
        {
            damodel = false;
            button2.Visibility = Visibility.Visible;
            cancel2.Visibility = Visibility.Hidden;
        }

        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            huomodel = true;
            button3.Visibility = Visibility.Hidden;
            cancel3.Visibility = Visibility.Visible;
        }

        private void Button13_Click(object sender, RoutedEventArgs e)
        {
            huomodel = false;
            button3.Visibility = Visibility.Visible;
            cancel3.Visibility = Visibility.Hidden;
        }

        private void Button4_Click(object sender, RoutedEventArgs e)
        {
            cjmodel = true;
            button4.Visibility = Visibility.Hidden;
            cancel4.Visibility = Visibility.Visible;
        }

        private void Button14_Click(object sender, RoutedEventArgs e)
        {
            cjmodel = false;
            button4.Visibility = Visibility.Visible;
            cancel4.Visibility = Visibility.Hidden;
        }

        private void Button5_Click(object sender, RoutedEventArgs e)
        {
            bushmodel = true;
            button5.Visibility = Visibility.Hidden;
            cancel5.Visibility = Visibility.Visible;
        }

        private void Button15_Click(object sender, RoutedEventArgs e)
        {
            bushmodel = false;
            button5.Visibility = Visibility.Visible;
            cancel5.Visibility = Visibility.Hidden;
        }

        private void Button6_Click(object sender, RoutedEventArgs e)
        {
            cgmodel = true;
            button6.Visibility = Visibility.Hidden;
            cancel6.Visibility = Visibility.Visible;
        }

        private void Button16_Click(object sender, RoutedEventArgs e)
        {
            cgmodel = false;
            button6.Visibility = Visibility.Visible;
            cancel6.Visibility = Visibility.Hidden;
        }
        #endregion

    }
}
