using SJ_Project.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace SJ_Project
{
    /// <summary>
    /// InfoWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ModelInfoWindow : Window
    {
        public ModelInfoWindow(float data, GwType type, bool mark)
        {
            InitializeComponent();

            NgName.Text = type.ToString();

            dataBox.Text = data.ToString();
            dataBox.Background = mark ? Brushes.SteelBlue : Brushes.Red;

            StartCloseTimer();
        }

        public ModelInfoWindow(float data1,float data2, GwType type, bool mark1,bool mark2)
        {
            InitializeComponent();

            NgName.Text = type.ToString();

            dataBox.Text = data1.ToString();
            dataBox.Background = mark1 ? Brushes.SteelBlue : Brushes.Red;

            dataBox_Copy.Text = data2.ToString();
            dataBox_Copy.Background = mark2 ? Brushes.SteelBlue : Brushes.Red;

            StartCloseTimer();
        }

        private void StartCloseTimer()
        {
            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(5); // 3秒
            timer.Tick += TimerTick; // 注册计时器到点后触发的回调
            timer.Start();
        }

        private void TimerTick(object sender, EventArgs e)
        {
            DispatcherTimer timer = (DispatcherTimer)sender;
            timer.Stop();
            timer.Tick -= TimerTick; // 取消注册
            this.Close();
        }

    }
}
