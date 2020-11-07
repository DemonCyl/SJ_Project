using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using SJ_Project.Entity;

namespace SJ_Project.Services
{
    public class DataService
    {
        private SerialPort serialPort;
        private ConfigData config;
        private ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static string cmd0 = @"DR1000000";
        private static string cmd1 = @"DR1010000";
        private static string cmd2 = @"DR1020000";
        private static string cmd3 = @"DR1030000";
        private static string cmd4 = @"DR1040000";
        private static string cmd5 = @"DR1050000";
        private byte[] cdd0 = { 0x44, 0x52, 0x31, 0x30, 0x30, 0x30, 0x30, 0x30, 0x30, 0x0d };
        private byte[] cdd1 = { 0x44, 0x52, 0x31, 0x30, 0x31, 0x30, 0x30, 0x30, 0x30, 0x0d };
        private byte[] cdd2 = { 0x44, 0x52, 0x31, 0x30, 0x32, 0x30, 0x30, 0x30, 0x30, 0x0d };
        private byte[] cdd3 = { 0x44, 0x52, 0x31, 0x30, 0x33, 0x30, 0x30, 0x30, 0x30, 0x0d };
        private byte[] cdd4 = { 0x44, 0x52, 0x31, 0x30, 0x34, 0x30, 0x30, 0x30, 0x30, 0x0d };
        private byte[] cdd5 = { 0x44, 0x52, 0x31, 0x30, 0x35, 0x30, 0x30, 0x30, 0x30, 0x0d };

        public DataService(ConfigData config)
        {
            this.config = config;
        }

        public bool GetConnection()
        {
            bool mark = false;
            if (serialPort == null)
            {
                serialPort = new SerialPort(config.PortName, config.BaudRate, Parity.None, 8, StopBits.One);
                serialPort.DtrEnable = true;
                serialPort.RtsEnable = true;
                serialPort.NewLine = "\r";
                serialPort.ReadTimeout = 3000;
                mark = OpenPort();
            }
            return mark;
        }

        private bool OpenPort()
        {
            string message = null;
            try//这里写成异常处理的形式以免串口打不开程序崩溃
            {
                serialPort.Open();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            if (serialPort.IsOpen)
            {
                log.Info("无线接收器连接成功！");
                return true;
            }
            else
            {
                log.Error("无线接收器打开失败!原因为： " + message);
                return false;
            }
        }

        public void Close()
        {
            if (serialPort.IsOpen)
            {
                serialPort.Close();
            }
        }

        public float ReadData(GwType type)
        {
            string re = null;
            float data = 0;
            byte[] cmd = null;
            if (serialPort.IsOpen)
            {
                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                try
                {
                    switch (type)
                    {
                        case GwType.小径:
                            cmd = cdd0;
                            break;
                        case GwType.大径活塞高度:
                            cmd = cdd1;
                            break;
                        case GwType.活塞高度:
                            cmd = cdd2;
                            break;
                        case GwType.槽径:
                            cmd = cdd3;
                            break;
                        case GwType.BUSH:
                            cmd = cdd4;
                            break;
                        case GwType.槽高:
                            cmd = cdd5;
                            break;
                    }
                    serialPort.Write(cmd, 0, cmd.Length);
                    re = serialPort.ReadLine();
                    data = TransformData(re);
                }
                catch (Exception ex)
                {
                    log.Error(ex.Message);
                    throw new Exception(ex.Message);
                }
            }
            return data;
        }

        /// <summary>
        /// 转换
        /// </summary>
        /// <param name="strData"></param>
        /// <returns></returns>
        public float TransformData(string strData)
        {
            float data = 0;
            // 00.000E-03  e.g. 00 000 -03


            var str1 = strData.Substring(7, 12);
            data = float.Parse(str1);

            return data;
        }

        public float Readtest(string cmda)
        {
            float data = 0;
            try
            {

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                log.Info("DR1FF0000");
                //serialPort.WriteLine(cmda);
                //var cdd = strToHexByte("DR1FF0000");
                log.Info(cdd0);
                serialPort.Write(cdd0, 0, cdd0.Length);

                var re = serialPort.ReadLine();
                log.Info(re);
                data = TransformData(re);
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
            }

            return data;
        }
    }
}
