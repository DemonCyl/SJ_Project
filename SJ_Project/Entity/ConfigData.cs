using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SJ_Project.Entity
{
    public class ConfigData
    {

        #region 接收器配置
        public string PortName { get; set; }

        public int BaudRate { get; set; }
        #endregion

        #region 三菱Plc配置

        public string PlcIpAddress { get; set; }

        public int PlcPort { get; set; }
        #endregion

        public float XiaoJingMax { get; set; }
        public float XiaoJingMin { get; set; }
        public int XiaoJingTime { get; set; }

        public float DaJingMax { get; set; }
        public float DaJingMin { get; set; }
        public int DaJingHuoSaiTime { get; set; }

        public float HuoSaiMax { get; set; }
        public float HuoSaiMin { get; set; }

        public float CaoJingMax { get; set; }
        public float CaoJingMin { get; set; }
        public int CaoJingTime { get; set; }

        public float BushMax { get; set; }
        public float BushMin { get; set; }
        public int BushTime { get; set; }

        public float CaoGaoMax { get; set; }
        public float CaoGaoMin { get; set; }
        public int CaoGaoTime { get; set; }

        public int FirstTime { get; set; }

        public int XiaoJingCount { get; set; }
        public int DaJingHuoSaiCount { get; set; }
        public int CaoJingCount { get; set; }
        public int BushCount { get; set; }
        public int CaoGaoCount { get; set; }
    }


    public enum GwType
    {
        小径 = 1,
        大径活塞高度 = 2,
        活塞高度 = 3,
        槽径 = 4,
        BUSH = 5,
        槽高 = 6
    }
}
