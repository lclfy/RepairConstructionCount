using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RepairConstructionCount
{
    public class SubFileTitle
    {
        public string _fileName { get; set; }
        //该车站名称
        public string _subStationName { get; set; }
        //0为施工，1为天窗
        public bool _repairOrConstruction { get; set; }
        //施工单位名称
        public string _departName { get; set; }
        public int _plannedCount_Column { get; set; }
        public int _plannedTime_Column { get; set; }
        public int _permitCount_Column { get; set; }
        public int _permitTime_Column { get; set; }
        //特殊情况列，在MainFileTitle文件内
        public List<SpecialCauses_Title> _specialCauses_title { get; set; }

        public SubFileTitle()
        {
            _fileName = "";
            _subStationName = "";
            _repairOrConstruction = false;
            _departName = "";
            _plannedCount_Column = 0;
            _plannedTime_Column = 0;
            _permitCount_Column = 0;
            _specialCauses_title = new List<SpecialCauses_Title>();
        }

    }
}
