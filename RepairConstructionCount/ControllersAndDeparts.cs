using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RepairConstructionCount
{
    public class ControllersAndDeparts : ICloneable
    {
        //调度台与施工单位类
        public string _codName { get; set; }
        public int _codPlannedCount { get; set; }
        public int _codPlannedTime { get; set; }
        public int _codPermitCount { get; set; }
        public int _codPermitTime { get; set; }

        //事故影响 自然灾害
        public int _causeByAccidentCount { get; set; }
        public int _causeByAccidentTime { get; set; }
        public int _causeByNatureCount { get; set; }
        public int _causeByNatureTime { get; set; }

        //部令取消 局令取消 车站取消 天气原因 单位未要 
        public int _causeByDepartCommandCount { get; set; }
        public int _causeByDepartCommandTime { get; set; }
        public int _causeByMainStreamCommandCount { get; set; }
        public int _causeByMainStreamCommandTime { get; set; }

        public int _causeByStationCount { get; set; }
        public int _causeByStationTime { get; set; }

        public int _causeByWeatherCount { get; set; }
        public int _causeByWeatherTime { get; set; }

        public int _causeByNotAskCount { get; set; }
        public int _causeByNotAskTime { get; set; }

        //备注
        public string extra { get; set; }

        public int _codColumn { get; set; }

        public ControllersAndDeparts()
        {
            _codName = "";
            extra = "";
        }

        public object Clone()
        {
            ControllersAndDeparts _cod = new ControllersAndDeparts();
            _cod._codName = _codName;
            _cod._codPermitCount = _codPermitCount;
            _cod._codPermitTime = _codPermitTime;
            _cod._codPlannedCount = _codPlannedCount;
            _cod._codPlannedTime = _codPlannedTime;
            _cod._causeByAccidentCount = _causeByAccidentCount;
            _cod._causeByAccidentTime = _causeByAccidentTime;
            _cod._causeByNatureCount = _causeByNatureCount;
            _cod._causeByNatureTime = _causeByNatureTime;
            _cod._causeByDepartCommandCount = _causeByDepartCommandCount;
            _cod._causeByDepartCommandTime = _causeByDepartCommandTime;
            _cod._causeByMainStreamCommandCount = _causeByMainStreamCommandCount;
            _cod._causeByMainStreamCommandTime = _causeByMainStreamCommandTime;
            _cod._causeByStationCount = _causeByStationCount;
            _cod._causeByStationTime = _causeByStationTime;
            _cod._causeByNotAskCount = _causeByNotAskCount;
            _cod._causeByNotAskTime = _causeByNotAskTime;
            _cod.extra = extra;
            _cod._codColumn = _codColumn;
            return _cod as object;//深复制
        }
    }
}
