using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RepairConstructionCount
{
    public class RailRepair : ICloneable
    {
        public int _id { get; set; }
        public string _stationName { get; set; }
        public string _departName { get; set; }
        public int _repairDate { get; set; }
        public int _plannedRepairCount { get; set; }
        public int _plannedRepairTime { get; set; }
        public int _askRepairCount { get; set; }
        public int _askRepairTime { get; set; }
        public int _permitRepairCount { get; set; }
        public int _permitRepairTime { get; set; }
        //SpecialCauses类在施工类中
        public List<SpecialCauses_ConsRepair> _specialCauses { get; set; }

        public RailRepair()
        {
            _id = 0;
            _stationName = "";
            _departName = "";
            _repairDate = 0;
            _plannedRepairCount = 0;
            _plannedRepairTime = 0;
            _askRepairCount = 0;
            _askRepairTime = 0;
            _permitRepairCount = 0;
            _permitRepairTime = 0;
            _specialCauses = new List<SpecialCauses_ConsRepair>();
        }

        public object Clone()
        {
            RailRepair _r = new RailRepair();
            _r._id = this._id;
            _r._stationName = this._stationName;
            _r._departName = this._departName;
            _r._repairDate = this._repairDate;
            _r._plannedRepairCount = this._plannedRepairCount;
            _r._plannedRepairTime = this._plannedRepairTime;
            _r._askRepairCount = this._askRepairCount;
            _r._askRepairTime = this._askRepairTime;
            _r._permitRepairCount = this._permitRepairCount;
            _r._permitRepairTime = this._permitRepairTime;
            _r._specialCauses = this._specialCauses;

            return _r as object;//深复制
        }
    }
}
