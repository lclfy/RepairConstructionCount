using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RepairConstructionCount
{
    public class RailRepair
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
        }
    }
}
