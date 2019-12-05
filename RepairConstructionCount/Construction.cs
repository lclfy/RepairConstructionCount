using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RepairConstructionCount
{
    public class Construction
    {
        /*
        Id
departName 单位名称
plannedConstructionCount 计划次数
plannedConstructionTime 计划时间
askConstructionCount申请次数
askConstructionTime申请时间
permitConstructionCount给点次数
permitConstructionTime给点时间
List<SpecialCauses> 特殊原因
*/
        public int _id { get; set; }
        public string _stationName { get; set; }
        public string _departName { get; set; }
        public int _constructionDate { get; set; }
        public int _plannedConstructionCount { get; set; }
        public int _plannedConstructionTime { get; set; }
        public int _askConstructionCount { get; set; }
        public int _askConstructionTime { get; set; }
        public int _permitConstructionCount { get; set; }
        public int _permitConstructionTime { get; set; }
        public List<SpecialCauses_Cons> _specialCauses {get;set;}

        public Construction()
        {
            _id = 0;
            _stationName = "";
            _departName = "";
            _constructionDate = 0;
            _plannedConstructionCount = 0;
            _plannedConstructionTime = 0;
            _askConstructionCount = 0;
            _askConstructionTime = 0;
            _permitConstructionCount = 0;
            _permitConstructionTime = 0;
            _specialCauses = new List<SpecialCauses_Cons>();
        }
    }

    //特殊原因
    public class SpecialCauses_Cons
    {
        //是否参加考核
        public bool _examine { get; set; }
        //原因名称
        public string _causesName { get; set; }
        //原因导致次数
        public int _causesCount { get; set; }
        //原因导致时间
        public int _causesTime { get; set; }

        public SpecialCauses_Cons()
        {
            _examine = false;
            _causesName = "";
            _causesCount = 0;
            _causesTime = 0;
        }

    }

}

