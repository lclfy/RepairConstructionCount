using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RepairConstructionCount
{
    public class MainFileTitles : ICloneable
    {
        //文件名
        public string _fileName { get; set; }
        //是施工还是天窗, 施工false，维修true；
        public bool _repairOrConstruction { get; set; }
        //按调度台统计所在列
        public int _statisticsByController_column { get; set; }
        //按施工单位统计
        public int _statisticsByDeparts_column { get; set; }
        //各调度台列
        public List<TrainControllerTitle> _trainControllerTitle { get; set; }
        //各施工单位列
        public List<DepartTitle> _departTitle { get; set; }
        public int _plannedCount_row { get; set; }
        public int _plannedTime_row { get; set; }
        public int _permitCount_row { get; set; }
        public int _permitTime_row { get; set; }
        public List<SpecialCauses_Title> _specialCauses_title { get; set; }
        //兑现率
        public double _demandRate_row { get; set; }


        public MainFileTitles()
        {
            _fileName = "";
            _repairOrConstruction = false;
            _statisticsByController_column = 0;
            _statisticsByDeparts_column = 0;
            _trainControllerTitle = new List<TrainControllerTitle>();
            _departTitle = new List<DepartTitle>();
            _plannedCount_row = 0;
            _plannedTime_row = 0;
            _permitCount_row = 0;
            _permitTime_row = 0;
            _specialCauses_title = new List<SpecialCauses_Title>();
            _demandRate_row = 0;
        }

        public object Clone()
        {
            MainFileTitles _r = new MainFileTitles();
            _r._fileName = this._fileName;
            _r._repairOrConstruction = this._repairOrConstruction;
            _r._statisticsByController_column = this._statisticsByController_column;
            _r._statisticsByDeparts_column = this._statisticsByDeparts_column;
            _r._trainControllerTitle = this._trainControllerTitle;
            _r._departTitle = this._departTitle;
            _r._plannedCount_row = this._plannedCount_row;
            _r._plannedTime_row = this._plannedTime_row;
            _r._permitCount_row = this._permitCount_row;
            _r._permitTime_row = this._permitTime_row;
            _r._specialCauses_title = this._specialCauses_title;
            _r._demandRate_row = this._demandRate_row;

            return _r as object;//深复制
        }

    }

    //按调度台统计-标题
    public class TrainControllerTitle
    {
        public string _controllerName { get; set; }
        public int _controllerColumn { get; set; }

        public TrainControllerTitle()
        {
            _controllerColumn = 0;
            _controllerName = "";
        }
    }

    //按施工单位统计-标题
    public class DepartTitle
    {
        public string _departName { get; set; }
        public int _departColumn { get; set; }

        public DepartTitle()
        {
            _departColumn = 0;
            _departName = "";
        }
    }

    //特殊原因所在的行
    public class SpecialCauses_Title
    {
        public string _specialCauseName { get; set; }
        public int _specialCauseCount_rowOrColumn { get; set; }
        public int _specialCauseTime_rowOrColumn { get; set; }

        public SpecialCauses_Title()
        {
            _specialCauseName = "";
            _specialCauseCount_rowOrColumn = 0;
            _specialCauseTime_rowOrColumn = 0;

        }
    }
}
