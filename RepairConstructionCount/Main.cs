using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using CCWin;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;

namespace RepairConstructionCount
{
    public partial class Main : Skin_Mac
    {
        //总表地址
        string mainFile = "";
        //分表地址
        List<string> subFileList = new List<string>();
        //主文件标题,分施工和维修
        List<MainFileTitles> mainFileTitle;
        //各分支文件与调度台关系
        List<string> stationInController;
        //副文件标题
        List<SubFileTitle> allSubFileTitles;
        //维修
        List<RailRepair> railRepairs;
        //调度台(维修)
        List<ControllersAndDeparts> repairControllers;
        //单位(天窗)
        List<ControllersAndDeparts> repairDepart;
        bool hasFilePath = false;
        string procceingProgress = "";
        //文件中的车站列表，用于纠错
        List<string> stationsList;
        string consText = "";
        string repairText = "";
        string s_text = "";
        int stationCount = 0;
        //显示进度
        private delegate void SetPos(int ipos, string vinfo);
        Thread fThread;
        Stopwatch sw = new Stopwatch();

        public Main()
        {
            mainFileTitle = new List<MainFileTitles>();
            allSubFileTitles = new List<SubFileTitle>();
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            refresh();
            stationCount = 0;
            fThread = new Thread(new ThreadStart(SleepT));
            stationInController = new List<string>();
            repairControllers = new List<ControllersAndDeparts>();
            repairDepart = new List<ControllersAndDeparts>();
            stationsList = new List<string>();
            start_btn.Enabled = false;
            processing_lbl.Text = procceingProgress;
        }

        private void refresh()
        {
            stationCount = 0;
            railRepairs = new List<RailRepair>();
            stationInController = new List<string>();
            repairControllers = new List<ControllersAndDeparts>();
            repairDepart = new List<ControllersAndDeparts>();
            stationsList = new List<string>();
            stationController_rtb.Clear();
            repairDeaprtList_rtb.Clear();
            stationsList_rtb.Clear();
        }


        private void selectPath(int mainOrSub)
        {
            refresh();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();   //显示选择文件对话框 
            openFileDialog1.Filter = "Excel 2003 文件 (*.xls)|*.xls";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            //main为0，sub为1
            if (mainOrSub == 0)
            {
                openFileDialog1.Multiselect = false;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    mainFileTitle = new List<MainFileTitles>();
                    mainFile = "";
                    foreach (string fileName in openFileDialog1.FileNames)
                    {
                        mainFile = fileName;
                        //施工和维修分别创建
                        MainFileTitles repairTitle = new MainFileTitles();
                        repairTitle._fileName = fileName;
                        repairTitle._repairOrConstruction = true;
                        mainFileTitle.Add(repairTitle);
                    }
                    this.mainExcelFile_lbl.Text = "已选择：" + mainFile.Split('\\')[mainFile.Split('\\').Count() - 1];
                }
            }
            else
            {
                openFileDialog1.Multiselect = true;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    allSubFileTitles = new List<SubFileTitle>();
                    subFileList = new List<string>();
                    int fileCount = 0;
                    foreach (string fileName in openFileDialog1.FileNames)
                    {
                        fileCount++;
                        subFileList.Add(fileName);
                    }
                    this.subExcelFile_lbl.Text = "已选择：" + fileCount + "个文件";
                }
            }
            if(mainFile.Length != 0 && mainFile != "" && subFileList != null && subFileList.Count != 0)
            {
                hasFilePath = true;
            }
            startBtnCheck();
        }

        private void startBtnCheck()
        {
            if (hasFilePath)
            {
                start_btn.Enabled = true;
            }
        }

        //读主图
        private void readMainFile()
        {
            procceingProgress = "正在读取文件…";
            try
            {
                if (mainFile == null)
                {
                    MessageBox.Show("请重新选择主文件~", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                FileStream fileStream = new FileStream(mainFile, FileMode.Open, FileAccess.Read, FileShare.Read);
                IWorkbook workbook = null;
                if (mainFile.IndexOf(".xls") > 0) // 2003版本  
                {
                    try
                    {
                        workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("出现错误，请重新选择文件，错误内容：" + e.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                ISheet repairSheet = workbook.GetSheet("高铁天窗");
                ISheet stationControllerSheet = workbook.GetSheet("统计设置");
                MainFileTitles repairTitle = (MainFileTitles)mainFileTitle[0].Clone();
                mainFileTitle.Clear();
                //再找维修的
                repairTitle = findMainTitles(repairSheet, repairTitle);
                mainFileTitle.Add(repairTitle);
                //再找调度台车站对应关系
                stationInController = FindStationController(stationControllerSheet, new List<string>());
                fileStream.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("请关闭所有打开的已选文件，错误内容：" + e.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private MainFileTitles findMainTitles(ISheet _sheet ,MainFileTitles _mainTitle)
        {
            int _statisticController = 0;
            int _statisticsDepart = 0;
            bool hasGotIt = false;
            for (int rowNum = 0; rowNum <= _sheet.LastRowNum; rowNum++)
            {
                IRow row = _sheet.GetRow(rowNum);
                if (row == null)
                {
                    continue;
                }
                for (int columnNum = 0; columnNum <= row.LastCellNum; columnNum++)
                {
                    if (row.GetCell(columnNum) != null && row.GetCell(columnNum).ToString().Length != 0)
                    {
                        ICell cell = row.GetCell(columnNum);
                        if (cell.ToString().Contains("按调度台统计"))
                        {
                            _mainTitle._statisticsByController_column = columnNum;
                            _statisticController = columnNum;
                            //开始找台
                            ICell _controllerCell = null;
                            int _controllerColumnNum = columnNum;
                            //从下一行开始找
                            IRow _controllerRow = _sheet.GetRow(rowNum + 1);
                            _controllerCell = _controllerRow.GetCell(_controllerColumnNum);
                            while (!_controllerCell.ToString().Contains("小计") && _controllerColumnNum <= _controllerRow.LastCellNum)
                            {
                                //找到了一个台
                                if(_controllerRow.GetCell(_controllerColumnNum) != null && _controllerRow.GetCell(_controllerColumnNum).ToString().Length != 0)
                                {
                                    TrainControllerTitle _tct = new TrainControllerTitle();
                                    _tct._controllerName = _controllerRow.GetCell(_controllerColumnNum).ToString();
                                    _tct._controllerColumn = _controllerColumnNum;
                                    _mainTitle._trainControllerTitle.Add(_tct);
                                }
                                _controllerColumnNum++;
                                _controllerCell = _controllerRow.GetCell(_controllerColumnNum);
                                if (_controllerCell.ToString().Contains("小计"))
                                {
                                    TrainControllerTitle _tct = new TrainControllerTitle();
                                    _tct._controllerName = "小计";
                                    _tct._controllerColumn = _controllerColumnNum;
                                    _mainTitle._trainControllerTitle.Add(_tct);
                                    break;
                                }
                            }
                        }
                        if (cell.ToString().Contains("按作业单位统计"))
                        {
                            _mainTitle._statisticsByDeparts_column = columnNum;
                            _statisticsDepart = columnNum;
                            //开始找单位
                            ICell _departCell = null;
                            int _departColumnNum = columnNum;
                            //从下一行开始找
                            IRow _departRow = _sheet.GetRow(rowNum + 1);
                            _departCell = _departRow.GetCell(_departColumnNum);
                            while (!_departCell.ToString().Contains("小计") && _departColumnNum <= _departRow.LastCellNum)
                            {
                                //找到了一个单位
                                if (_departRow.GetCell(_departColumnNum) != null && _departRow.GetCell(_departColumnNum).ToString().Length != 0)
                                {
                                    DepartTitle _dpt = new DepartTitle();
                                    _dpt._departName = _departRow.GetCell(_departColumnNum).ToString();
                                    _dpt._departColumn = _departColumnNum;
                                    _mainTitle._departTitle.Add(_dpt);
                                }
                                _departColumnNum++;
                                _departCell = _departRow.GetCell(_departColumnNum);
                                if (_departCell.ToString().Contains("小计"))
                                {
                                    DepartTitle _dpt = new DepartTitle();
                                    _dpt._departName = "小计";
                                    _dpt._departColumn = _departColumnNum;
                                    _mainTitle._departTitle.Add(_dpt);
                                    break;
                                }
                            }
                        }
                        if (cell.ToString().Contains("基数") && !hasGotIt)
                        {
                            _mainTitle._plannedCount_row = rowNum;
                            _mainTitle._plannedTime_row = rowNum + 1;
                        }
                        if (cell.ToString().Contains("给点") && !hasGotIt)
                        {
                            _mainTitle._permitCount_row = rowNum;
                            _mainTitle._permitTime_row = rowNum + 1;
                        }
                        if(_mainTitle._plannedCount_row != 0 && _mainTitle._permitCount_row != 0)
                        {
                            hasGotIt = true;
                        }
                        if (cell.ToString().Contains("不\n参\n加\n考\n核"))
                        {
                            int notExaminedRowNum = rowNum;
                            bool firstSearch = true;
                            int notExaminedColumnNum = columnNum + 1;
                            //先往右，再逐渐往下
                            while(!_sheet.GetRow(notExaminedRowNum).GetCell(notExaminedColumnNum - 1).ToString().Contains("参\n加\n考\n核") || firstSearch)
                            {
                                firstSearch = false;
                                if(_sheet.GetRow(notExaminedRowNum).GetCell(notExaminedColumnNum) != null &&
                                    _sheet.GetRow(notExaminedRowNum).GetCell(notExaminedColumnNum).ToString().Length != 0)
                                {
                                    SpecialCauses_Title _spct = new SpecialCauses_Title();
                                    _spct._specialCauseName = _sheet.GetRow(notExaminedRowNum).GetCell(notExaminedColumnNum).ToString();
                                    _spct._specialCauseCount_rowOrColumn = notExaminedRowNum;
                                    _spct._specialCauseTime_rowOrColumn = notExaminedRowNum + 1;
                                    _mainTitle._specialCauses_title.Add(_spct);
                                }
                                notExaminedRowNum++;
                            }
                        }
                        if (cell.ToString().Equals("参\n加\n考\n核"))
                        {
                            int examinedRowNum = rowNum;
                            int examinedColumnNum = columnNum + 1;
                            //先往右，再逐渐往下
                            while (!_sheet.GetRow(examinedRowNum).GetCell(examinedColumnNum - 1).ToString().Contains("兑现率"))
                            {
                                if (_sheet.GetRow(examinedRowNum).GetCell(examinedColumnNum) != null &&
                                    _sheet.GetRow(examinedRowNum).GetCell(examinedColumnNum).ToString().Length != 0)
                                {
                                    SpecialCauses_Title _spct = new SpecialCauses_Title();
                                    _spct._specialCauseName = _sheet.GetRow(examinedRowNum).GetCell(examinedColumnNum).ToString();
                                    _spct._specialCauseCount_rowOrColumn = examinedRowNum;
                                    _spct._specialCauseTime_rowOrColumn = examinedRowNum + 1;
                                    _mainTitle._specialCauses_title.Add(_spct);
                                }
                                examinedRowNum++;
                            }
                        }
                        if (cell.ToString().Contains("兑现率"))
                        {
                            _mainTitle._demandRate_row = rowNum;
                        }
                    }
                }

            }
            return _mainTitle;
        }

        private List<string> FindStationController(ISheet _sheet ,List<string> _StationController)
        {
            int stationColumn = 0;
            int controllerColumn = 0;
            int titleRow = 0;
            bool hasGotIt = false;
            //找标题
            for(int rowNum=0; rowNum <= _sheet.LastRowNum; rowNum++)
            {
                IRow row = _sheet.GetRow(rowNum);
                if(row == null)
                {
                    continue;
                }
                for(int columnNum = 0; columnNum<=row.LastCellNum; columnNum++)
                {
                    if(row.GetCell(columnNum)!= null && row.GetCell(columnNum).ToString().Length != 0)
                    {
                        ICell cell = row.GetCell(columnNum);
                        if (cell.ToString().Contains("站场名称"))
                        {
                            titleRow = rowNum;
                            stationColumn = columnNum;
                        }
                        if (cell.ToString().Contains("所属调度台"))
                        {
                            controllerColumn = columnNum;
                        }
                        if(stationColumn != 0 && controllerColumn != 0)
                        {
                            hasGotIt = true;
                            break;
                        }
                    }
                }
                if (hasGotIt)
                {
                    break;
                }
            }
            //找数据
            for(int rowNum = titleRow + 1; rowNum <= _sheet.LastRowNum; rowNum++)
            {
                IRow row = _sheet.GetRow(rowNum);
                if(row == null)
                {
                    continue;
                }
                ICell stationCell = row.GetCell(stationColumn);
                ICell controllerCell = row.GetCell(controllerColumn);
                if(stationCell == null || controllerCell == null)
                {
                    continue;
                }
                string station_Controller = stationCell.ToString().Trim() + "_" + controllerCell.ToString().Trim();
                _StationController.Add(station_Controller);
                //添加调度台
                ControllersAndDeparts _tempController = new ControllersAndDeparts();
                _tempController._codName = controllerCell.ToString().Trim();
                //找调度台所在的列
                foreach(TrainControllerTitle _tct in mainFileTitle[0]._trainControllerTitle)
                {
                    if (_tct._controllerName.Equals(_tempController._codName))
                    {//名字匹配到了
                        _tempController._codColumn = _tct._controllerColumn;
                    }
                }
                bool findSame = false;
                foreach(ControllersAndDeparts _cod in repairControllers)
                {
                    if (_cod._codName.Equals(_tempController._codName))
                    {
                        findSame = true;
                        break;
                    }
                }
                if (!findSame)
                {
                    repairControllers.Add((ControllersAndDeparts)_tempController.Clone());
                }
            }
            return _StationController;
        }

        //读副图
        private void readSubFiles()
        {
            //try
            {
                //procceingProgress = "正在读取文件…";
                foreach (string _subFile in subFileList)
                {
                    procceingProgress = "正在读取：" + _subFile.Split('\\')[_subFile.Split('\\').Count() - 1];
                    if (_subFile == null)
                    {
                        MessageBox.Show("请重新选择所有副文件~", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    FileStream fileStream = new FileStream(_subFile, FileMode.Open, FileAccess.Read, FileShare.Read);
                    IWorkbook workbook = null;
                    if (_subFile.IndexOf(".xls") > 0) // 2003版本  
                    {
                        try
                        {
                            workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("出现错误，请重新选择文件，错误内容：" + e.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    List<ISheet> repairSheet = new List<ISheet>();
                    for(int sheetCount = 0;sheetCount < workbook.NumberOfSheets; sheetCount++)
                    {
                        repairSheet.Add(workbook.GetSheetAt(sheetCount));
                    }
                    string currentStation = "";
                    //维修表头重写
                    List<SubFileTitle> _subFileTitles = new List<SubFileTitle>();
                    foreach(ISheet _rs in repairSheet)
                    {
                        currentStation = "";
                        int titleRowNum = -1;
                        int contentRowNum = -1;
                        int endRowNum = -1;
                        List<RailRepair> _rrS = new List<RailRepair>();
                        for(int ij= 0; ij <= _rs.LastRowNum; ij++)
                        { //找标题行和标题行，内容行，结束行,
                            IRow row = _rs.GetRow(ij);
                            if(row != null)
                            {
                                if(ij ==0 && (row.GetCell(0) == null || row.GetCell(0).ToString().Length == 0))
                                {
                                    MessageBox.Show("该表内某标签无站名，无法统计至调度台：" + _subFile + "\n第" + (repairSheet.IndexOf(_rs)+1) + "个标签","提示",MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                if(row.GetCell(0) != null)
                                {
                                    if (row.GetCell(0).ToString().Contains("综合天窗日统计表"))
                                    {//车站名称
                                        currentStation = row.GetCell(0).ToString().Trim().Replace(" ","").Replace("（","(").Replace("）",")").Replace("线路所)",")").Replace("站)",")").Split('(')[1].Split(')')[0];
                                        //车站名称加入列表
                                        bool hasGotIt = false;
                                        foreach(string _station in stationsList)
                                        {
                                            if (currentStation.Equals(_station.Trim()))
                                            {
                                                hasGotIt = true;
                                                continue;
                                            }
                                        }
                                        if(hasGotIt == false)
                                        {
                                            stationsList.Add(currentStation);
                                            stationCount++;
                                            s_text = s_text + currentStation + "\n";
                                        }
                                    }
                                    if (row.GetCell(0).ToString().Contains("日期星期"))
                                    {//标题
                                        titleRowNum = ij;
                                    }
                                    if (row.GetCell(0).ToString().Contains("周计"))
                                    {//结束
                                        endRowNum = ij;
                                    }
                                    if (row.GetCell(1).ToString().Contains("计划次数"))
                                    {
                                        contentRowNum = ij;
                                    }
                                }
                            }
                            else
                            {
                                continue;
                            }
                        }
                        if(titleRowNum != -1 && endRowNum != -1)
                        {
                            {
                                IRow row = _rs.GetRow(titleRowNum);
                                //开始找设备单位名称和他们的计划/给点时间
                                for (int columns = 1; columns <= row.LastCellNum; columns++)
                                {
                                    if (row.GetCell(columns) != null && row.GetCell(columns).ToString().Length != 0 && !row.GetCell(columns).ToString().Trim().Equals("") && !row.GetCell(columns).ToString().Trim().Contains("车站值班员"))
                                    {//找到了一个设备单位
                                        string tempDepartName = row.GetCell(columns).ToString().Trim();
                                        int planCountColumn = columns;
                                        int planTimeColumn = columns + 1;
                                        int askCountColumn = columns+2;
                                        int askTimeColumn = columns +3;
                                        int permitCountColumn = columns +4;
                                        int permitTimeColumn = columns +5;
                                        //往下找次数与时间
                                        for (int im = contentRowNum +1; im < endRowNum; im++)
                                        {
                                            RailRepair _tempRR = new RailRepair();
                                            _tempRR._stationName = currentStation;
                                            _tempRR._departName = tempDepartName;
                                            IRow _contentRow = _rs.GetRow(im);
                                            if (row.GetCell(planCountColumn) != null)
                                            {
                                                int _temp = -1;
                                                int.TryParse(_contentRow.GetCell(planCountColumn).ToString().Trim(), out _temp);
                                                _tempRR._plannedRepairCount = _temp;
                                            }
                                            if (row.GetCell(planTimeColumn) != null)
                                            {
                                                int _temp = -1;
                                                int.TryParse(_contentRow.GetCell(planTimeColumn).ToString().Trim(), out _temp);
                                                _tempRR._plannedRepairTime = _temp;
                                            }
                                            if (row.GetCell(askCountColumn) != null)
                                            {
                                                int _temp = -1;
                                                int.TryParse(_contentRow.GetCell(askCountColumn).ToString().Trim(), out _temp);
                                                _tempRR._askRepairCount = _temp;
                                            }
                                            if (row.GetCell(askTimeColumn) != null)
                                            {
                                                int _temp = -1;
                                                int.TryParse(_contentRow.GetCell(askTimeColumn).ToString().Trim(), out _temp);
                                                _tempRR._askRepairTime = _temp;
                                            }
                                            if (row.GetCell(permitCountColumn) != null)
                                            {
                                                int _temp = -1;
                                                int.TryParse(_contentRow.GetCell(permitCountColumn).ToString().Trim(), out _temp);
                                                _tempRR._permitRepairCount = _temp;
                                            }
                                            if (row.GetCell(permitTimeColumn) != null)
                                            {
                                                int _temp = -1;
                                                int.TryParse(_contentRow.GetCell(permitTimeColumn).ToString().Trim(), out _temp);
                                                _tempRR._permitRepairTime = _temp;
                                            }
                                            railRepairs.Add(_tempRR);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("选定的文件格式出错:第一列无_日期星期_或_周计\n" + _subFile);
                        }
                    }
                    fileStream.Close();
                    //添加
                    foreach(SubFileTitle _sft in _subFileTitles)
                    {
                        SubFileTitle _tempSFT = new SubFileTitle();
                        _tempSFT._fileName = _sft._fileName;
                        _tempSFT._subDeparts = _sft._subDeparts;
                        _tempSFT._subStationName = _sft._subStationName;
                        _tempSFT._specialCauses_title = _sft._specialCauses_title;
                        allSubFileTitles.Add(_tempSFT);
                    }
                }
                //匹配调度台
                matchControllersWithStations();
                //遍历所有的设备单位
                searchAllDeparts();
                //匹配设备单位
                matchDepartsWithStations();
                //填表
                FillTheForm();
                fThread.Abort();
            }
            /*
            catch (Exception e)
            {
                fThread.Abort();
                MessageBox.Show("运行出现问题，错误内容：" + e.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            */

        }

        //计算并填写
        //匹配调度台
        private void matchControllersWithStations()
        {
            procceingProgress = "正在处理…";
            for (int count = 0; count < repairControllers.Count; count++)
            {
                foreach (string _stationControllers in stationInController)
                {
                    if (_stationControllers.Contains(repairControllers[count]._codName))
                    {
                        foreach (RailRepair _r in railRepairs)
                        {
                            if (_r._stationName.Trim().Equals(_stationControllers.Split('_')[0].Trim()))
                            {//匹配到
                                repairControllers[count]._codPlannedCount += _r._plannedRepairCount;
                                repairControllers[count]._codPlannedTime += _r._plannedRepairTime;
                                repairControllers[count]._codPermitCount += _r._plannedRepairCount;
                                repairControllers[count]._codPermitTime += _r._plannedRepairTime;
                                //特殊
                                foreach (SpecialCauses_ConsRepair _spcc in _r._specialCauses)
                                {
                                    if (_spcc._causesName.Contains("事故抢险"))
                                    {
                                        repairControllers[count]._causeByAccidentCount += _spcc._causesCount;
                                        repairControllers[count]._causeByAccidentTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("自然灾害"))
                                    {
                                        repairControllers[count]._causeByNatureCount += _spcc._causesCount;
                                        repairControllers[count]._causeByNatureTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("部令取消"))
                                    {
                                        repairControllers[count]._causeByDepartCommandCount += _spcc._causesCount;
                                        repairControllers[count]._causeByDepartCommandTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("局令取消"))
                                    {
                                        repairControllers[count]._causeByMainStreamCommandCount += _spcc._causesCount;
                                        repairControllers[count]._causeByMainStreamCommandTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("单位未要"))
                                    {
                                        repairControllers[count]._causeByNotAskCount += _spcc._causesCount;
                                        repairControllers[count]._causeByNotAskTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("天气影响"))
                                    {
                                        repairControllers[count]._causeByWeatherCount += _spcc._causesCount;
                                        repairControllers[count]._causeByWeatherTime += _spcc._causesTime;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        //遍历设备单位并把他们找到合适的列
        private void searchAllDeparts()
        {
            //合适的列
            if(mainFileTitle.Count == 0)
            {
                return;
            }
            int columnCount = mainFileTitle[0]._statisticsByDeparts_column;

            foreach (RailRepair _r in railRepairs)
            {
                bool hasSame = false;
                foreach (ControllersAndDeparts _tempCod in repairDepart)
                {
                    if (_tempCod._codName.Equals(_r._departName))
                    {
                        hasSame = true;
                        break;
                    }
                }
                if (hasSame)
                {
                    continue;
                }
                else
                {
                    if(_r._departName.Length != 0)
                    {
                        ControllersAndDeparts _cod = new ControllersAndDeparts();
                        _cod._codName = _r._departName;
                        _cod._codColumn = columnCount;
                        repairDepart.Add(_cod);
                        columnCount++;
                    }

                }
            }
        }

        //匹配设备单位
        private void matchDepartsWithStations()
        {
            for (int count = 0; count < repairDepart.Count; count++)
            {
                        foreach (RailRepair _r in railRepairs)
                        {
                            if (_r._departName.Trim().Equals(repairDepart[count]._codName.Trim()))
                            {//匹配到
                        repairDepart[count]._codPlannedCount += _r._plannedRepairCount;
                        repairDepart[count]._codPlannedTime += _r._plannedRepairTime;
                        repairDepart[count]._codPermitCount += _r._plannedRepairCount;
                        repairDepart[count]._codPermitTime += _r._plannedRepairTime;
                        //维修作业
                                //特殊
                                foreach (SpecialCauses_ConsRepair _spcc in _r._specialCauses)
                                {
                                    if (_spcc._causesName.Contains("事故抢险"))
                                    {
                                repairDepart[count]._causeByAccidentCount += _spcc._causesCount;
                                repairDepart[count]._causeByAccidentTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("自然灾害"))
                                    {
                                repairDepart[count]._causeByNatureCount += _spcc._causesCount;
                                repairDepart[count]._causeByNatureTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("部令取消"))
                                    {
                                repairDepart[count]._causeByDepartCommandCount += _spcc._causesCount;
                                repairDepart[count]._causeByDepartCommandTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("局令取消"))
                                    {
                                repairDepart[count]._causeByMainStreamCommandCount += _spcc._causesCount;
                                repairDepart[count]._causeByMainStreamCommandTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("单位未要"))
                                    {
                                repairDepart[count]._causeByNotAskCount += _spcc._causesCount;
                                repairDepart[count]._causeByNotAskTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("天气影响"))
                                    {
                                repairDepart[count]._causeByWeatherCount += _spcc._causesCount;
                                repairDepart[count]._causeByWeatherTime += _spcc._causesTime;
                                    }
                                }
                            }
                        }
            }
        }

        //填表
        private void FillTheForm()
        {
            IWorkbook workbook = null;  //新建IWorkbook对象 
            FileStream fileStream = new FileStream(mainFile, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            if (mainFile.IndexOf(".xls") > 0) // 2003版本  
            {
                try
                {
                    workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
                }
                catch (Exception e)
                {
                    MessageBox.Show("主表文件出现损坏\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            ICellStyle normalStyle = workbook.CreateCellStyle();
            normalStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightBlue.Index;
            normalStyle.FillPattern = FillPattern.SolidForeground;
            normalStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LightBlue.Index;
            normalStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            normalStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            normalStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            normalStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            HSSFFont normalFont = (HSSFFont)workbook.CreateFont();
            normalFont.FontName = "宋体";//字体  
            normalFont.Color = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            normalFont.FontHeightInPoints = short.Parse("12");//字号  
            normalFont.IsBold = true;
            normalStyle.SetFont(normalFont);

        

            //维修
            int accidentCountRow_rep = 0;
            int accidentTimeRow_rep = 0;

            int natureCountRow_rep = 0;
            int natureTimeRow_rep = 0;

            int departComCountRow_rep = 0;
            int departComTimeRow_rep = 0;

            int mainStreamComCountRow_rep = 0;
            int mainStreamComTimeRow_rep = 0;

            int stationCountRow_rep = 0;
            int stationTimeRow_rep = 0;

            int weatherCountRow_rep = 0;
            int weatherTimeRow_rep = 0;

            int unitCountRow_rep = 0;
            int unitTimeRow_rep = 0;
            try
            {
                accidentCountRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("事故影响"))._specialCauseCount_rowOrColumn;
                accidentTimeRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("事故影响"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                natureCountRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("自然灾害"))._specialCauseCount_rowOrColumn;
                natureTimeRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("自然灾害"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                departComCountRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("部令取消"))._specialCauseCount_rowOrColumn;
                departComTimeRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("部令取消"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                mainStreamComCountRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("局令取消"))._specialCauseCount_rowOrColumn;
                mainStreamComTimeRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("局令取消"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                stationCountRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("车站取消"))._specialCauseCount_rowOrColumn;
                stationTimeRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("车站取消"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                weatherCountRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("天气影响"))._specialCauseCount_rowOrColumn;
                weatherTimeRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("天气影响"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                unitCountRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("单位未要"))._specialCauseCount_rowOrColumn;
                unitTimeRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("单位未要"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }

            ///维修
            MainFileTitles _mfRepair = mainFileTitle[0];
            ISheet sheetRepair = workbook.GetSheet("高铁天窗");
            IRow rowPlanCount_repair = sheetRepair.GetRow(_mfRepair._plannedCount_row);
            IRow rowPlanTime_repair = sheetRepair.GetRow(_mfRepair._plannedTime_row);

            IRow rowPermitCount_repair = sheetRepair.GetRow(_mfRepair._permitCount_row);
            IRow rowPermitTime_repair = sheetRepair.GetRow(_mfRepair._permitTime_row);

            IRow rowAccidentCount_repair = sheetRepair.GetRow(accidentCountRow_rep);
            IRow rowAccidentTime_repair = sheetRepair.GetRow(accidentTimeRow_rep);

            IRow rowNatureCount_repair = sheetRepair.GetRow(natureCountRow_rep);
            IRow rowNatureTime_repair = sheetRepair.GetRow(natureTimeRow_rep);

            IRow rowDepartCount_repair = sheetRepair.GetRow(departComCountRow_rep);
            IRow rowDepartTime_repair = sheetRepair.GetRow(departComTimeRow_rep);

            IRow rowMainStreamCount_repair = sheetRepair.GetRow(mainStreamComCountRow_rep);
            IRow rowMainStreamTime_repair = sheetRepair.GetRow(mainStreamComTimeRow_rep);

            IRow rowStationCount_repair = sheetRepair.GetRow(stationCountRow_rep);
            IRow rowStationTime_repair = sheetRepair.GetRow(stationTimeRow_rep);

            IRow rowWeatherCount_repair = sheetRepair.GetRow(weatherCountRow_rep);
            IRow rowWeatherTime_repair = sheetRepair.GetRow(weatherTimeRow_rep);

            IRow rowUnitCount_repair = sheetRepair.GetRow(unitCountRow_rep);
            IRow rowUnitTime_repair = sheetRepair.GetRow(unitTimeRow_rep);

            foreach (ControllersAndDeparts _cod in repairControllers)
            {
                //1-4行为普通情况
                if (rowPlanCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowPlanCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPlanCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPlannedCount);
                }

                if (rowPlanTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowPlanTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPlanTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPlannedTime);
                }

                if (rowPermitCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowPermitCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPermitCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPermitCount);
                }

                if (rowPermitTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowPermitTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPermitTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPermitTime);
                }

                if (rowAccidentCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowAccidentCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowAccidentCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByAccidentCount);
                }

                if (rowAccidentTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowAccidentTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowAccidentTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByAccidentTime);
                }

                if (rowNatureCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowNatureCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowNatureCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNatureCount);
                }

                if (rowNatureTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowNatureTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowNatureTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNatureTime);
                }

                if (rowDepartCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowDepartCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowDepartCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByDepartCommandCount);
                }

                if (rowDepartTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowDepartTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowDepartTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByDepartCommandTime);
                }

                if (rowMainStreamCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowMainStreamCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowMainStreamCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByMainStreamCommandCount);
                }

                if (rowMainStreamTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowMainStreamTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowMainStreamTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByMainStreamCommandTime);
                }

                if (rowStationCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowStationCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowStationCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByStationCount);
                }

                if (rowStationTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowStationTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowStationTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByStationTime);
                }

                if (rowWeatherCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowWeatherCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowWeatherCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByWeatherCount);
                }

                if (rowWeatherTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowWeatherTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowWeatherTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByWeatherTime);
                }

                if (rowUnitCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowUnitCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowUnitCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNotAskCount);
                }

                if (rowUnitTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowUnitTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowUnitTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNotAskTime);
                }
            }

            repairText = "";
            int count = 1;
            //设备单位
            foreach (ControllersAndDeparts _cod in repairDepart)
            {

                //把单位名称填上
                if(sheetRepair.GetRow(mainFileTitle[0]._plannedCount_row - 1).GetCell(_cod._codColumn) == null)
                {
                    sheetRepair.CreateRow(mainFileTitle[0]._plannedCount_row - 1).CreateCell(_cod._codColumn).SetCellValue(_cod._codName);
                }
                else
                {
                    sheetRepair.GetRow(mainFileTitle[0]._plannedCount_row - 1).GetCell(_cod._codColumn).SetCellValue(_cod._codName);
                }
                repairText = repairText + count + "、" + _cod._codName + "\n";
                count++;

                //1-4行为普通情况
                if (rowPlanCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowPlanCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPlanCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPlannedCount);
                }

                if (rowPlanTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowPlanTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPlanTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPlannedTime);
                }

                if (rowPermitCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowPermitCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPermitCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPermitCount);
                }

                if (rowPermitTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowPermitTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPermitTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPermitTime);
                }

                if (rowAccidentCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowAccidentCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowAccidentCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByAccidentCount);
                }

                if (rowAccidentTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowAccidentTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowAccidentTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByAccidentTime);
                }

                if (rowNatureCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowNatureCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowNatureCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNatureCount);
                }

                if (rowNatureTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowNatureTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowNatureTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNatureTime);
                }

                if (rowDepartCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowDepartCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowDepartCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByDepartCommandCount);
                }

                if (rowDepartTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowDepartTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowDepartTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByDepartCommandTime);
                }

                if (rowMainStreamCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowMainStreamCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowMainStreamCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByMainStreamCommandCount);
                }

                if (rowMainStreamTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowMainStreamTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowMainStreamTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByMainStreamCommandTime);
                }

                if (rowStationCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowStationCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowStationCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByStationCount);
                }

                if (rowStationTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowStationTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowStationTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByStationTime);
                }

                if (rowWeatherCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowWeatherCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowWeatherCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByWeatherCount);
                }

                if (rowWeatherTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowWeatherTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowWeatherTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByWeatherTime);
                }

                if (rowUnitCount_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowUnitCount_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowUnitCount_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNotAskCount);
                }

                if (rowUnitTime_repair != null && _cod._codColumn != 0)
                {
                    ICell cell = rowUnitTime_repair.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowUnitTime_repair.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNotAskTime);
                }
            }

            /*重新修改文件指定单元格样式*/
            FileStream fs1 = File.OpenWrite(mainFile);
            workbook.Write(fs1);
            fs1.Close();
            fileStream.Close();
            workbook.Close();

            sw.Stop();
            TimeSpan ts2 = sw.Elapsed;
            procceingProgress = "处理完成，表格将打开。共耗时" + ts2.TotalSeconds+ "秒。";
            sw.Reset();
            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
            //info.WorkingDirectory = Application.StartupPath;
            info.FileName = mainFile;
            info.Arguments = "";
            try
            {
                System.Diagnostics.Process.Start(info);
            }
            catch (System.ComponentModel.Win32Exception we)
            {
                MessageBox.Show(this, we.Message);
                return;
            }
        }
       
        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void readMainFile_btn_Click(object sender, EventArgs e)
        {
            selectPath(0);
        }

        private void readSubFile_btn_Click(object sender, EventArgs e)
        {
            selectPath(1);
        }

        //实现进度
        private void SetTextMesssage(int ipos, string vinfo)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    SetPos setpos = new SetPos(SetTextMesssage);
                    this.Invoke(setpos, new object[] { ipos, vinfo });
                }
                else
                {
                    this.processing_lbl.Text = procceingProgress;
                    if(stationController_rtb.Text .Length == 0)
                    {
                        int _counter = 1;
                        string allTemp = "";
                        foreach (string _temp in stationInController)
                        {
                            allTemp = allTemp + _counter.ToString() + "、" + _temp.Replace("_", "->") + "\n";
                            stationController_rtb.Text = allTemp;
                            _counter++;
                        }
                    }
                    if(repairDeaprtList_rtb.Text.Length == 0)
                    {
                        repairDeaprtList_rtb.Text = repairText;
                    }
                    if(stationsList_rtb.Text.Length == 0)
                    {
                        stationsList_rtb.Text = "共"+stationCount+"个车站\n"+s_text;
                    }
                }
            }
            catch(Exception e)
            {
                fThread.Abort();
            }

        }

        private void SleepT()
        {
            for (int i = 0; i < 5000; i++)
            {
                System.Threading.Thread.Sleep(20);
                SetTextMesssage(20 * i / 100, i.ToString() + "\r\n");
            }
        }

        private void start_btn_Click(object sender, EventArgs e)
        {
            refresh();
            readMainFile();
            readSubFiles();
            if (stationController_rtb.Text.Length == 0)
            {
                int _counter = 1;
                string allTemp = "";
                foreach (string _temp in stationInController)
                {
                    allTemp = allTemp + _counter.ToString() + "、" + _temp.Replace("_", "->") + "\n";
                    stationController_rtb.Text = allTemp;
                    _counter++;
                }
            }
            if (repairDeaprtList_rtb.Text.Length == 0)
            {
                repairDeaprtList_rtb.Text = repairText;
            }
            if (stationsList_rtb.Text.Length == 0)
            {
                stationsList_rtb.Text = "共"+stationCount+"个车站\n"+ s_text;
            }
            /*
            if (!fThread.IsAlive)
            {
                sw.Start();
                refresh();
                fThread = new Thread(new ThreadStart(SleepT));
                fThread.Start();
                Thread readMainFileThread = new Thread(new ThreadStart(readMainFile));
                readMainFileThread.Start();
                Thread readSubFileThread = new Thread(new ThreadStart(readSubFiles));
                readSubFileThread.Start();
            }
            */

        }

        private void label3_Click(object sender, EventArgs e)
        {
            Info form = new Info();
            form.Show();
        }
    }
}
