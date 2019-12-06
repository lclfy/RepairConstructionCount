using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CCWin;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;
using System.Threading;

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
        //List<SubFileTitle> subFileTitles;
        //施工
        List<Construction> constructions;
        //维修
        List<RailRepair> railRepairs;
        //调度台(施工)
        List<ControllersAndDeparts> consControllers;
        //调度台(维修)
        List<ControllersAndDeparts> repairControllers;
        //单位(施工)
        List<ControllersAndDeparts> constructDepart;
        //单位(天窗)
        List<ControllersAndDeparts> repairDepart;
        bool hasFilePath = false;
        string procceingProgress = "";
        //显示进度
        private delegate void SetPos(int ipos, string vinfo);
        Thread fThread;

        public Main()
        {
            mainFileTitle = new List<MainFileTitles>();
            //subFileTitles = new List<SubFileTitle>();
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            refresh();
            fThread = new Thread(new ThreadStart(SleepT));
            stationInController = new List<string>();
            consControllers = new List<ControllersAndDeparts>();
            repairControllers = new List<ControllersAndDeparts>();
            constructDepart = new List<ControllersAndDeparts>();
            repairDepart = new List<ControllersAndDeparts>();
            start_btn.Enabled = false;
            processing_lbl.Text = procceingProgress;
        }

        private void refresh()
        {
            constructions = new List<Construction>();
            railRepairs = new List<RailRepair>();
            stationInController = new List<string>();
            consControllers = new List<ControllersAndDeparts>();
            repairControllers = new List<ControllersAndDeparts>();
            constructDepart = new List<ControllersAndDeparts>();
            repairDepart = new List<ControllersAndDeparts>();
            stationController_rtb.Clear();
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
                        MainFileTitles constructtionTitle = new MainFileTitles();
                        constructtionTitle._fileName = fileName;
                        constructtionTitle._repairOrConstruction = false;
                        mainFileTitle.Add(constructtionTitle);
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
                   // subFileTitles = new List<SubFileTitle>();
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
                ISheet constructionSheet = workbook.GetSheet("施工天窗");
                ISheet repairSheet = workbook.GetSheet("高铁天窗");
                ISheet stationControllerSheet = workbook.GetSheet("统计设置");
                MainFileTitles consTitle = (MainFileTitles)mainFileTitle[0].Clone();
                MainFileTitles repairTitle = (MainFileTitles)mainFileTitle[1].Clone();
                mainFileTitle.Clear();
                //先找施工的
                consTitle = findMainTitles(constructionSheet, consTitle);
                //再找维修的
                repairTitle = findMainTitles(repairSheet, repairTitle);
                mainFileTitle.Add(consTitle);
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
            procceingProgress = "正在处理：" + mainFile.Split('\\')[mainFile.Split('\\').Count() - 1];
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
                foreach(ControllersAndDeparts _cod in consControllers)
                {
                    if (_cod._codName.Equals(_tempController._codName))
                    {
                        findSame = true;
                        break;
                    }
                }
                if (!findSame)
                {
                    consControllers.Add((ControllersAndDeparts)_tempController.Clone());
                    repairControllers.Add((ControllersAndDeparts)_tempController.Clone());
                }
            }
            return _StationController;
        }

        //读副图
        private void readSubFiles()
        {
            try
            {
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
                    ISheet constructionSheet = workbook.GetSheet("施工天窗");
                    ISheet repairSheet = workbook.GetSheet("高铁天窗");
                    //直接找，先施工天窗
                    //次数在左，时间在右
                    string currentStation = "";
                    int plannedColumn_cons = 0;
                    int askColumn_cons = 0;
                    int permitColumn_cons = 0;
                    int dateColumn_cons = 0;
                    int departCloumn_cons = 0;
                    int titleRow_cons = 0;

                    int plannedColumn_repair = 0;
                    int askColumn_repair = 0;
                    int permitColumn_repair = 0;
                    int dateColumn_repair = 0;
                    int departCloumn_repair = 0;
                    int titleRow_repair = 0;
                    //居然有多个“给点”。。取第一个
                    bool hasGotIt = false;
                    //施工表头
                    for (int rowNum = 0; rowNum <= constructionSheet.LastRowNum; rowNum++)
                    {
                        IRow row = constructionSheet.GetRow(rowNum);
                        if (row == null)
                        {
                            continue;
                        }

                        for (int columnNum = 0; columnNum <= row.LastCellNum; columnNum++)
                        {
                            if (row.GetCell(columnNum) == null || row.GetCell(columnNum).ToString().Length == 0)
                            {
                                continue;
                            }
                            if (row.GetCell(columnNum).ToString().Contains("提报车站"))
                            {
                                currentStation = row.GetCell(columnNum + 2).ToString();
                            }
                            if (row.GetCell(columnNum).ToString().Contains("计划"))
                            {
                                plannedColumn_cons = columnNum;
                            }
                            if (row.GetCell(columnNum).ToString().Contains("申请"))
                            {
                                askColumn_cons = columnNum;
                            }
                            if (row.GetCell(columnNum).ToString().Contains("给点") && !hasGotIt)
                            {
                                permitColumn_cons = columnNum;
                                hasGotIt = true;
                            }
                            if (row.GetCell(columnNum).ToString().Contains("日期"))
                            {
                                dateColumn_cons = columnNum;
                                titleRow_cons = rowNum;
                            }
                            if (row.GetCell(columnNum).ToString().Contains("单位"))
                            {
                                departCloumn_cons = columnNum;
                            }

                        }
                    }
                    hasGotIt = false;
                    //施工内容
                    for (int rowNum = 0; rowNum <= constructionSheet.LastRowNum; rowNum++)
                    {
                        IRow row = constructionSheet.GetRow(rowNum);
                        //标题行，名称行
                        if (row == null)
                        {
                            continue;
                        }
                        IRow titleRow = constructionSheet.GetRow(titleRow_cons);
                        IRow nameRow = constructionSheet.GetRow(titleRow_cons - 1);
                        ICell cell = row.GetCell(dateColumn_cons);
                        if (cell != null && cell.ToString().Length != 0)
                        {
                            int date = 0;
                            int.TryParse(cell.ToString(), out date);
                            if (date > 0)
                            {
                                Construction _cons = new Construction();
                                _cons._constructionDate = date;
                                _cons._stationName = currentStation;
                                cell = row.GetCell(departCloumn_cons);
                                _cons._departName = cell.ToString();

                                //计划
                                int temp = 0;
                                cell = row.GetCell(plannedColumn_cons);
                                int.TryParse(cell.ToString(), out temp);
                                _cons._plannedConstructionCount = temp;
                                temp = 0;

                                cell = row.GetCell(plannedColumn_cons + 1);
                                int.TryParse(cell.ToString(), out temp);
                                _cons._plannedConstructionTime = temp;
                                temp = 0;

                                //申请
                                cell = row.GetCell(askColumn_cons);
                                int.TryParse(cell.ToString(), out temp);
                                _cons._askConstructionCount = temp;
                                temp = 0;

                                cell = row.GetCell(askColumn_cons + 1);
                                int.TryParse(cell.ToString(), out temp);
                                _cons._askConstructionTime = temp;
                                temp = 0;

                                //给点
                                cell = row.GetCell(permitColumn_cons);
                                int.TryParse(cell.ToString(), out temp);
                                _cons._permitConstructionCount = temp;
                                temp = 0;

                                cell = row.GetCell(permitColumn_cons + 1);
                                int.TryParse(cell.ToString(), out temp);
                                _cons._permitConstructionTime = temp;
                                temp = 0;

                                //往右继续找,每找到一个之后向上找对应标题行(先找次数，右边一列就是时间)
                                List<SpecialCauses_ConsRepair> _spccList = new List<SpecialCauses_ConsRepair>();
                                for (int tempColumnNum = permitColumn_cons + 1; tempColumnNum <= row.LastCellNum; tempColumnNum++)
                                {
                                    SpecialCauses_ConsRepair _spcc = new SpecialCauses_ConsRepair();
                                    if (row.GetCell(tempColumnNum) != null && row.GetCell(tempColumnNum).ToString().Length != 0)
                                    {//找到了
                                        if (row.GetCell(tempColumnNum).ToString().Equals("0"))
                                        {
                                            continue;
                                        }
                                        //往上找标题栏看是什么
                                        if (titleRow.GetCell(tempColumnNum).ToString().Contains("次数"))
                                        {//是次数，意味着上面就是名称
                                            if (nameRow.GetCell(tempColumnNum) != null && nameRow.GetCell(tempColumnNum).ToString().Length != 0)
                                            {
                                                if (nameRow.GetCell(tempColumnNum).ToString().Contains("小计") || nameRow.GetCell(tempColumnNum).ToString().Contains("基数")
                                                    || nameRow.GetCell(tempColumnNum).ToString().Contains("给点"))
                                                {
                                                    continue;
                                                }
                                                _spcc._causesName = nameRow.GetCell(tempColumnNum).ToString();
                                                int _temp = 0;
                                                int.TryParse(row.GetCell(tempColumnNum).ToString(), out _temp);
                                                _spcc._causesCount = _temp;
                                                if (_temp == 0)
                                                {
                                                    continue;
                                                }
                                                _temp = 0;
                                                //时间在右边一格
                                                int.TryParse(row.GetCell(tempColumnNum + 1).ToString(), out _temp);
                                                _spcc._causesTime = _temp;
                                                _temp = 0;
                                                //根据名称判断是不是列入考核
                                                if (_spcc._causesName.Contains("事故抢险") || _spcc._causesName.Contains("自然灾害"))
                                                {
                                                    _spcc._examine = false;
                                                }
                                                else if (_spcc._causesName.Contains("部令取消") || _spcc._causesName.Contains("局令取消") ||
                                                    _spcc._causesName.Contains("天气影响") || _spcc._causesName.Contains("车站未给") ||
                                                    _spcc._causesName.Contains("单位未要"))
                                                {
                                                    _spcc._examine = true;

                                                }
                                                //添加进去
                                                _spccList.Add(_spcc);
                                            }
                                        }
                                    }
                                }
                                _cons._specialCauses = _spccList;
                                constructions.Add(_cons);
                            }
                        }
                    }
                    //然后是维修表头
                    for (int rowNum = 0; rowNum <= repairSheet.LastRowNum; rowNum++)
                    {
                        IRow row = repairSheet.GetRow(rowNum);
                        if (row == null)
                        {
                            continue;
                        }
                        for (int columnNum = 0; columnNum <= row.LastCellNum; columnNum++)
                        {
                            if (row.GetCell(columnNum) == null || row.GetCell(columnNum).ToString().Length == 0)
                            {
                                continue;
                            }
                            if (row.GetCell(columnNum).ToString().Contains("提报车站"))
                            {
                                currentStation = row.GetCell(columnNum + 2).ToString();
                            }
                            if (row.GetCell(columnNum).ToString().Contains("计划"))
                            {
                                plannedColumn_repair = columnNum;
                            }
                            if (row.GetCell(columnNum).ToString().Contains("申请"))
                            {
                                askColumn_repair = columnNum;
                            }
                            if (row.GetCell(columnNum).ToString().Contains("给点") && !hasGotIt)
                            {
                                permitColumn_repair = columnNum;
                                hasGotIt = true;
                            }
                            if (row.GetCell(columnNum).ToString().Contains("日期"))
                            {
                                dateColumn_repair = columnNum;
                                titleRow_repair = rowNum;
                            }
                            if (row.GetCell(columnNum).ToString().Contains("单位"))
                            {
                                departCloumn_repair = columnNum;
                            }
                        }
                    }
                    //维修内容
                    int lastDateRow = titleRow_repair + 1;
                    for (int rowNum = titleRow_repair + 1; rowNum <= repairSheet.LastRowNum; rowNum++)
                    {
                        //从标题行下面一行开始找
                        IRow row = repairSheet.GetRow(rowNum);
                        //标题行，名称行
                        if (row == null)
                        {
                            continue;
                        }
                        IRow titleRow = repairSheet.GetRow(titleRow_repair);
                        IRow nameRow = repairSheet.GetRow(titleRow_repair - 1);
                        //从单位栏开始找，找左上角日期(不是日期的跳过)
                        ICell cell = row.GetCell(dateColumn_repair + 1);
                        if (cell != null && cell.ToString().Length != 0)
                        {
                            int date = 0;
                            for (int tempRowNum = rowNum; tempRowNum >= lastDateRow; tempRowNum--)
                            {//往上找
                                IRow _dateRow = repairSheet.GetRow(tempRowNum);

                                if (_dateRow.GetCell(dateColumn_repair) == null && _dateRow.GetCell(dateColumn_repair).ToString().Length == 0)
                                {
                                    continue;
                                }
                                else
                                {
                                    if (!Regex.IsMatch(_dateRow.GetCell(dateColumn_repair).ToString(), "^[0-9]*$"))
                                    {//日期不是数字，跳过
                                        string a = _dateRow.GetCell(dateColumn_repair).ToString();
                                        break;
                                    }
                                    int.TryParse(_dateRow.GetCell(dateColumn_repair).ToString(), out date);
                                    if (date != 0)
                                    {
                                        lastDateRow = tempRowNum;
                                    }
                                }
                            }
                            if (date > 0)
                            {
                                //需要一直往下找直到下一个Date出现
                                RailRepair _repair = new RailRepair();
                                _repair._repairDate = date;
                                _repair._stationName = currentStation;
                                cell = row.GetCell(departCloumn_repair);
                                _repair._departName = cell.ToString();

                                //计划
                                int temp = 0;
                                cell = row.GetCell(plannedColumn_repair);
                                int.TryParse(cell.ToString(), out temp);
                                _repair._plannedRepairCount = temp;
                                temp = 0;

                                cell = row.GetCell(plannedColumn_repair + 1);
                                int.TryParse(cell.ToString(), out temp);
                                _repair._plannedRepairTime = temp;
                                temp = 0;

                                //申请
                                cell = row.GetCell(askColumn_repair);
                                int.TryParse(cell.ToString(), out temp);
                                _repair._askRepairCount = temp;
                                temp = 0;

                                cell = row.GetCell(askColumn_repair + 1);
                                int.TryParse(cell.ToString(), out temp);
                                _repair._askRepairTime = temp;
                                temp = 0;

                                //给点
                                cell = row.GetCell(permitColumn_repair);
                                int.TryParse(cell.ToString(), out temp);
                                _repair._permitRepairCount = temp;
                                temp = 0;

                                cell = row.GetCell(permitColumn_repair + 1);
                                int.TryParse(cell.ToString(), out temp);
                                _repair._permitRepairTime = temp;
                                temp = 0;

                                //往右继续找,每找到一个之后向上找对应标题行(先找次数，右边一列就是时间)
                                List<SpecialCauses_ConsRepair> _spccList = new List<SpecialCauses_ConsRepair>();
                                for (int tempColumnNum = permitColumn_repair + 1; tempColumnNum <= row.LastCellNum; tempColumnNum++)
                                {
                                    SpecialCauses_ConsRepair _spcc = new SpecialCauses_ConsRepair();
                                    if (row.GetCell(tempColumnNum) != null && row.GetCell(tempColumnNum).ToString().Length != 0)
                                    {//找到了
                                        if (row.GetCell(tempColumnNum).ToString().Equals("0"))
                                        {
                                            continue;
                                        }
                                        //往上找标题栏看是什么
                                        if (titleRow.GetCell(tempColumnNum).ToString().Contains("次数"))
                                        {//是次数，意味着上面就是名称
                                            if (nameRow.GetCell(tempColumnNum) != null && nameRow.GetCell(tempColumnNum).ToString().Length != 0)
                                            {
                                                if (nameRow.GetCell(tempColumnNum).ToString().Contains("小计") || nameRow.GetCell(tempColumnNum).ToString().Contains("基数")
                                                    || nameRow.GetCell(tempColumnNum).ToString().Contains("给点"))
                                                {
                                                    continue;
                                                }
                                                _spcc._causesName = nameRow.GetCell(tempColumnNum).ToString();
                                                int _temp = 0;
                                                int.TryParse(row.GetCell(tempColumnNum).ToString(), out _temp);
                                                _spcc._causesCount = _temp;
                                                if (_temp == 0)
                                                {
                                                    continue;
                                                }
                                                _temp = 0;
                                                //时间在右边一格
                                                int.TryParse(row.GetCell(tempColumnNum + 1).ToString(), out _temp);
                                                _spcc._causesTime = _temp;
                                                _temp = 0;
                                                //根据名称判断是不是列入考核
                                                if (_spcc._causesName.Contains("事故抢险") || _spcc._causesName.Contains("自然灾害"))
                                                {
                                                    _spcc._examine = false;
                                                }
                                                else if (_spcc._causesName.Contains("部令取消") || _spcc._causesName.Contains("局令取消") ||
                                                    _spcc._causesName.Contains("天气影响") || _spcc._causesName.Contains("车站未给") ||
                                                    _spcc._causesName.Contains("单位未要"))
                                                {
                                                    _spcc._examine = true;

                                                }
                                                //添加进去
                                                _spccList.Add(_spcc);
                                            }
                                        }
                                    }
                                }
                                _repair._specialCauses = _spccList;
                                if (_repair._repairDate > 0)
                                {//防止混入小计
                                    railRepairs.Add(_repair);
                                }
                            }
                        }
                    }
                    fileStream.Close();
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
            catch (Exception e)
            {
                fThread.Abort();
                MessageBox.Show("请关闭所有打开的已选文件，错误内容：" + e.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        //计算并填写
        //匹配调度台
        private void matchControllersWithStations()
        {
            procceingProgress = "正在处理…";
            for (int count = 0; count<consControllers.Count;count++)
                {
                foreach (string _stationControllers in stationInController)
                {
                    //如果调度台和字符匹配到
                    if (_stationControllers.Contains(consControllers[count]._codName))
                    {
                        //挨个匹配对象加入进去
                        foreach (Construction _c in constructions)
                        {
                            if (_c._stationName.Trim().Equals(_stationControllers.Split('_')[0].Trim()))
                            {//匹配到
                                if(_c._plannedConstructionCount != 0)
                                {
                                    int aa = 0;
                                }
                                consControllers[count]._codPlannedCount += _c._plannedConstructionCount;
                                consControllers[count]._codPlannedTime += _c._plannedConstructionTime;
                                consControllers[count]._codPermitCount += _c._permitConstructionCount;
                                consControllers[count]._codPermitTime += _c._permitConstructionTime;
                                //特殊
                                foreach(SpecialCauses_ConsRepair _spcc in _c._specialCauses)
                                {
                                    if (_spcc._causesName.Contains("事故影响"))
                                    {
                                        consControllers[count]._causeByAccidentCount += _spcc._causesCount;
                                        consControllers[count]._causeByAccidentTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("自然灾害"))
                                    {
                                        consControllers[count]._causeByNatureCount += _spcc._causesCount;
                                        consControllers[count]._causeByNatureTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("部令取消"))
                                    {
                                        consControllers[count]._causeByDepartCommandCount += _spcc._causesCount;
                                        consControllers[count]._causeByDepartCommandTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("局令取消"))
                                    {
                                        consControllers[count]._causeByMainStreamCommandCount += _spcc._causesCount;
                                        consControllers[count]._causeByMainStreamCommandTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("单位未要"))
                                    {
                                        consControllers[count]._causeByNotAskCount += _spcc._causesCount;
                                        consControllers[count]._causeByNotAskTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("天气影响"))
                                    {
                                        consControllers[count]._causeByWeatherCount += _spcc._causesCount;
                                        consControllers[count]._causeByWeatherTime += _spcc._causesTime;
                                    }
                                }
                            }
                        }
                    }
                }
            }
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
                                    if (_spcc._causesName.Contains("事故影响"))
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

        //遍历设备单位
        private void searchAllDeparts()
        {
            foreach(Construction _cons in constructions)
            {
                bool hasSame = false;
                foreach(ControllersAndDeparts _tempCod in constructDepart)
                {
                    if (_tempCod._codName.Equals(_cons._departName))
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
                    if(_cons._departName.Length != 0)
                    {
                        ControllersAndDeparts _cod = new ControllersAndDeparts();
                        _cod._codName = _cons._departName;
                        constructDepart.Add(_cod);
                    }
                }
            }

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
                        repairDepart.Add(_cod);
                    }

                }
            }
        }

        //匹配设备单位
        private void matchDepartsWithStations()
        {
            for (int count = 0; count < constructDepart.Count; count++)
            {
                        foreach (Construction _c in constructions)
                        {
                            if (_c._departName.Trim().Equals(constructDepart[count]._codName.Trim()))
                            {//匹配到
                                if (_c._plannedConstructionCount != 0)
                                {
                                    int aa = 0;
                                }
                                constructDepart[count]._codPlannedCount += _c._plannedConstructionCount;
                                constructDepart[count]._codPlannedTime += _c._plannedConstructionTime;
                                constructDepart[count]._codPermitCount += _c._permitConstructionCount;
                                constructDepart[count]._codPermitTime += _c._permitConstructionTime;
                        if (!constructDepart[count].extra.Contains(_c._stationName))
                        {
                            constructDepart[count].extra += _c._stationName + " ";
                        }

                        //特殊
                        foreach (SpecialCauses_ConsRepair _spcc in _c._specialCauses)
                                {
                                    if (_spcc._causesName.Contains("事故影响"))
                                    {
                                        constructDepart[count]._causeByAccidentCount += _spcc._causesCount;
                                        constructDepart[count]._causeByAccidentTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("自然灾害"))
                                    {
                                        constructDepart[count]._causeByNatureCount += _spcc._causesCount;
                                        constructDepart[count]._causeByNatureTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("部令取消"))
                                    {
                                        constructDepart[count]._causeByDepartCommandCount += _spcc._causesCount;
                                        constructDepart[count]._causeByDepartCommandTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("局令取消"))
                                    {
                                        constructDepart[count]._causeByMainStreamCommandCount += _spcc._causesCount;
                                        constructDepart[count]._causeByMainStreamCommandTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("单位未要"))
                                    {
                                        constructDepart[count]._causeByNotAskCount += _spcc._causesCount;
                                        constructDepart[count]._causeByNotAskTime += _spcc._causesTime;
                                    }
                                    if (_spcc._causesName.Contains("天气影响"))
                                    {
                                        constructDepart[count]._causeByWeatherCount += _spcc._causesCount;
                                        constructDepart[count]._causeByWeatherTime += _spcc._causesTime;
                                    }
                                }
                            }
                }
            }
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
                                    if (_spcc._causesName.Contains("事故影响"))
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

            //特殊行，先找施工后找维修
            int accidentCountRow_cons = 0;
            int accidentTimeRow_cons = 0;

            int natureCountRow_cons = 0;
            int natureTimeRow_cons = 0;

            int departComCountRow_cons = 0;
            int departComTimeRow_cons = 0;

            int mainStreamComCountRow_cons = 0;
            int mainStreamComTimeRow_cons = 0;

            int stationCountRow_cons = 0;
            int stationTimeRow_cons = 0;

            int weatherCountRow_cons = 0;
            int weatherTimeRow_cons = 0;

            int unitCountRow_cons = 0;
            int unitTimeRow_cons = 0;
            try
            {
                accidentCountRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("事故影响"))._specialCauseCount_rowOrColumn;
                accidentTimeRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("事故影响"))._specialCauseTime_rowOrColumn;
            }
            catch(Exception e)
            {

            }
            try
            {
                natureCountRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("自然灾害"))._specialCauseCount_rowOrColumn;
                natureTimeRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("自然灾害"))._specialCauseTime_rowOrColumn;
            }
            catch(Exception e)
            {

            }
            try
            {
                departComCountRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("部令取消"))._specialCauseCount_rowOrColumn;
                departComTimeRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("部令取消"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                mainStreamComCountRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("局令取消"))._specialCauseCount_rowOrColumn;
                mainStreamComTimeRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("局令取消"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                stationCountRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("车站取消"))._specialCauseCount_rowOrColumn;
                stationTimeRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("车站取消"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                weatherCountRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("天气影响"))._specialCauseCount_rowOrColumn;
                weatherTimeRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("天气影响"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                unitCountRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("单位未要"))._specialCauseCount_rowOrColumn;
                unitTimeRow_cons = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("单位未要"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }

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
                accidentCountRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("事故影响"))._specialCauseCount_rowOrColumn;
                accidentTimeRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("事故影响"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                natureCountRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("自然灾害"))._specialCauseCount_rowOrColumn;
                natureTimeRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("自然灾害"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                departComCountRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("部令取消"))._specialCauseCount_rowOrColumn;
                departComTimeRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("部令取消"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                mainStreamComCountRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("局令取消"))._specialCauseCount_rowOrColumn;
                mainStreamComTimeRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("局令取消"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                stationCountRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("车站取消"))._specialCauseCount_rowOrColumn;
                stationTimeRow_rep = mainFileTitle[0]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("车站取消"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                weatherCountRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("天气影响"))._specialCauseCount_rowOrColumn;
                weatherTimeRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("天气影响"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }
            try
            {
                unitCountRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("单位未要"))._specialCauseCount_rowOrColumn;
                unitTimeRow_rep = mainFileTitle[1]._specialCauses_title.Find(target => target._specialCauseName.Trim().Equals("单位未要"))._specialCauseTime_rowOrColumn;
            }
            catch (Exception e)
            {

            }

            MainFileTitles _mfCons = mainFileTitle[0];
            ISheet sheetCons = workbook.GetSheet("施工天窗");
            //先填调度台的
            IRow rowPlanCount_cons = sheetCons.GetRow(_mfCons._plannedCount_row);
            IRow rowPlanTime_cons = sheetCons.GetRow(_mfCons._plannedTime_row);

            IRow rowPermitCount_cons = sheetCons.GetRow(_mfCons._permitCount_row);
            IRow rowPermitTime_cons = sheetCons.GetRow(_mfCons._permitTime_row);

            IRow rowAccidentCount_cons = sheetCons.GetRow(accidentCountRow_cons);
            IRow rowAccidentTime_cons = sheetCons.GetRow(accidentTimeRow_cons);

            IRow rowNatureCount_cons = sheetCons.GetRow(natureCountRow_cons);
            IRow rowNatureTime_cons = sheetCons.GetRow(natureTimeRow_cons);

            IRow rowDepartCount_cons = sheetCons.GetRow(departComCountRow_cons);
            IRow rowDepartTime_cons = sheetCons.GetRow(departComTimeRow_cons);

            IRow rowMainStreamCount_cons = sheetCons.GetRow(mainStreamComCountRow_cons);
            IRow rowMainStreamTime_cons = sheetCons.GetRow(mainStreamComTimeRow_cons);

            IRow rowStationCount_cons = sheetCons.GetRow(stationCountRow_cons);
            IRow rowStationTime_cons = sheetCons.GetRow(stationTimeRow_cons);

            IRow rowWeatherCount_cons = sheetCons.GetRow(weatherCountRow_cons);
            IRow rowWeatherTime_cons = sheetCons.GetRow(weatherTimeRow_cons);

            IRow rowUnitCount_cons = sheetCons.GetRow(unitCountRow_cons);
            IRow rowUnitTime_cons = sheetCons.GetRow(unitTimeRow_cons);

            foreach (ControllersAndDeparts _cod in consControllers)
            {
                if(_cod._codColumn == 0)
                {
                    continue;
                }
                //1-4行为普通情况
                if (rowPlanCount_cons != null)
                {
                    ICell cell = rowPlanCount_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPlanCount_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPlannedCount);
                }

                if (rowPlanTime_cons != null)
                {
                    ICell cell = rowPlanTime_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPlanTime_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPlannedTime);
                }

                if (rowPermitCount_cons != null)
                {
                    ICell cell = rowPermitCount_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPermitCount_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPermitCount);
                }

                if (rowPermitTime_cons != null)
                {
                    ICell cell = rowPermitTime_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowPermitTime_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._codPermitTime);
                }

                if(rowAccidentCount_cons != null)
                {
                    ICell cell = rowAccidentCount_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowAccidentCount_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByAccidentCount);
                }

                if (rowAccidentTime_cons != null)
                {
                    ICell cell = rowAccidentTime_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowAccidentTime_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByAccidentTime);
                }

                if (rowNatureCount_cons != null)
                {
                    ICell cell = rowNatureCount_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowNatureCount_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNatureCount);
                }

                if (rowNatureTime_cons != null)
                {
                    ICell cell = rowNatureTime_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowNatureTime_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNatureTime);
                }

                if (rowDepartCount_cons != null)
                {
                    ICell cell = rowDepartCount_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowDepartCount_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByDepartCommandCount);
                }

                if (rowDepartTime_cons != null)
                {
                    ICell cell = rowDepartTime_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowDepartTime_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByDepartCommandTime);
                }

                if (rowMainStreamCount_cons != null)
                {
                    ICell cell = rowMainStreamCount_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowMainStreamCount_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByMainStreamCommandCount);
                }

                if (rowMainStreamTime_cons != null)
                {
                    ICell cell = rowMainStreamTime_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowMainStreamTime_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByMainStreamCommandTime);
                }

                if (rowStationCount_cons != null)
                {
                    ICell cell = rowStationCount_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowStationCount_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByStationCount);
                }

                if (rowStationTime_cons != null)
                {
                    ICell cell = rowStationTime_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowStationTime_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByStationTime);
                }

                if (rowWeatherCount_cons != null)
                {
                    ICell cell = rowWeatherCount_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowWeatherCount_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByWeatherCount);
                }

                if (rowWeatherTime_cons != null)
                {
                    ICell cell = rowWeatherTime_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowWeatherTime_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByWeatherTime);
                }

                if (rowUnitCount_cons != null)
                {
                    ICell cell = rowUnitCount_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowUnitCount_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNotAskCount);
                }

                if (rowUnitTime_cons != null)
                {
                    ICell cell = rowUnitTime_cons.GetCell(_cod._codColumn);
                    if (cell == null)
                    {
                        cell = rowUnitTime_cons.CreateCell(_cod._codColumn);
                    }
                    cell.SetCellValue(_cod._causeByNotAskTime);
                }

            }

            //

            ///维修
            MainFileTitles _mfRepair = mainFileTitle[1];
            ISheet sheetRepair = workbook.GetSheet("高铁天窗");
            IRow rowPlanCount_repair = sheetRepair.GetRow(_mfRepair._plannedCount_row);
            IRow rowPlanTime_repair = sheetRepair.GetRow(_mfRepair._plannedTime_row);

            IRow rowPermitCount_repair = sheetRepair.GetRow(_mfRepair._permitCount_row);
            IRow rowPermitTime_repair = sheetRepair.GetRow(_mfRepair._permitTime_row);

            IRow rowAccidentCount_repair = sheetCons.GetRow(accidentCountRow_rep);
            IRow rowAccidentTime_repair = sheetCons.GetRow(accidentTimeRow_rep);

            IRow rowNatureCount_repair = sheetCons.GetRow(natureCountRow_rep);
            IRow rowNatureTime_repair = sheetCons.GetRow(natureTimeRow_rep);

            IRow rowDepartCount_repair = sheetCons.GetRow(departComCountRow_rep);
            IRow rowDepartTime_repair = sheetCons.GetRow(departComTimeRow_rep);

            IRow rowMainStreamCount_repair = sheetCons.GetRow(mainStreamComCountRow_rep);
            IRow rowMainStreamTime_repair = sheetCons.GetRow(mainStreamComTimeRow_rep);

            IRow rowStationCount_repair = sheetCons.GetRow(stationCountRow_rep);
            IRow rowStationTime_repair = sheetCons.GetRow(stationTimeRow_rep);

            IRow rowWeatherCount_repair = sheetCons.GetRow(weatherCountRow_rep);
            IRow rowWeatherTime_repair = sheetCons.GetRow(weatherTimeRow_rep);

            IRow rowUnitCount_repair = sheetCons.GetRow(unitCountRow_rep);
            IRow rowUnitTime_repair = sheetCons.GetRow(unitTimeRow_rep);

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

            /*重新修改文件指定单元格样式*/
            FileStream fs1 = File.OpenWrite(mainFile);
            workbook.Write(fs1);
            fs1.Close();
            fileStream.Close();
            workbook.Close();

            procceingProgress = "处理完成";
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
            if (!fThread.IsAlive)
            {
                refresh();
                fThread = new Thread(new ThreadStart(SleepT));
                fThread.Start();
                Thread readMainFileThread = new Thread(new ThreadStart(readMainFile));
                readMainFileThread.Start();
                Thread readSubFilesThread = new Thread(new ThreadStart(readSubFiles));
                readSubFilesThread.Start();
            }
        }
    }
}
