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
        //副文件标题
        //List<SubFileTitle> subFileTitles;
        //施工
        List<Construction> constructions;
        //维修
        List<RailRepair> railRepairs;
        bool hasFilePath = false;

        public Main()
        {
            mainFileTitle = new List<MainFileTitles>();
            //subFileTitles = new List<SubFileTitle>();
            refresh();
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            start_btn.Enabled = false;
        }

        private void refresh()
        {
            constructions = new List<Construction>();
            railRepairs = new List<RailRepair>();
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
                    this.mainExcelFile_lbl.Text = "已选择：" + mainFile;
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
            MainFileTitles consTitle = mainFileTitle[0];
            MainFileTitles repairTitle = mainFileTitle[1];
            //先找施工的
            consTitle = findMainTitles(constructionSheet, consTitle);
            //再找维修的
            repairTitle = findMainTitles(repairSheet, repairTitle);
            fileStream.Close();
        }

        private MainFileTitles findMainTitles(ISheet _sheet ,MainFileTitles _mainTitle)
        {
            int _statisticController = 0;
            int _statisticsDepart = 0;
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
                        if (cell.ToString().Contains("基数"))
                        {
                            _mainTitle._plannedCount_row = rowNum;
                            _mainTitle._plannedTime_row = rowNum + 1;
                        }
                        if (cell.ToString().Contains("给点"))
                        {
                            _mainTitle._permitCount_row = rowNum;
                            _mainTitle._permitTime_row = rowNum + 1;
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

        //读副图
        private void readSubFiles()
        {
            foreach(string _subFile in subFileList)
            {
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
                //施工表头
                for (int rowNum = 0; rowNum<=constructionSheet.LastRowNum;rowNum++)
                {
                    IRow row = constructionSheet.GetRow(rowNum);
                    for(int columnNum = 0; columnNum<= row.LastCellNum; columnNum++)
                    {
                        if (row.GetCell(columnNum).ToString().Contains("提报车站"))
                        {
                            currentStation = row.GetCell(columnNum + 1).ToString();
                        }
                        if (row.GetCell(columnNum).ToString().Contains("计划"))
                        {
                            plannedColumn_cons = columnNum;
                        }
                        if (row.GetCell(columnNum).ToString().Contains("申请"))
                        {
                            askColumn_cons = columnNum;
                        }
                        if (row.GetCell(columnNum).ToString().Contains("给点"))
                        {
                            permitColumn_cons = columnNum;
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
                //施工内容
                for (int rowNum = 0; rowNum <= constructionSheet.LastRowNum; rowNum++)
                {
                    IRow row = constructionSheet.GetRow(rowNum);
                    ICell cell = row.GetCell(dateColumn_cons);
                    if(cell != null && cell.ToString().Length != 0)
                    {
                        int date = 0;
                        int.TryParse(cell.ToString(), out date);
                        if (date > 0)
                        {
                            Construction _cons = new Construction();
                            _cons._constructionDate = date;

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

                            cell = row.GetCell(askColumn_cons +1);
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

                            //往右继续找,每找到一个之后向上找对应标题行

                        }
                    }
                }
                //然后是维修表头
                for (int rowNum = 0; rowNum <= repairSheet.LastRowNum; rowNum++)
                {
                    IRow row = constructionSheet.GetRow(rowNum);
                    for (int columnNum = 0; columnNum <= row.LastCellNum; columnNum++)
                    {
                        if (row.GetCell(columnNum).ToString().Contains("提报车站"))
                        {
                            currentStation = row.GetCell(columnNum + 1).ToString();
                        }
                        if (row.GetCell(columnNum).ToString().Contains("计划"))
                        {
                            plannedColumn_repair = columnNum;
                        }
                        if (row.GetCell(columnNum).ToString().Contains("申请"))
                        {
                            askColumn_repair = columnNum;
                        }
                        if (row.GetCell(columnNum).ToString().Contains("给点"))
                        {
                            permitColumn_repair = columnNum;
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
                for (int rowNum = 0; rowNum <= constructionSheet.LastRowNum; rowNum++)
                {
                    IRow row = constructionSheet.GetRow(rowNum);
                    for (int columnNum = 0; columnNum <= row.LastCellNum; columnNum++)
                    {

                    }
                }
                fileStream.Close();
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

        private void start_btn_Click(object sender, EventArgs e)
        {
            readMainFile();
        }
    }
}
