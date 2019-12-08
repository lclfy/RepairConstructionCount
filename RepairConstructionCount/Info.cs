using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CCWin;

namespace RepairConstructionCount
{
    public partial class Info : Skin_Mac
    {
        public Info()
        {
            InitializeComponent();
        }

        private void Info_Load(object sender, EventArgs e)
        {
            label1.Text = "使用说明：\n1、调度台信息位于月度总表中“统计设置”，仅读取“站场名称”与“所属调度台”两列，其余内容可不填。\n" +
                "请确保“站场名称”与各站点所在的子文件中标题“提报车站”保持一致，识别为全字匹配。\n" +
                "2、处理前请手动删除月度总表内各施工单位名称，避免覆盖填写\n" +
                "3、天窗修单位为自动识别（方法与施工一样），因此名称会改为各单位在子表中的名称(可在各子表格中全部修改)\n" +
                "在“统计设置”中设置的单位不起作用，或等处理完后在总表内将“工务（洛）”等改掉即可\n" +
                "4、小计未进行计算，点击有数据的单元格可自动计算，或输入公式计算";
            label2.Text = "施工维修天窗统计工具 build20191207\n" +
                "Contact：17638570597";
        }

    }
}
