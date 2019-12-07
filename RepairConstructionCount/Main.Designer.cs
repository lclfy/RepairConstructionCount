namespace RepairConstructionCount
{
    partial class Main
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.readMainFile_btn = new CCWin.SkinControl.SkinButton();
            this.readSubFile_btn = new CCWin.SkinControl.SkinButton();
            this.start_btn = new CCWin.SkinControl.SkinButton();
            this.mainExcelFile_lbl = new System.Windows.Forms.Label();
            this.subExcelFile_lbl = new System.Windows.Forms.Label();
            this.processing_lbl = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.stationController_rtb = new System.Windows.Forms.RichTextBox();
            this.constructDepartList_rtb = new System.Windows.Forms.RichTextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.repairDeaprtList_rtb = new System.Windows.Forms.RichTextBox();
            this.repairDeaprtList_lbl = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // readMainFile_btn
            // 
            this.readMainFile_btn.BackColor = System.Drawing.Color.Transparent;
            this.readMainFile_btn.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.readMainFile_btn.DownBack = null;
            this.readMainFile_btn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.readMainFile_btn.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.readMainFile_btn.Location = new System.Drawing.Point(638, 114);
            this.readMainFile_btn.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.readMainFile_btn.MouseBack = null;
            this.readMainFile_btn.Name = "readMainFile_btn";
            this.readMainFile_btn.NormlBack = null;
            this.readMainFile_btn.Size = new System.Drawing.Size(202, 66);
            this.readMainFile_btn.TabIndex = 1;
            this.readMainFile_btn.Text = "读总表文件";
            this.readMainFile_btn.UseVisualStyleBackColor = false;
            this.readMainFile_btn.Click += new System.EventHandler(this.readMainFile_btn_Click);
            // 
            // readSubFile_btn
            // 
            this.readSubFile_btn.BackColor = System.Drawing.Color.Transparent;
            this.readSubFile_btn.BaseColor = System.Drawing.Color.Crimson;
            this.readSubFile_btn.BorderColor = System.Drawing.Color.Crimson;
            this.readSubFile_btn.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.readSubFile_btn.DownBack = null;
            this.readSubFile_btn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.readSubFile_btn.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.readSubFile_btn.Location = new System.Drawing.Point(851, 114);
            this.readSubFile_btn.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.readSubFile_btn.MouseBack = null;
            this.readSubFile_btn.Name = "readSubFile_btn";
            this.readSubFile_btn.NormlBack = null;
            this.readSubFile_btn.Size = new System.Drawing.Size(202, 66);
            this.readSubFile_btn.TabIndex = 2;
            this.readSubFile_btn.Text = "(全选)读子表文件";
            this.readSubFile_btn.UseVisualStyleBackColor = false;
            this.readSubFile_btn.Click += new System.EventHandler(this.readSubFile_btn_Click);
            // 
            // start_btn
            // 
            this.start_btn.BackColor = System.Drawing.Color.Transparent;
            this.start_btn.BaseColor = System.Drawing.Color.OrangeRed;
            this.start_btn.BorderColor = System.Drawing.Color.OrangeRed;
            this.start_btn.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.start_btn.DownBack = null;
            this.start_btn.DownBaseColor = System.Drawing.Color.Red;
            this.start_btn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.start_btn.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.start_btn.Location = new System.Drawing.Point(638, 191);
            this.start_btn.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.start_btn.MouseBack = null;
            this.start_btn.MouseBaseColor = System.Drawing.Color.DarkSalmon;
            this.start_btn.Name = "start_btn";
            this.start_btn.NormlBack = null;
            this.start_btn.Size = new System.Drawing.Size(416, 74);
            this.start_btn.TabIndex = 3;
            this.start_btn.Text = "执行";
            this.start_btn.UseVisualStyleBackColor = false;
            this.start_btn.Click += new System.EventHandler(this.start_btn_Click);
            // 
            // mainExcelFile_lbl
            // 
            this.mainExcelFile_lbl.AutoSize = true;
            this.mainExcelFile_lbl.Location = new System.Drawing.Point(79, 137);
            this.mainExcelFile_lbl.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.mainExcelFile_lbl.Name = "mainExcelFile_lbl";
            this.mainExcelFile_lbl.Size = new System.Drawing.Size(115, 21);
            this.mainExcelFile_lbl.TabIndex = 4;
            this.mainExcelFile_lbl.Text = "总表文件：";
            // 
            // subExcelFile_lbl
            // 
            this.subExcelFile_lbl.AutoSize = true;
            this.subExcelFile_lbl.Location = new System.Drawing.Point(79, 193);
            this.subExcelFile_lbl.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.subExcelFile_lbl.Name = "subExcelFile_lbl";
            this.subExcelFile_lbl.Size = new System.Drawing.Size(115, 21);
            this.subExcelFile_lbl.TabIndex = 5;
            this.subExcelFile_lbl.Text = "分表文件：";
            this.subExcelFile_lbl.Click += new System.EventHandler(this.label2_Click);
            // 
            // processing_lbl
            // 
            this.processing_lbl.AutoSize = true;
            this.processing_lbl.Location = new System.Drawing.Point(79, 285);
            this.processing_lbl.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.processing_lbl.Name = "processing_lbl";
            this.processing_lbl.Size = new System.Drawing.Size(115, 21);
            this.processing_lbl.TabIndex = 6;
            this.processing_lbl.Text = "正在处理：";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(79, 360);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(304, 21);
            this.label1.TabIndex = 8;
            this.label1.Text = "“统计设置”中的调度车站关系";
            // 
            // stationController_rtb
            // 
            this.stationController_rtb.Location = new System.Drawing.Point(83, 420);
            this.stationController_rtb.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.stationController_rtb.Name = "stationController_rtb";
            this.stationController_rtb.ReadOnly = true;
            this.stationController_rtb.Size = new System.Drawing.Size(299, 518);
            this.stationController_rtb.TabIndex = 9;
            this.stationController_rtb.Text = "";
            // 
            // constructDepartList_rtb
            // 
            this.constructDepartList_rtb.Location = new System.Drawing.Point(424, 420);
            this.constructDepartList_rtb.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.constructDepartList_rtb.Name = "constructDepartList_rtb";
            this.constructDepartList_rtb.ReadOnly = true;
            this.constructDepartList_rtb.Size = new System.Drawing.Size(299, 518);
            this.constructDepartList_rtb.TabIndex = 11;
            this.constructDepartList_rtb.Text = "";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(765, 360);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(284, 21);
            this.label2.TabIndex = 10;
            this.label2.Text = "(已找到的)天窗作业单位列表";
            // 
            // repairDeaprtList_rtb
            // 
            this.repairDeaprtList_rtb.Location = new System.Drawing.Point(750, 420);
            this.repairDeaprtList_rtb.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.repairDeaprtList_rtb.Name = "repairDeaprtList_rtb";
            this.repairDeaprtList_rtb.ReadOnly = true;
            this.repairDeaprtList_rtb.Size = new System.Drawing.Size(299, 518);
            this.repairDeaprtList_rtb.TabIndex = 13;
            this.repairDeaprtList_rtb.Text = "";
            // 
            // repairDeaprtList_lbl
            // 
            this.repairDeaprtList_lbl.AutoSize = true;
            this.repairDeaprtList_lbl.Location = new System.Drawing.Point(458, 360);
            this.repairDeaprtList_lbl.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.repairDeaprtList_lbl.Name = "repairDeaprtList_lbl";
            this.repairDeaprtList_lbl.Size = new System.Drawing.Size(242, 21);
            this.repairDeaprtList_lbl.TabIndex = 12;
            this.repairDeaprtList_lbl.Text = "(已找到的)施工单位列表";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 10.71429F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label3.Location = new System.Drawing.Point(846, 68);
            this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(216, 26);
            this.label3.TabIndex = 14;
            this.label3.Text = "<关于/使用说明>";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1159, 999);
            this.Controls.Add(this.readSubFile_btn);
            this.Controls.Add(this.readMainFile_btn);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.repairDeaprtList_rtb);
            this.Controls.Add(this.repairDeaprtList_lbl);
            this.Controls.Add(this.constructDepartList_rtb);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.stationController_rtb);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.processing_lbl);
            this.Controls.Add(this.subExcelFile_lbl);
            this.Controls.Add(this.mainExcelFile_lbl);
            this.Controls.Add(this.start_btn);
            this.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.Name = "Main";
            this.Text = "施工维修天窗统计工具";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private CCWin.SkinControl.SkinButton readMainFile_btn;
        private CCWin.SkinControl.SkinButton readSubFile_btn;
        private CCWin.SkinControl.SkinButton start_btn;
        private System.Windows.Forms.Label mainExcelFile_lbl;
        private System.Windows.Forms.Label subExcelFile_lbl;
        private System.Windows.Forms.Label processing_lbl;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RichTextBox stationController_rtb;
        private System.Windows.Forms.RichTextBox constructDepartList_rtb;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RichTextBox repairDeaprtList_rtb;
        private System.Windows.Forms.Label repairDeaprtList_lbl;
        private System.Windows.Forms.Label label3;
    }
}

