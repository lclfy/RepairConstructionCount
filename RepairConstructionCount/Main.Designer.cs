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
            this.label2 = new System.Windows.Forms.Label();
            this.repairDeaprtList_rtb = new System.Windows.Forms.RichTextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.stationsList_rtb = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // readMainFile_btn
            // 
            this.readMainFile_btn.BackColor = System.Drawing.Color.Transparent;
            this.readMainFile_btn.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.readMainFile_btn.DownBack = null;
            this.readMainFile_btn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.readMainFile_btn.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.readMainFile_btn.Location = new System.Drawing.Point(696, 130);
            this.readMainFile_btn.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.readMainFile_btn.MouseBack = null;
            this.readMainFile_btn.Name = "readMainFile_btn";
            this.readMainFile_btn.NormlBack = null;
            this.readMainFile_btn.Size = new System.Drawing.Size(220, 75);
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
            this.readSubFile_btn.Location = new System.Drawing.Point(928, 130);
            this.readSubFile_btn.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.readSubFile_btn.MouseBack = null;
            this.readSubFile_btn.Name = "readSubFile_btn";
            this.readSubFile_btn.NormlBack = null;
            this.readSubFile_btn.Size = new System.Drawing.Size(220, 75);
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
            this.start_btn.Location = new System.Drawing.Point(696, 218);
            this.start_btn.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.start_btn.MouseBack = null;
            this.start_btn.MouseBaseColor = System.Drawing.Color.DarkSalmon;
            this.start_btn.Name = "start_btn";
            this.start_btn.NormlBack = null;
            this.start_btn.Size = new System.Drawing.Size(454, 85);
            this.start_btn.TabIndex = 3;
            this.start_btn.Text = "执行";
            this.start_btn.UseVisualStyleBackColor = false;
            this.start_btn.Click += new System.EventHandler(this.start_btn_Click);
            // 
            // mainExcelFile_lbl
            // 
            this.mainExcelFile_lbl.AutoSize = true;
            this.mainExcelFile_lbl.Location = new System.Drawing.Point(86, 157);
            this.mainExcelFile_lbl.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.mainExcelFile_lbl.Name = "mainExcelFile_lbl";
            this.mainExcelFile_lbl.Size = new System.Drawing.Size(130, 24);
            this.mainExcelFile_lbl.TabIndex = 4;
            this.mainExcelFile_lbl.Text = "总表文件：";
            // 
            // subExcelFile_lbl
            // 
            this.subExcelFile_lbl.AutoSize = true;
            this.subExcelFile_lbl.Location = new System.Drawing.Point(86, 221);
            this.subExcelFile_lbl.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.subExcelFile_lbl.Name = "subExcelFile_lbl";
            this.subExcelFile_lbl.Size = new System.Drawing.Size(130, 24);
            this.subExcelFile_lbl.TabIndex = 5;
            this.subExcelFile_lbl.Text = "分表文件：";
            this.subExcelFile_lbl.Click += new System.EventHandler(this.label2_Click);
            // 
            // processing_lbl
            // 
            this.processing_lbl.AutoSize = true;
            this.processing_lbl.Location = new System.Drawing.Point(86, 326);
            this.processing_lbl.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.processing_lbl.Name = "processing_lbl";
            this.processing_lbl.Size = new System.Drawing.Size(130, 24);
            this.processing_lbl.TabIndex = 6;
            this.processing_lbl.Text = "正在处理：";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(86, 411);
            this.label1.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(346, 24);
            this.label1.TabIndex = 8;
            this.label1.Text = "“统计设置”中的调度车站关系";
            // 
            // stationController_rtb
            // 
            this.stationController_rtb.Location = new System.Drawing.Point(91, 480);
            this.stationController_rtb.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.stationController_rtb.Name = "stationController_rtb";
            this.stationController_rtb.ReadOnly = true;
            this.stationController_rtb.Size = new System.Drawing.Size(341, 591);
            this.stationController_rtb.TabIndex = 9;
            this.stationController_rtb.Text = "";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(814, 411);
            this.label2.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(322, 24);
            this.label2.TabIndex = 10;
            this.label2.Text = "(已找到的)天窗作业单位列表";
            // 
            // repairDeaprtList_rtb
            // 
            this.repairDeaprtList_rtb.Location = new System.Drawing.Point(818, 480);
            this.repairDeaprtList_rtb.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.repairDeaprtList_rtb.Name = "repairDeaprtList_rtb";
            this.repairDeaprtList_rtb.ReadOnly = true;
            this.repairDeaprtList_rtb.Size = new System.Drawing.Size(326, 591);
            this.repairDeaprtList_rtb.TabIndex = 13;
            this.repairDeaprtList_rtb.Text = "";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 10.71429F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label3.Location = new System.Drawing.Point(865, 326);
            this.label3.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(285, 29);
            this.label3.TabIndex = 14;
            this.label3.Text = "<点此查看使用说明>";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 10.71429F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label4.Location = new System.Drawing.Point(964, 1089);
            this.label4.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(237, 29);
            this.label4.TabIndex = 15;
            this.label4.Text = "Build 20200609";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(457, 411);
            this.label5.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(226, 24);
            this.label5.TabIndex = 16;
            this.label5.Text = "(已找到的)车站列表";
            // 
            // stationsList_rtb
            // 
            this.stationsList_rtb.Location = new System.Drawing.Point(461, 480);
            this.stationsList_rtb.Name = "stationsList_rtb";
            this.stationsList_rtb.ReadOnly = true;
            this.stationsList_rtb.Size = new System.Drawing.Size(331, 591);
            this.stationsList_rtb.TabIndex = 17;
            this.stationsList_rtb.Text = "";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1264, 1142);
            this.Controls.Add(this.stationsList_rtb);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.readSubFile_btn);
            this.Controls.Add(this.readMainFile_btn);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.repairDeaprtList_rtb);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.stationController_rtb);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.processing_lbl);
            this.Controls.Add(this.subExcelFile_lbl);
            this.Controls.Add(this.mainExcelFile_lbl);
            this.Controls.Add(this.start_btn);
            this.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
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
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RichTextBox repairDeaprtList_rtb;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.RichTextBox stationsList_rtb;
    }
}

