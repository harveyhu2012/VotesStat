namespace VotesStat
{
    partial class FormMain
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
            this.buttonLoad = new System.Windows.Forms.Button();
            this.buttonCal = new System.Windows.Forms.Button();
            this.checkedListBoxDept = new System.Windows.Forms.CheckedListBox();
            this.checkedListBoxVoteds = new System.Windows.Forms.CheckedListBox();
            this.labelMode = new System.Windows.Forms.Label();
            this.buttonClear = new System.Windows.Forms.Button();
            this.listBoxFiles = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonLoad
            // 
            this.buttonLoad.Location = new System.Drawing.Point(31, 23);
            this.buttonLoad.Name = "buttonLoad";
            this.buttonLoad.Size = new System.Drawing.Size(91, 30);
            this.buttonLoad.TabIndex = 0;
            this.buttonLoad.Text = "读取excel";
            this.buttonLoad.UseVisualStyleBackColor = true;
            this.buttonLoad.Click += new System.EventHandler(this.buttonLoad_Click);
            // 
            // buttonCal
            // 
            this.buttonCal.Location = new System.Drawing.Point(128, 23);
            this.buttonCal.Name = "buttonCal";
            this.buttonCal.Size = new System.Drawing.Size(91, 30);
            this.buttonCal.TabIndex = 1;
            this.buttonCal.Text = "计算评分";
            this.buttonCal.UseVisualStyleBackColor = true;
            this.buttonCal.Click += new System.EventHandler(this.buttonCal_Click);
            // 
            // checkedListBoxDept
            // 
            this.checkedListBoxDept.CheckOnClick = true;
            this.checkedListBoxDept.FormattingEnabled = true;
            this.checkedListBoxDept.Location = new System.Drawing.Point(31, 230);
            this.checkedListBoxDept.Name = "checkedListBoxDept";
            this.checkedListBoxDept.Size = new System.Drawing.Size(285, 388);
            this.checkedListBoxDept.TabIndex = 2;
            this.checkedListBoxDept.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBoxDept_ItemCheck);
            // 
            // checkedListBoxVoteds
            // 
            this.checkedListBoxVoteds.CheckOnClick = true;
            this.checkedListBoxVoteds.FormattingEnabled = true;
            this.checkedListBoxVoteds.Location = new System.Drawing.Point(345, 230);
            this.checkedListBoxVoteds.Name = "checkedListBoxVoteds";
            this.checkedListBoxVoteds.Size = new System.Drawing.Size(285, 388);
            this.checkedListBoxVoteds.TabIndex = 3;
            this.checkedListBoxVoteds.DoubleClick += new System.EventHandler(this.checkedListBoxVoteds_DoubleClick);
            // 
            // labelMode
            // 
            this.labelMode.AutoSize = true;
            this.labelMode.Location = new System.Drawing.Point(465, 33);
            this.labelMode.Name = "labelMode";
            this.labelMode.Size = new System.Drawing.Size(41, 12);
            this.labelMode.TabIndex = 4;
            this.labelMode.Text = "模式：";
            // 
            // buttonClear
            // 
            this.buttonClear.Location = new System.Drawing.Point(225, 23);
            this.buttonClear.Name = "buttonClear";
            this.buttonClear.Size = new System.Drawing.Size(91, 30);
            this.buttonClear.TabIndex = 5;
            this.buttonClear.Text = "清空数据";
            this.buttonClear.UseVisualStyleBackColor = true;
            this.buttonClear.Click += new System.EventHandler(this.buttonClear_Click);
            // 
            // listBoxFiles
            // 
            this.listBoxFiles.FormattingEnabled = true;
            this.listBoxFiles.ItemHeight = 12;
            this.listBoxFiles.Location = new System.Drawing.Point(31, 92);
            this.listBoxFiles.Name = "listBoxFiles";
            this.listBoxFiles.Size = new System.Drawing.Size(599, 100);
            this.listBoxFiles.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 77);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "已读文件列表：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 215);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 8;
            this.label2.Text = "选择科室：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(345, 215);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "选择成员:";
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(696, 642);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBoxFiles);
            this.Controls.Add(this.buttonClear);
            this.Controls.Add(this.labelMode);
            this.Controls.Add(this.checkedListBoxVoteds);
            this.Controls.Add(this.checkedListBoxDept);
            this.Controls.Add(this.buttonCal);
            this.Controls.Add(this.buttonLoad);
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "员工考评数据统计";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonLoad;
        private System.Windows.Forms.Button buttonCal;
        private System.Windows.Forms.CheckedListBox checkedListBoxDept;
        private System.Windows.Forms.CheckedListBox checkedListBoxVoteds;
        private System.Windows.Forms.Label labelMode;
        private System.Windows.Forms.Button buttonClear;
        private System.Windows.Forms.ListBox listBoxFiles;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}

