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
            this.SuspendLayout();
            // 
            // buttonLoad
            // 
            this.buttonLoad.Location = new System.Drawing.Point(31, 32);
            this.buttonLoad.Name = "buttonLoad";
            this.buttonLoad.Size = new System.Drawing.Size(91, 30);
            this.buttonLoad.TabIndex = 0;
            this.buttonLoad.Text = "读取excel";
            this.buttonLoad.UseVisualStyleBackColor = true;
            this.buttonLoad.Click += new System.EventHandler(this.buttonLoad_Click);
            // 
            // buttonCal
            // 
            this.buttonCal.Location = new System.Drawing.Point(138, 33);
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
            this.checkedListBoxDept.Location = new System.Drawing.Point(31, 101);
            this.checkedListBoxDept.Name = "checkedListBoxDept";
            this.checkedListBoxDept.Size = new System.Drawing.Size(268, 388);
            this.checkedListBoxDept.TabIndex = 2;
            this.checkedListBoxDept.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBoxDept_ItemCheck);
            // 
            // checkedListBoxVoteds
            // 
            this.checkedListBoxVoteds.CheckOnClick = true;
            this.checkedListBoxVoteds.FormattingEnabled = true;
            this.checkedListBoxVoteds.Location = new System.Drawing.Point(324, 101);
            this.checkedListBoxVoteds.Name = "checkedListBoxVoteds";
            this.checkedListBoxVoteds.Size = new System.Drawing.Size(268, 388);
            this.checkedListBoxVoteds.TabIndex = 3;
            // 
            // labelMode
            // 
            this.labelMode.AutoSize = true;
            this.labelMode.Location = new System.Drawing.Point(465, 42);
            this.labelMode.Name = "labelMode";
            this.labelMode.Size = new System.Drawing.Size(41, 12);
            this.labelMode.TabIndex = 4;
            this.labelMode.Text = "模式：";
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(627, 523);
            this.Controls.Add(this.labelMode);
            this.Controls.Add(this.checkedListBoxVoteds);
            this.Controls.Add(this.checkedListBoxDept);
            this.Controls.Add(this.buttonCal);
            this.Controls.Add(this.buttonLoad);
            this.Name = "FormMain";
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
    }
}

