namespace KDRS_Metadata
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.priorityHigh = new System.Windows.Forms.CheckBox();
            this.priorityMedium = new System.Windows.Forms.CheckBox();
            this.prioritySystem = new System.Windows.Forms.CheckBox();
            this.priorityEmpty = new System.Windows.Forms.CheckBox();
            this.priorityLow = new System.Windows.Forms.CheckBox();
            this.priorityNull = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.priorityDummy = new System.Windows.Forms.CheckBox();
            this.priorityStat = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.includeTables = new System.Windows.Forms.CheckBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.btnCopyLog = new System.Windows.Forms.Button();
            this.btnSaveLog = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 63);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "             ";
            // 
            // priorityHigh
            // 
            this.priorityHigh.AutoSize = true;
            this.priorityHigh.Checked = true;
            this.priorityHigh.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityHigh.Location = new System.Drawing.Point(15, 35);
            this.priorityHigh.Margin = new System.Windows.Forms.Padding(4);
            this.priorityHigh.Name = "priorityHigh";
            this.priorityHigh.Size = new System.Drawing.Size(78, 29);
            this.priorityHigh.TabIndex = 2;
            this.priorityHigh.Text = "High";
            this.priorityHigh.UseVisualStyleBackColor = true;
            // 
            // priorityMedium
            // 
            this.priorityMedium.AutoSize = true;
            this.priorityMedium.Checked = true;
            this.priorityMedium.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityMedium.Location = new System.Drawing.Point(15, 76);
            this.priorityMedium.Margin = new System.Windows.Forms.Padding(4);
            this.priorityMedium.Name = "priorityMedium";
            this.priorityMedium.Size = new System.Drawing.Size(108, 29);
            this.priorityMedium.TabIndex = 3;
            this.priorityMedium.Text = "Medium";
            this.priorityMedium.UseVisualStyleBackColor = true;
            // 
            // prioritySystem
            // 
            this.prioritySystem.AutoSize = true;
            this.prioritySystem.Checked = true;
            this.prioritySystem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.prioritySystem.Location = new System.Drawing.Point(183, 35);
            this.prioritySystem.Margin = new System.Windows.Forms.Padding(4);
            this.prioritySystem.Name = "prioritySystem";
            this.prioritySystem.Size = new System.Drawing.Size(104, 29);
            this.prioritySystem.TabIndex = 4;
            this.prioritySystem.Text = "System";
            this.prioritySystem.UseVisualStyleBackColor = true;
            // 
            // priorityEmpty
            // 
            this.priorityEmpty.AutoSize = true;
            this.priorityEmpty.Checked = true;
            this.priorityEmpty.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityEmpty.Location = new System.Drawing.Point(183, 76);
            this.priorityEmpty.Margin = new System.Windows.Forms.Padding(4);
            this.priorityEmpty.Name = "priorityEmpty";
            this.priorityEmpty.Size = new System.Drawing.Size(93, 29);
            this.priorityEmpty.TabIndex = 5;
            this.priorityEmpty.Text = "Empty";
            this.priorityEmpty.UseVisualStyleBackColor = true;
            // 
            // priorityLow
            // 
            this.priorityLow.AutoSize = true;
            this.priorityLow.Checked = true;
            this.priorityLow.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityLow.Location = new System.Drawing.Point(15, 116);
            this.priorityLow.Margin = new System.Windows.Forms.Padding(4);
            this.priorityLow.Name = "priorityLow";
            this.priorityLow.Size = new System.Drawing.Size(74, 29);
            this.priorityLow.TabIndex = 6;
            this.priorityLow.Text = "Low";
            this.priorityLow.UseVisualStyleBackColor = true;
            // 
            // priorityNull
            // 
            this.priorityNull.AutoSize = true;
            this.priorityNull.Checked = true;
            this.priorityNull.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityNull.Location = new System.Drawing.Point(352, 31);
            this.priorityNull.Margin = new System.Windows.Forms.Padding(4);
            this.priorityNull.Name = "priorityNull";
            this.priorityNull.Size = new System.Drawing.Size(68, 29);
            this.priorityNull.TabIndex = 7;
            this.priorityNull.Text = "null";
            this.priorityNull.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.groupBox1.Controls.Add(this.priorityDummy);
            this.groupBox1.Controls.Add(this.priorityStat);
            this.groupBox1.Controls.Add(this.priorityNull);
            this.groupBox1.Controls.Add(this.priorityLow);
            this.groupBox1.Controls.Add(this.priorityEmpty);
            this.groupBox1.Controls.Add(this.prioritySystem);
            this.groupBox1.Controls.Add(this.priorityMedium);
            this.groupBox1.Controls.Add(this.priorityHigh);
            this.groupBox1.Location = new System.Drawing.Point(20, 451);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(524, 172);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Table priorities";
            // 
            // priorityDummy
            // 
            this.priorityDummy.AutoSize = true;
            this.priorityDummy.Checked = true;
            this.priorityDummy.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityDummy.Location = new System.Drawing.Point(352, 76);
            this.priorityDummy.Margin = new System.Windows.Forms.Padding(4);
            this.priorityDummy.Name = "priorityDummy";
            this.priorityDummy.Size = new System.Drawing.Size(105, 29);
            this.priorityDummy.TabIndex = 9;
            this.priorityDummy.Text = "Dummy";
            this.priorityDummy.UseVisualStyleBackColor = true;
            // 
            // priorityStat
            // 
            this.priorityStat.AutoSize = true;
            this.priorityStat.Checked = true;
            this.priorityStat.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityStat.Location = new System.Drawing.Point(183, 116);
            this.priorityStat.Margin = new System.Windows.Forms.Padding(4);
            this.priorityStat.Name = "priorityStat";
            this.priorityStat.Size = new System.Drawing.Size(73, 29);
            this.priorityStat.TabIndex = 8;
            this.priorityStat.Text = "Stat";
            this.priorityStat.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.groupBox2.Controls.Add(this.includeTables);
            this.groupBox2.Location = new System.Drawing.Point(556, 455);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox2.Size = new System.Drawing.Size(337, 168);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Tables";
            // 
            // includeTables
            // 
            this.includeTables.AutoSize = true;
            this.includeTables.Checked = true;
            this.includeTables.CheckState = System.Windows.Forms.CheckState.Checked;
            this.includeTables.Location = new System.Drawing.Point(13, 28);
            this.includeTables.Margin = new System.Windows.Forms.Padding(6);
            this.includeTables.Name = "includeTables";
            this.includeTables.Size = new System.Drawing.Size(226, 29);
            this.includeTables.TabIndex = 0;
            this.includeTables.Text = "Include table columns";
            this.includeTables.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(20, 92);
            this.textBox1.Margin = new System.Windows.Forms.Padding(6);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(1532, 289);
            this.textBox1.TabIndex = 11;
            // 
            // btnCopyLog
            // 
            this.btnCopyLog.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCopyLog.Location = new System.Drawing.Point(24, 634);
            this.btnCopyLog.Margin = new System.Windows.Forms.Padding(6);
            this.btnCopyLog.Name = "btnCopyLog";
            this.btnCopyLog.Size = new System.Drawing.Size(160, 50);
            this.btnCopyLog.TabIndex = 12;
            this.btnCopyLog.Text = "Copy log";
            this.btnCopyLog.UseVisualStyleBackColor = true;
            this.btnCopyLog.Click += new System.EventHandler(this.btnCopyLog_Click);
            // 
            // btnSaveLog
            // 
            this.btnSaveLog.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSaveLog.Location = new System.Drawing.Point(194, 634);
            this.btnSaveLog.Margin = new System.Windows.Forms.Padding(6);
            this.btnSaveLog.Name = "btnSaveLog";
            this.btnSaveLog.Size = new System.Drawing.Size(160, 50);
            this.btnSaveLog.TabIndex = 13;
            this.btnSaveLog.Text = "Save log";
            this.btnSaveLog.UseVisualStyleBackColor = true;
            this.btnSaveLog.Click += new System.EventHandler(this.btnSaveLog_Click);
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1576, 736);
            this.Controls.Add(this.btnSaveLog);
            this.Controls.Add(this.btnCopyLog);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "Form1";
            this.Text = "KDRS Metadata vn.n";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form1_DragEnter);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;

        private System.Windows.Forms.CheckBox priorityHigh;
        private System.Windows.Forms.CheckBox priorityMedium;
        private System.Windows.Forms.CheckBox prioritySystem;
        private System.Windows.Forms.CheckBox priorityEmpty;
        private System.Windows.Forms.CheckBox priorityLow;
        private System.Windows.Forms.CheckBox priorityNull;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox priorityDummy;
        private System.Windows.Forms.CheckBox priorityStat;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox includeTables;
        private System.Windows.Forms.TextBox textBox1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button btnCopyLog;
        private System.Windows.Forms.Button btnSaveLog;
    }
}

