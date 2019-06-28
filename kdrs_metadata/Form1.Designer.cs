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
            this.label2 = new System.Windows.Forms.Label();
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
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 34);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "             ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 48);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "           ";
            // 
            // priorityHigh
            // 
            this.priorityHigh.AutoSize = true;
            this.priorityHigh.Checked = true;
            this.priorityHigh.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityHigh.Location = new System.Drawing.Point(8, 19);
            this.priorityHigh.Margin = new System.Windows.Forms.Padding(2);
            this.priorityHigh.Name = "priorityHigh";
            this.priorityHigh.Size = new System.Drawing.Size(48, 17);
            this.priorityHigh.TabIndex = 2;
            this.priorityHigh.Text = "High";
            this.priorityHigh.UseVisualStyleBackColor = true;
            // 
            // priorityMedium
            // 
            this.priorityMedium.AutoSize = true;
            this.priorityMedium.Checked = true;
            this.priorityMedium.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityMedium.Location = new System.Drawing.Point(8, 41);
            this.priorityMedium.Margin = new System.Windows.Forms.Padding(2);
            this.priorityMedium.Name = "priorityMedium";
            this.priorityMedium.Size = new System.Drawing.Size(63, 17);
            this.priorityMedium.TabIndex = 3;
            this.priorityMedium.Text = "Medium";
            this.priorityMedium.UseVisualStyleBackColor = true;
            // 
            // prioritySystem
            // 
            this.prioritySystem.AutoSize = true;
            this.prioritySystem.Checked = true;
            this.prioritySystem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.prioritySystem.Location = new System.Drawing.Point(100, 19);
            this.prioritySystem.Margin = new System.Windows.Forms.Padding(2);
            this.prioritySystem.Name = "prioritySystem";
            this.prioritySystem.Size = new System.Drawing.Size(60, 17);
            this.prioritySystem.TabIndex = 4;
            this.prioritySystem.Text = "System";
            this.prioritySystem.UseVisualStyleBackColor = true;
            // 
            // priorityEmpty
            // 
            this.priorityEmpty.AutoSize = true;
            this.priorityEmpty.Checked = true;
            this.priorityEmpty.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityEmpty.Location = new System.Drawing.Point(100, 41);
            this.priorityEmpty.Margin = new System.Windows.Forms.Padding(2);
            this.priorityEmpty.Name = "priorityEmpty";
            this.priorityEmpty.Size = new System.Drawing.Size(55, 17);
            this.priorityEmpty.TabIndex = 5;
            this.priorityEmpty.Text = "Empty";
            this.priorityEmpty.UseVisualStyleBackColor = true;
            // 
            // priorityLow
            // 
            this.priorityLow.AutoSize = true;
            this.priorityLow.Checked = true;
            this.priorityLow.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityLow.Location = new System.Drawing.Point(8, 63);
            this.priorityLow.Margin = new System.Windows.Forms.Padding(2);
            this.priorityLow.Name = "priorityLow";
            this.priorityLow.Size = new System.Drawing.Size(46, 17);
            this.priorityLow.TabIndex = 6;
            this.priorityLow.Text = "Low";
            this.priorityLow.UseVisualStyleBackColor = true;
            // 
            // priorityNull
            // 
            this.priorityNull.AutoSize = true;
            this.priorityNull.Checked = true;
            this.priorityNull.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityNull.Location = new System.Drawing.Point(192, 17);
            this.priorityNull.Margin = new System.Windows.Forms.Padding(2);
            this.priorityNull.Name = "priorityNull";
            this.priorityNull.Size = new System.Drawing.Size(42, 17);
            this.priorityNull.TabIndex = 7;
            this.priorityNull.Text = "null";
            this.priorityNull.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.priorityDummy);
            this.groupBox1.Controls.Add(this.priorityStat);
            this.groupBox1.Controls.Add(this.priorityNull);
            this.groupBox1.Controls.Add(this.priorityLow);
            this.groupBox1.Controls.Add(this.priorityEmpty);
            this.groupBox1.Controls.Add(this.prioritySystem);
            this.groupBox1.Controls.Add(this.priorityMedium);
            this.groupBox1.Controls.Add(this.priorityHigh);
            this.groupBox1.Location = new System.Drawing.Point(11, 187);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(286, 93);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Table priorities";
            // 
            // priorityDummy
            // 
            this.priorityDummy.AutoSize = true;
            this.priorityDummy.Checked = true;
            this.priorityDummy.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityDummy.Location = new System.Drawing.Point(192, 41);
            this.priorityDummy.Margin = new System.Windows.Forms.Padding(2);
            this.priorityDummy.Name = "priorityDummy";
            this.priorityDummy.Size = new System.Drawing.Size(61, 17);
            this.priorityDummy.TabIndex = 9;
            this.priorityDummy.Text = "Dummy";
            this.priorityDummy.UseVisualStyleBackColor = true;
            // 
            // priorityStat
            // 
            this.priorityStat.AutoSize = true;
            this.priorityStat.Checked = true;
            this.priorityStat.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityStat.Location = new System.Drawing.Point(100, 63);
            this.priorityStat.Margin = new System.Windows.Forms.Padding(2);
            this.priorityStat.Name = "priorityStat";
            this.priorityStat.Size = new System.Drawing.Size(45, 17);
            this.priorityStat.TabIndex = 8;
            this.priorityStat.Text = "Stat";
            this.priorityStat.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.includeTables);
            this.groupBox2.Location = new System.Drawing.Point(303, 189);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(184, 91);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Tables";
            // 
            // includeTables
            // 
            this.includeTables.AutoSize = true;
            this.includeTables.Checked = true;
            this.includeTables.CheckState = System.Windows.Forms.CheckState.Checked;
            this.includeTables.Location = new System.Drawing.Point(7, 15);
            this.includeTables.Name = "includeTables";
            this.includeTables.Size = new System.Drawing.Size(129, 17);
            this.includeTables.TabIndex = 0;
            this.includeTables.Text = "Include table columns";
            this.includeTables.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(11, 64);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(572, 69);
            this.textBox1.TabIndex = 11;
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(596, 341);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "KDRS Metadata";
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
        private System.Windows.Forms.Label label2;
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
    }
}

