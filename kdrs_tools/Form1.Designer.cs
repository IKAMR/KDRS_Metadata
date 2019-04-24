namespace Metadata_XLS
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
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "             ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 17);
            this.label2.TabIndex = 1;
            this.label2.Text = "           ";
            // 
            // priorityHigh
            // 
            this.priorityHigh.AutoSize = true;
            this.priorityHigh.Checked = true;
            this.priorityHigh.CheckState = System.Windows.Forms.CheckState.Checked;
            this.priorityHigh.Location = new System.Drawing.Point(11, 23);
            this.priorityHigh.Name = "priorityHigh";
            this.priorityHigh.Size = new System.Drawing.Size(59, 21);
            this.priorityHigh.TabIndex = 2;
            this.priorityHigh.Text = "High";
            this.priorityHigh.UseVisualStyleBackColor = true;
            // 
            // priorityMedium
            // 
            this.priorityMedium.AutoSize = true;
            this.priorityMedium.Location = new System.Drawing.Point(11, 50);
            this.priorityMedium.Name = "priorityMedium";
            this.priorityMedium.Size = new System.Drawing.Size(79, 21);
            this.priorityMedium.TabIndex = 3;
            this.priorityMedium.Text = "Medium";
            this.priorityMedium.UseVisualStyleBackColor = true;
            // 
            // prioritySystem
            // 
            this.prioritySystem.AutoSize = true;
            this.prioritySystem.Location = new System.Drawing.Point(134, 23);
            this.prioritySystem.Name = "prioritySystem";
            this.prioritySystem.Size = new System.Drawing.Size(76, 21);
            this.prioritySystem.TabIndex = 4;
            this.prioritySystem.Text = "System";
            this.prioritySystem.UseVisualStyleBackColor = true;
            // 
            // priorityEmpty
            // 
            this.priorityEmpty.AutoSize = true;
            this.priorityEmpty.Location = new System.Drawing.Point(134, 50);
            this.priorityEmpty.Name = "priorityEmpty";
            this.priorityEmpty.Size = new System.Drawing.Size(69, 21);
            this.priorityEmpty.TabIndex = 5;
            this.priorityEmpty.Text = "Empty";
            this.priorityEmpty.UseVisualStyleBackColor = true;
            // 
            // priorityLow
            // 
            this.priorityLow.AutoSize = true;
            this.priorityLow.Location = new System.Drawing.Point(11, 77);
            this.priorityLow.Name = "priorityLow";
            this.priorityLow.Size = new System.Drawing.Size(55, 21);
            this.priorityLow.TabIndex = 6;
            this.priorityLow.Text = "Low";
            this.priorityLow.UseVisualStyleBackColor = true;
            // 
            // priorityNull
            // 
            this.priorityNull.AutoSize = true;
            this.priorityNull.Location = new System.Drawing.Point(134, 77);
            this.priorityNull.Name = "priorityNull";
            this.priorityNull.Size = new System.Drawing.Size(52, 21);
            this.priorityNull.TabIndex = 7;
            this.priorityNull.Text = "null";
            this.priorityNull.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.priorityNull);
            this.groupBox1.Controls.Add(this.priorityLow);
            this.groupBox1.Controls.Add(this.priorityEmpty);
            this.groupBox1.Controls.Add(this.prioritySystem);
            this.groupBox1.Controls.Add(this.priorityMedium);
            this.groupBox1.Controls.Add(this.priorityHigh);
            this.groupBox1.Location = new System.Drawing.Point(5, 118);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(260, 114);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Table priorities";
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(594, 244);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "IKAVA Metadata2XLS2Metadata converter";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form1_DragEnter);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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
    }
}

