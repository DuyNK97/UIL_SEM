namespace CIM.Forms
{
    partial class FormSetting
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
            this.lbResultSetting = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.pathLocal = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.pathNAS = new System.Windows.Forms.TextBox();
            this.dayDeleteCSV = new System.Windows.Forms.TextBox();
            this.btnSaveConfigPath = new System.Windows.Forms.Button();
            this.cbAutoDeleteCSV = new System.Windows.Forms.CheckBox();
            this.btnSaveNAS = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbResultSetting
            // 
            this.lbResultSetting.AutoSize = true;
            this.lbResultSetting.Font = new System.Drawing.Font("Microsoft Tai Le", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbResultSetting.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lbResultSetting.Location = new System.Drawing.Point(17, 241);
            this.lbResultSetting.Name = "lbResultSetting";
            this.lbResultSetting.Size = new System.Drawing.Size(14, 21);
            this.lbResultSetting.TabIndex = 21;
            this.lbResultSetting.Text = " ";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.pathLocal);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.pathNAS);
            this.groupBox1.Controls.Add(this.dayDeleteCSV);
            this.groupBox1.Controls.Add(this.btnSaveConfigPath);
            this.groupBox1.Controls.Add(this.cbAutoDeleteCSV);
            this.groupBox1.Controls.Add(this.btnSaveNAS);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(21, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1194, 210);
            this.groupBox1.TabIndex = 20;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Custom CSV";
            // 
            // pathLocal
            // 
            this.pathLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pathLocal.Location = new System.Drawing.Point(421, 23);
            this.pathLocal.Multiline = true;
            this.pathLocal.Name = "pathLocal";
            this.pathLocal.Size = new System.Drawing.Size(560, 40);
            this.pathLocal.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(589, 160);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(48, 17);
            this.label3.TabIndex = 15;
            this.label3.Text = "(days)";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(247, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(161, 20);
            this.label1.TabIndex = 5;
            this.label1.Text = "Path Save CSV Local";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(247, 94);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(158, 20);
            this.label2.TabIndex = 7;
            this.label2.Text = "Path Save CSV MES";
            // 
            // pathNAS
            // 
            this.pathNAS.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pathNAS.Location = new System.Drawing.Point(421, 85);
            this.pathNAS.Multiline = true;
            this.pathNAS.Name = "pathNAS";
            this.pathNAS.Size = new System.Drawing.Size(560, 40);
            this.pathNAS.TabIndex = 9;
            // 
            // dayDeleteCSV
            // 
            this.dayDeleteCSV.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dayDeleteCSV.Location = new System.Drawing.Point(421, 148);
            this.dayDeleteCSV.Multiline = true;
            this.dayDeleteCSV.Name = "dayDeleteCSV";
            this.dayDeleteCSV.Size = new System.Drawing.Size(163, 40);
            this.dayDeleteCSV.TabIndex = 14;
            // 
            // btnSaveConfigPath
            // 
            this.btnSaveConfigPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveConfigPath.Location = new System.Drawing.Point(1005, 52);
            this.btnSaveConfigPath.Name = "btnSaveConfigPath";
            this.btnSaveConfigPath.Size = new System.Drawing.Size(162, 45);
            this.btnSaveConfigPath.TabIndex = 11;
            this.btnSaveConfigPath.Text = "Save";
            this.btnSaveConfigPath.UseVisualStyleBackColor = true;
            this.btnSaveConfigPath.Click += new System.EventHandler(this.btnSaveConfigPath_Click);
            // 
            // cbAutoDeleteCSV
            // 
            this.cbAutoDeleteCSV.AutoSize = true;
            this.cbAutoDeleteCSV.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbAutoDeleteCSV.Location = new System.Drawing.Point(216, 157);
            this.cbAutoDeleteCSV.Name = "cbAutoDeleteCSV";
            this.cbAutoDeleteCSV.Size = new System.Drawing.Size(192, 24);
            this.cbAutoDeleteCSV.TabIndex = 12;
            this.cbAutoDeleteCSV.Text = "Auto Delete CSV Local";
            this.cbAutoDeleteCSV.UseVisualStyleBackColor = true;
            this.cbAutoDeleteCSV.CheckedChanged += new System.EventHandler(this.cbAutoDeleteCSV_CheckedChanged);
            // 
            // btnSaveNAS
            // 
            this.btnSaveNAS.AutoSize = true;
            this.btnSaveNAS.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveNAS.Location = new System.Drawing.Point(29, 60);
            this.btnSaveNAS.Name = "btnSaveNAS";
            this.btnSaveNAS.Size = new System.Drawing.Size(186, 29);
            this.btnSaveNAS.TabIndex = 2;
            this.btnSaveNAS.Text = "Save Folder MES";
            this.btnSaveNAS.UseVisualStyleBackColor = true;
            this.btnSaveNAS.CheckedChanged += new System.EventHandler(this.btnSaveNAS_CheckedChanged);
            // 
            // FormSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1348, 429);
            this.Controls.Add(this.lbResultSetting);
            this.Controls.Add(this.groupBox1);
            this.Name = "FormSetting";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormSetting";
            this.Load += new System.EventHandler(this.FormSetting_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbResultSetting;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox pathLocal;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox pathNAS;
        private System.Windows.Forms.TextBox dayDeleteCSV;
        private System.Windows.Forms.Button btnSaveConfigPath;
        private System.Windows.Forms.CheckBox cbAutoDeleteCSV;
        private System.Windows.Forms.CheckBox btnSaveNAS;
    }
}