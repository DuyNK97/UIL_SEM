namespace CIM
{
    partial class SearchForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnSearch = new System.Windows.Forms.Button();
            this.txtQR = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.QR = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box1_glue_amount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fbondccode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.foutput = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box1_glue_dischargevolume_vision = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.insulator_bar_code = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sbondcode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.souput = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box1_glueoverflow_vison = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box1_heated_air_curing = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box2_glue_amount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box2_glue_dischargevolume_vision = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FPCBbar_code = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box2_glueoverflow_vision = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box2_heated_air_curing = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box2_heigh = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tbondcode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.toutput = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box3_glue_amount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box3_glue_dischargevolume_vision = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box3_heated_air_curing = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box3_glueoverflow_vision = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tighness_and_location_vision = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.height_parallelism = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.resistance = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.air_leakagetest_detail = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.air_leakagetest_result = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.testtime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Remark = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.btnSearch);
            this.panel1.Controls.Add(this.txtQR);
            this.panel1.Location = new System.Drawing.Point(1, 35);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1374, 86);
            this.panel1.TabIndex = 4;
            // 
            // btnSearch
            // 
            this.btnSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSearch.Location = new System.Drawing.Point(1125, 20);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(233, 48);
            this.btnSearch.TabIndex = 3;
            this.btnSearch.Text = "Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // txtQR
            // 
            this.txtQR.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtQR.Location = new System.Drawing.Point(782, 31);
            this.txtQR.Name = "txtQR";
            this.txtQR.Size = new System.Drawing.Size(328, 28);
            this.txtQR.TabIndex = 2;
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Location = new System.Drawing.Point(3, 130);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1370, 505);
            this.panel2.TabIndex = 5;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Tai Le", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.QR,
            this.box1_glue_amount,
            this.fbondccode,
            this.foutput,
            this.box1_glue_dischargevolume_vision,
            this.insulator_bar_code,
            this.sbondcode,
            this.souput,
            this.box1_glueoverflow_vison,
            this.box1_heated_air_curing,
            this.box2_glue_amount,
            this.box2_glue_dischargevolume_vision,
            this.FPCBbar_code,
            this.box2_glueoverflow_vision,
            this.box2_heated_air_curing,
            this.box2_heigh,
            this.tbondcode,
            this.toutput,
            this.box3_glue_amount,
            this.box3_glue_dischargevolume_vision,
            this.box3_heated_air_curing,
            this.box3_glueoverflow_vision,
            this.tighness_and_location_vision,
            this.height_parallelism,
            this.resistance,
            this.air_leakagetest_detail,
            this.air_leakagetest_result,
            this.testtime,
            this.Remark});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridView1.Name = "dataGridView1";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Tai Le", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1370, 505);
            this.dataGridView1.TabIndex = 4;
            this.dataGridView1.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView1_CellFormatting);
            // 
            // QR
            // 
            this.QR.HeaderText = "Top Housing\nQR";
            this.QR.MinimumWidth = 6;
            this.QR.Name = "QR";
            this.QR.Width = 300;
            // 
            // box1_glue_amount
            // 
            this.box1_glue_amount.HeaderText = "1st Glue Amount";
            this.box1_glue_amount.MinimumWidth = 6;
            this.box1_glue_amount.Name = "box1_glue_amount";
            this.box1_glue_amount.Width = 125;
            // 
            // fbondccode
            // 
            this.fbondccode.HeaderText = "1st Bond Code";
            this.fbondccode.Name = "fbondccode";
            this.fbondccode.ReadOnly = true;
            // 
            // foutput
            // 
            this.foutput.HeaderText = "1st F-Output/E-Input";
            this.foutput.Name = "foutput";
            this.foutput.ReadOnly = true;
            // 
            // box1_glue_dischargevolume_vision
            // 
            this.box1_glue_dischargevolume_vision.HeaderText = "1st  Glue discharge\nvolume Vision";
            this.box1_glue_dischargevolume_vision.MinimumWidth = 6;
            this.box1_glue_dischargevolume_vision.Name = "box1_glue_dischargevolume_vision";
            this.box1_glue_dischargevolume_vision.Width = 125;
            // 
            // insulator_bar_code
            // 
            this.insulator_bar_code.HeaderText = "Insulator\nbar code";
            this.insulator_bar_code.MinimumWidth = 6;
            this.insulator_bar_code.Name = "insulator_bar_code";
            this.insulator_bar_code.Width = 125;
            // 
            // sbondcode
            // 
            this.sbondcode.HeaderText = "2nd Bond Code";
            this.sbondcode.Name = "sbondcode";
            this.sbondcode.ReadOnly = true;
            // 
            // souput
            // 
            this.souput.HeaderText = "2nd F-Output/E-Input";
            this.souput.Name = "souput";
            this.souput.ReadOnly = true;
            // 
            // box1_glueoverflow_vison
            // 
            this.box1_glueoverflow_vison.HeaderText = "1st Glue\noverflow vision";
            this.box1_glueoverflow_vison.MinimumWidth = 6;
            this.box1_glueoverflow_vison.Name = "box1_glueoverflow_vison";
            this.box1_glueoverflow_vison.Width = 125;
            // 
            // box1_heated_air_curing
            // 
            this.box1_heated_air_curing.HeaderText = "1st heated Air curing";
            this.box1_heated_air_curing.MinimumWidth = 6;
            this.box1_heated_air_curing.Name = "box1_heated_air_curing";
            this.box1_heated_air_curing.Width = 150;
            // 
            // box2_glue_amount
            // 
            this.box2_glue_amount.HeaderText = "2nd Glue Amount";
            this.box2_glue_amount.MinimumWidth = 6;
            this.box2_glue_amount.Name = "box2_glue_amount";
            this.box2_glue_amount.Width = 125;
            // 
            // box2_glue_dischargevolume_vision
            // 
            this.box2_glue_dischargevolume_vision.HeaderText = "2nd Glue discharge\nvolume Vision";
            this.box2_glue_dischargevolume_vision.MinimumWidth = 6;
            this.box2_glue_dischargevolume_vision.Name = "box2_glue_dischargevolume_vision";
            this.box2_glue_dischargevolume_vision.Width = 125;
            // 
            // FPCBbar_code
            // 
            this.FPCBbar_code.HeaderText = "FPCB\nbar code";
            this.FPCBbar_code.MinimumWidth = 6;
            this.FPCBbar_code.Name = "FPCBbar_code";
            this.FPCBbar_code.Width = 125;
            // 
            // box2_glueoverflow_vision
            // 
            this.box2_glueoverflow_vision.HeaderText = "2nd Glue\noverflow vision";
            this.box2_glueoverflow_vision.MinimumWidth = 6;
            this.box2_glueoverflow_vision.Name = "box2_glueoverflow_vision";
            this.box2_glueoverflow_vision.Width = 125;
            // 
            // box2_heated_air_curing
            // 
            this.box2_heated_air_curing.HeaderText = "2nd heated Air curing";
            this.box2_heated_air_curing.MinimumWidth = 6;
            this.box2_heated_air_curing.Name = "box2_heated_air_curing";
            this.box2_heated_air_curing.Width = 150;
            // 
            // box2_heigh
            // 
            this.box2_heigh.HeaderText = "Box 2 height";
            this.box2_heigh.MinimumWidth = 6;
            this.box2_heigh.Name = "box2_heigh";
            this.box2_heigh.Width = 125;
            // 
            // tbondcode
            // 
            this.tbondcode.HeaderText = "3rd Bond Code";
            this.tbondcode.Name = "tbondcode";
            this.tbondcode.ReadOnly = true;
            // 
            // toutput
            // 
            this.toutput.HeaderText = "3rd F-Output/E-Input";
            this.toutput.Name = "toutput";
            this.toutput.ReadOnly = true;
            // 
            // box3_glue_amount
            // 
            this.box3_glue_amount.HeaderText = "3rd Glue Amount";
            this.box3_glue_amount.MinimumWidth = 6;
            this.box3_glue_amount.Name = "box3_glue_amount";
            this.box3_glue_amount.Width = 125;
            // 
            // box3_glue_dischargevolume_vision
            // 
            this.box3_glue_dischargevolume_vision.HeaderText = "3rd Glue discharge\nvolume Vision";
            this.box3_glue_dischargevolume_vision.MinimumWidth = 6;
            this.box3_glue_dischargevolume_vision.Name = "box3_glue_dischargevolume_vision";
            this.box3_glue_dischargevolume_vision.Width = 125;
            // 
            // box3_heated_air_curing
            // 
            this.box3_heated_air_curing.HeaderText = "3rd heated Air curing";
            this.box3_heated_air_curing.MinimumWidth = 6;
            this.box3_heated_air_curing.Name = "box3_heated_air_curing";
            this.box3_heated_air_curing.Width = 150;
            // 
            // box3_glueoverflow_vision
            // 
            this.box3_glueoverflow_vision.HeaderText = "3rd Glue\noverflow vision";
            this.box3_glueoverflow_vision.MinimumWidth = 6;
            this.box3_glueoverflow_vision.Name = "box3_glueoverflow_vision";
            this.box3_glueoverflow_vision.Width = 125;
            // 
            // tighness_and_location_vision
            // 
            this.tighness_and_location_vision.HeaderText = "Tightness and location vision";
            this.tighness_and_location_vision.MinimumWidth = 6;
            this.tighness_and_location_vision.Name = "tighness_and_location_vision";
            this.tighness_and_location_vision.Width = 125;
            // 
            // height_parallelism
            // 
            this.height_parallelism.HeaderText = "Height /\n Parallelism";
            this.height_parallelism.MinimumWidth = 6;
            this.height_parallelism.Name = "height_parallelism";
            this.height_parallelism.Width = 125;
            // 
            // resistance
            // 
            this.resistance.HeaderText = "Resistance";
            this.resistance.MinimumWidth = 6;
            this.resistance.Name = "resistance";
            this.resistance.Width = 125;
            // 
            // air_leakagetest_detail
            // 
            this.air_leakagetest_detail.HeaderText = "Air Leakage\nTest Detail";
            this.air_leakagetest_detail.MinimumWidth = 6;
            this.air_leakagetest_detail.Name = "air_leakagetest_detail";
            this.air_leakagetest_detail.Width = 125;
            // 
            // air_leakagetest_result
            // 
            this.air_leakagetest_result.HeaderText = "Air Leakage\nTest Result";
            this.air_leakagetest_result.MinimumWidth = 6;
            this.air_leakagetest_result.Name = "air_leakagetest_result";
            this.air_leakagetest_result.Width = 125;
            // 
            // testtime
            // 
            this.testtime.HeaderText = "Test time";
            this.testtime.MinimumWidth = 6;
            this.testtime.Name = "testtime";
            this.testtime.Width = 250;
            // 
            // Remark
            // 
            this.Remark.HeaderText = "Remark";
            this.Remark.MinimumWidth = 6;
            this.Remark.Name = "Remark";
            this.Remark.Width = 125;
            // 
            // SearchForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(1376, 638);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Microsoft Tai Le", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "SearchForm";
            this.Text = "SearchForm";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.ZoomScaleRect = new System.Drawing.Rectangle(19, 19, 1376, 450);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.TextBox txtQR;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn QR;
        private System.Windows.Forms.DataGridViewTextBoxColumn box1_glue_amount;
        private System.Windows.Forms.DataGridViewTextBoxColumn fbondccode;
        private System.Windows.Forms.DataGridViewTextBoxColumn foutput;
        private System.Windows.Forms.DataGridViewTextBoxColumn box1_glue_dischargevolume_vision;
        private System.Windows.Forms.DataGridViewTextBoxColumn insulator_bar_code;
        private System.Windows.Forms.DataGridViewTextBoxColumn sbondcode;
        private System.Windows.Forms.DataGridViewTextBoxColumn souput;
        private System.Windows.Forms.DataGridViewTextBoxColumn box1_glueoverflow_vison;
        private System.Windows.Forms.DataGridViewTextBoxColumn box1_heated_air_curing;
        private System.Windows.Forms.DataGridViewTextBoxColumn box2_glue_amount;
        private System.Windows.Forms.DataGridViewTextBoxColumn box2_glue_dischargevolume_vision;
        private System.Windows.Forms.DataGridViewTextBoxColumn FPCBbar_code;
        private System.Windows.Forms.DataGridViewTextBoxColumn box2_glueoverflow_vision;
        private System.Windows.Forms.DataGridViewTextBoxColumn box2_heated_air_curing;
        private System.Windows.Forms.DataGridViewTextBoxColumn box2_heigh;
        private System.Windows.Forms.DataGridViewTextBoxColumn tbondcode;
        private System.Windows.Forms.DataGridViewTextBoxColumn toutput;
        private System.Windows.Forms.DataGridViewTextBoxColumn box3_glue_amount;
        private System.Windows.Forms.DataGridViewTextBoxColumn box3_glue_dischargevolume_vision;
        private System.Windows.Forms.DataGridViewTextBoxColumn box3_heated_air_curing;
        private System.Windows.Forms.DataGridViewTextBoxColumn box3_glueoverflow_vision;
        private System.Windows.Forms.DataGridViewTextBoxColumn tighness_and_location_vision;
        private System.Windows.Forms.DataGridViewTextBoxColumn height_parallelism;
        private System.Windows.Forms.DataGridViewTextBoxColumn resistance;
        private System.Windows.Forms.DataGridViewTextBoxColumn air_leakagetest_detail;
        private System.Windows.Forms.DataGridViewTextBoxColumn air_leakagetest_result;
        private System.Windows.Forms.DataGridViewTextBoxColumn testtime;
        private System.Windows.Forms.DataGridViewTextBoxColumn Remark;
    }
}