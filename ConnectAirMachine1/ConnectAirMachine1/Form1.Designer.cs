namespace ConnectAirMachine1
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
            this.txtClient = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnGetResultAirMachine = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtClient
            // 
            this.txtClient.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtClient.Location = new System.Drawing.Point(51, 40);
            this.txtClient.Margin = new System.Windows.Forms.Padding(4);
            this.txtClient.Multiline = true;
            this.txtClient.Name = "txtClient";
            this.txtClient.Size = new System.Drawing.Size(363, 51);
            this.txtClient.TabIndex = 5;
            this.txtClient.Text = "5";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(48, 20);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 16);
            this.label1.TabIndex = 6;
            this.label1.Text = "Client";
            // 
            // btnGetResultAirMachine
            // 
            this.btnGetResultAirMachine.Location = new System.Drawing.Point(461, 40);
            this.btnGetResultAirMachine.Name = "btnGetResultAirMachine";
            this.btnGetResultAirMachine.Size = new System.Drawing.Size(170, 51);
            this.btnGetResultAirMachine.TabIndex = 11;
            this.btnGetResultAirMachine.Text = "Get Result Air Machine";
            this.btnGetResultAirMachine.UseVisualStyleBackColor = true;
            this.btnGetResultAirMachine.Click += new System.EventHandler(this.btnGetResultAirMachine_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(719, 146);
            this.Controls.Add(this.btnGetResultAirMachine);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtClient);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtClient;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnGetResultAirMachine;
    }
}

