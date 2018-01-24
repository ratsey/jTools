namespace JTools
{
    partial class frmSetup
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSetup));
            this.txtMLOExe = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMLOGUID = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtMLODataFile = new System.Windows.Forms.TextBox();
            this.chkDebug = new System.Windows.Forms.CheckBox();
            this.txtLogFile = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnFindMLO = new System.Windows.Forms.Button();
            this.btnFindMLOData = new System.Windows.Forms.Button();
            this.btnFindLog = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtMLOExe
            // 
            this.txtMLOExe.Location = new System.Drawing.Point(162, 51);
            this.txtMLOExe.Name = "txtMLOExe";
            this.txtMLOExe.Size = new System.Drawing.Size(233, 20);
            this.txtMLOExe.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Location of MLO.EXE";
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(353, 278);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 5;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 130);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(139, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "GUID of MLO Root Task ID";
            // 
            // txtMLOGUID
            // 
            this.txtMLOGUID.Location = new System.Drawing.Point(162, 127);
            this.txtMLOGUID.Name = "txtMLOGUID";
            this.txtMLOGUID.Size = new System.Drawing.Size(233, 20);
            this.txtMLOGUID.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 89);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "MLO Data File";
            // 
            // txtMLODataFile
            // 
            this.txtMLODataFile.Location = new System.Drawing.Point(162, 89);
            this.txtMLODataFile.Name = "txtMLODataFile";
            this.txtMLODataFile.Size = new System.Drawing.Size(233, 20);
            this.txtMLODataFile.TabIndex = 12;
            // 
            // chkDebug
            // 
            this.chkDebug.AutoSize = true;
            this.chkDebug.Location = new System.Drawing.Point(18, 189);
            this.chkDebug.Name = "chkDebug";
            this.chkDebug.Size = new System.Drawing.Size(88, 17);
            this.chkDebug.TabIndex = 14;
            this.chkDebug.Text = "Debug Mode";
            this.chkDebug.UseVisualStyleBackColor = true;
            // 
            // txtLogFile
            // 
            this.txtLogFile.Location = new System.Drawing.Point(162, 219);
            this.txtLogFile.Name = "txtLogFile";
            this.txtLogFile.Size = new System.Drawing.Size(233, 20);
            this.txtLogFile.TabIndex = 16;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 219);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Log File";
            // 
            // btnFindMLO
            // 
            this.btnFindMLO.Location = new System.Drawing.Point(401, 51);
            this.btnFindMLO.Name = "btnFindMLO";
            this.btnFindMLO.Size = new System.Drawing.Size(27, 23);
            this.btnFindMLO.TabIndex = 17;
            this.btnFindMLO.Text = "...";
            this.btnFindMLO.UseVisualStyleBackColor = true;
            this.btnFindMLO.Click += new System.EventHandler(this.btnFindMLO_Click);
            // 
            // btnFindMLOData
            // 
            this.btnFindMLOData.Location = new System.Drawing.Point(401, 87);
            this.btnFindMLOData.Name = "btnFindMLOData";
            this.btnFindMLOData.Size = new System.Drawing.Size(27, 23);
            this.btnFindMLOData.TabIndex = 18;
            this.btnFindMLOData.Text = "...";
            this.btnFindMLOData.UseVisualStyleBackColor = true;
            this.btnFindMLOData.Click += new System.EventHandler(this.btnMLODataFile_Click);
            // 
            // btnFindLog
            // 
            this.btnFindLog.Location = new System.Drawing.Point(401, 217);
            this.btnFindLog.Name = "btnFindLog";
            this.btnFindLog.Size = new System.Drawing.Size(27, 23);
            this.btnFindLog.TabIndex = 19;
            this.btnFindLog.Text = "...";
            this.btnFindLog.UseVisualStyleBackColor = true;
            this.btnFindLog.Click += new System.EventHandler(this.btnLogFile_Click);
            // 
            // frmSetup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(440, 313);
            this.Controls.Add(this.btnFindLog);
            this.Controls.Add(this.btnFindMLOData);
            this.Controls.Add(this.btnFindMLO);
            this.Controls.Add(this.txtLogFile);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.chkDebug);
            this.Controls.Add(this.txtMLODataFile);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtMLOGUID);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtMLOExe);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnOK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmSetup";
            this.Text = "OneNote MLO Setup";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtMLOExe;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtMLOGUID;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtMLODataFile;
        private System.Windows.Forms.CheckBox chkDebug;
        private System.Windows.Forms.TextBox txtLogFile;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnFindMLO;
        private System.Windows.Forms.Button btnFindMLOData;
        private System.Windows.Forms.Button btnFindLog;
    }
}