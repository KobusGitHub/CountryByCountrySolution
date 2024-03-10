namespace CBC_V2
{
    partial class Form2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.btnGenerate = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtSource = new System.Windows.Forms.TextBox();
            this.brnBrowseSource = new System.Windows.Forms.Button();
            this.btnBrowseDest = new System.Windows.Forms.Button();
            this.txtDestFolder = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.txtDestFileName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.lstLog = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rad8 = new System.Windows.Forms.RadioButton();
            this.rad16 = new System.Windows.Forms.RadioButton();
            this.ckSars = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGenerate
            // 
            this.btnGenerate.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnGenerate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGenerate.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGenerate.Location = new System.Drawing.Point(193, 298);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(146, 23);
            this.btnGenerate.TabIndex = 0;
            this.btnGenerate.Text = "Generate CBC xml";
            this.btnGenerate.UseVisualStyleBackColor = false;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Verdana", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.SteelBlue;
            this.label1.Location = new System.Drawing.Point(56, 70);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(119, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Source .xlsx File:";
            // 
            // txtSource
            // 
            this.txtSource.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtSource.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSource.ForeColor = System.Drawing.Color.Black;
            this.txtSource.Location = new System.Drawing.Point(181, 62);
            this.txtSource.Name = "txtSource";
            this.txtSource.ReadOnly = true;
            this.txtSource.Size = new System.Drawing.Size(233, 21);
            this.txtSource.TabIndex = 2;
            // 
            // brnBrowseSource
            // 
            this.brnBrowseSource.BackColor = System.Drawing.Color.LightSteelBlue;
            this.brnBrowseSource.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.brnBrowseSource.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.brnBrowseSource.ForeColor = System.Drawing.Color.Black;
            this.brnBrowseSource.Location = new System.Drawing.Point(420, 62);
            this.brnBrowseSource.Name = "brnBrowseSource";
            this.brnBrowseSource.Size = new System.Drawing.Size(92, 23);
            this.brnBrowseSource.TabIndex = 3;
            this.brnBrowseSource.Text = "Browse...";
            this.brnBrowseSource.UseVisualStyleBackColor = false;
            this.brnBrowseSource.Click += new System.EventHandler(this.brnBrowseSource_Click);
            // 
            // btnBrowseDest
            // 
            this.btnBrowseDest.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnBrowseDest.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowseDest.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowseDest.ForeColor = System.Drawing.Color.Black;
            this.btnBrowseDest.Location = new System.Drawing.Point(420, 105);
            this.btnBrowseDest.Name = "btnBrowseDest";
            this.btnBrowseDest.Size = new System.Drawing.Size(92, 23);
            this.btnBrowseDest.TabIndex = 6;
            this.btnBrowseDest.Text = "Browse...";
            this.btnBrowseDest.UseVisualStyleBackColor = false;
            this.btnBrowseDest.Click += new System.EventHandler(this.btnBrowseDest_Click);
            // 
            // txtDestFolder
            // 
            this.txtDestFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtDestFolder.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDestFolder.ForeColor = System.Drawing.Color.Black;
            this.txtDestFolder.Location = new System.Drawing.Point(181, 108);
            this.txtDestFolder.Name = "txtDestFolder";
            this.txtDestFolder.Size = new System.Drawing.Size(233, 21);
            this.txtDestFolder.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Verdana", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.SteelBlue;
            this.label2.Location = new System.Drawing.Point(44, 113);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(131, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Destination Folder:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.SteelBlue;
            this.label3.Location = new System.Drawing.Point(157, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(289, 18);
            this.label3.TabIndex = 7;
            this.label3.Text = "Country By Country Report v2.0";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtDestFileName
            // 
            this.txtDestFileName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtDestFileName.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDestFileName.ForeColor = System.Drawing.Color.Black;
            this.txtDestFileName.Location = new System.Drawing.Point(181, 153);
            this.txtDestFileName.Name = "txtDestFileName";
            this.txtDestFileName.Size = new System.Drawing.Size(233, 21);
            this.txtDestFileName.TabIndex = 9;
            this.txtDestFileName.Text = "CBC";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Verdana", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.SteelBlue;
            this.label4.Location = new System.Drawing.Point(21, 155);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(154, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Destination File Name:";
            // 
            // lstLog
            // 
            this.lstLog.FormattingEnabled = true;
            this.lstLog.HorizontalScrollbar = true;
            this.lstLog.Location = new System.Drawing.Point(24, 344);
            this.lstLog.Name = "lstLog";
            this.lstLog.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.lstLog.Size = new System.Drawing.Size(488, 290);
            this.lstLog.TabIndex = 10;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rad8);
            this.groupBox1.Controls.Add(this.rad16);
            this.groupBox1.Font = new System.Drawing.Font("Verdana", 8F, System.Drawing.FontStyle.Bold);
            this.groupBox1.ForeColor = System.Drawing.Color.SteelBlue;
            this.groupBox1.Location = new System.Drawing.Point(59, 198);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(355, 53);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "UTF Encoding:";
            // 
            // rad8
            // 
            this.rad8.AutoSize = true;
            this.rad8.Location = new System.Drawing.Point(122, 26);
            this.rad8.Name = "rad8";
            this.rad8.Size = new System.Drawing.Size(228, 17);
            this.rad8.TabIndex = 1;
            this.rad8.Text = "UTF-8   (SA or other Countries)";
            this.rad8.UseVisualStyleBackColor = true;
            // 
            // rad16
            // 
            this.rad16.AutoSize = true;
            this.rad16.Checked = true;
            this.rad16.Location = new System.Drawing.Point(122, 0);
            this.rad16.Name = "rad16";
            this.rad16.Size = new System.Drawing.Size(137, 17);
            this.rad16.TabIndex = 0;
            this.rad16.TabStop = true;
            this.rad16.Text = "UTF-16 (SA only)";
            this.rad16.UseVisualStyleBackColor = true;
            // 
            // ckSars
            // 
            this.ckSars.AutoSize = true;
            this.ckSars.Checked = true;
            this.ckSars.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ckSars.ForeColor = System.Drawing.Color.SteelBlue;
            this.ckSars.Location = new System.Drawing.Point(181, 262);
            this.ckSars.Name = "ckSars";
            this.ckSars.Size = new System.Drawing.Size(15, 14);
            this.ckSars.TabIndex = 12;
            this.ckSars.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Verdana", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.SteelBlue;
            this.label5.Location = new System.Drawing.Point(75, 262);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(89, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Is SARS CBC:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Verdana", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.SteelBlue;
            this.label6.Location = new System.Drawing.Point(202, 263);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(68, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "(SA only)";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(540, 669);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.ckSars);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lstLog);
            this.Controls.Add(this.txtDestFileName);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnBrowseDest);
            this.Controls.Add(this.txtDestFolder);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.brnBrowseSource);
            this.Controls.Add(this.txtSource);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnGenerate);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form2";
            this.Text = "CBC Report";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion


        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSource;
        private System.Windows.Forms.Button brnBrowseSource;
        private System.Windows.Forms.Button btnBrowseDest;
        private System.Windows.Forms.TextBox txtDestFolder;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TextBox txtDestFileName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListBox lstLog;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rad8;
        private System.Windows.Forms.RadioButton rad16;
        private System.Windows.Forms.CheckBox ckSars;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
    }
}

