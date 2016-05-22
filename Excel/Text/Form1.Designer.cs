namespace Text
{
    partial class Form1
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnImport = new System.Windows.Forms.Button();
            this.tbName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbFile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tbKeywords = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnSavePath = new System.Windows.Forms.Button();
            this.tbSavePath = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnComfirm = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(287, 51);
            this.btnImport.Margin = new System.Windows.Forms.Padding(2);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(32, 22);
            this.btnImport.TabIndex = 0;
            this.btnImport.Text = "...";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // tbName
            // 
            this.tbName.Location = new System.Drawing.Point(86, 18);
            this.tbName.Margin = new System.Windows.Forms.Padding(2);
            this.tbName.Name = "tbName";
            this.tbName.Size = new System.Drawing.Size(198, 21);
            this.tbName.TabIndex = 1;
            this.tbName.Text = "vivo";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(43, 26);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "产品：";
            // 
            // tbFile
            // 
            this.tbFile.Location = new System.Drawing.Point(86, 53);
            this.tbFile.Margin = new System.Windows.Forms.Padding(2);
            this.tbFile.Name = "tbFile";
            this.tbFile.Size = new System.Drawing.Size(198, 21);
            this.tbFile.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 61);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "txt文件：";
            // 
            // tbKeywords
            // 
            this.tbKeywords.Location = new System.Drawing.Point(86, 117);
            this.tbKeywords.Margin = new System.Windows.Forms.Padding(2);
            this.tbKeywords.Multiline = true;
            this.tbKeywords.Name = "tbKeywords";
            this.tbKeywords.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbKeywords.Size = new System.Drawing.Size(198, 124);
            this.tbKeywords.TabIndex = 3;
            this.tbKeywords.Text = "苏宁\r\n大中\r\n国美";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(32, 125);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "关键字：";
            // 
            // btnSavePath
            // 
            this.btnSavePath.Location = new System.Drawing.Point(287, 82);
            this.btnSavePath.Margin = new System.Windows.Forms.Padding(2);
            this.btnSavePath.Name = "btnSavePath";
            this.btnSavePath.Size = new System.Drawing.Size(32, 21);
            this.btnSavePath.TabIndex = 0;
            this.btnSavePath.Text = "...";
            this.btnSavePath.UseVisualStyleBackColor = true;
            this.btnSavePath.Click += new System.EventHandler(this.btnSavePath_Click);
            // 
            // tbSavePath
            // 
            this.tbSavePath.Location = new System.Drawing.Point(86, 85);
            this.tbSavePath.Margin = new System.Windows.Forms.Padding(2);
            this.tbSavePath.Name = "tbSavePath";
            this.tbSavePath.Size = new System.Drawing.Size(198, 21);
            this.tbSavePath.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 93);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "保存路径：";
            // 
            // btnComfirm
            // 
            this.btnComfirm.Location = new System.Drawing.Point(86, 245);
            this.btnComfirm.Margin = new System.Windows.Forms.Padding(2);
            this.btnComfirm.Name = "btnComfirm";
            this.btnComfirm.Size = new System.Drawing.Size(196, 36);
            this.btnComfirm.TabIndex = 5;
            this.btnComfirm.Text = "确  定";
            this.btnComfirm.UseVisualStyleBackColor = true;
            this.btnComfirm.Click += new System.EventHandler(this.btnComfirm_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(370, 298);
            this.Controls.Add(this.btnComfirm);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tbKeywords);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbSavePath);
            this.Controls.Add(this.tbFile);
            this.Controls.Add(this.btnSavePath);
            this.Controls.Add(this.tbName);
            this.Controls.Add(this.btnImport);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "转化Text";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.TextBox tbName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbKeywords;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnSavePath;
        private System.Windows.Forms.TextBox tbSavePath;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnComfirm;
    }
}

