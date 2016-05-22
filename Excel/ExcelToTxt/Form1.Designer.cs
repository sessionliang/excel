namespace ExcelToTxt
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnComfirm = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tbSavePath = new System.Windows.Forms.TextBox();
            this.tbFile = new System.Windows.Forms.TextBox();
            this.btnSavePath = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnComfirm
            // 
            this.btnComfirm.Location = new System.Drawing.Point(86, 121);
            this.btnComfirm.Margin = new System.Windows.Forms.Padding(2);
            this.btnComfirm.Name = "btnComfirm";
            this.btnComfirm.Size = new System.Drawing.Size(196, 36);
            this.btnComfirm.TabIndex = 16;
            this.btnComfirm.Text = "确  定";
            this.btnComfirm.UseVisualStyleBackColor = true;
            this.btnComfirm.Click += new System.EventHandler(this.btnComfirm_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 83);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 11;
            this.label4.Text = "保存路径：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 51);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 12);
            this.label2.TabIndex = 12;
            this.label2.Text = "excel文件：";
            // 
            // tbSavePath
            // 
            this.tbSavePath.Location = new System.Drawing.Point(86, 75);
            this.tbSavePath.Margin = new System.Windows.Forms.Padding(2);
            this.tbSavePath.Name = "tbSavePath";
            this.tbSavePath.Size = new System.Drawing.Size(198, 21);
            this.tbSavePath.TabIndex = 8;
            // 
            // tbFile
            // 
            this.tbFile.Location = new System.Drawing.Point(86, 43);
            this.tbFile.Margin = new System.Windows.Forms.Padding(2);
            this.tbFile.Name = "tbFile";
            this.tbFile.Size = new System.Drawing.Size(198, 21);
            this.tbFile.TabIndex = 9;
            // 
            // btnSavePath
            // 
            this.btnSavePath.Location = new System.Drawing.Point(287, 72);
            this.btnSavePath.Margin = new System.Windows.Forms.Padding(2);
            this.btnSavePath.Name = "btnSavePath";
            this.btnSavePath.Size = new System.Drawing.Size(32, 21);
            this.btnSavePath.TabIndex = 6;
            this.btnSavePath.Text = "...";
            this.btnSavePath.UseVisualStyleBackColor = true;
            this.btnSavePath.Click += new System.EventHandler(this.btnSavePath_Click);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(287, 41);
            this.btnImport.Margin = new System.Windows.Forms.Padding(2);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(32, 22);
            this.btnImport.TabIndex = 7;
            this.btnImport.Text = "...";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(345, 201);
            this.Controls.Add(this.btnComfirm);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tbSavePath);
            this.Controls.Add(this.tbFile);
            this.Controls.Add(this.btnSavePath);
            this.Controls.Add(this.btnImport);
            this.Name = "Form1";
            this.Text = "Excel转换为Txt";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnComfirm;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbSavePath;
        private System.Windows.Forms.TextBox tbFile;
        private System.Windows.Forms.Button btnSavePath;
        private System.Windows.Forms.Button btnImport;
    }
}

