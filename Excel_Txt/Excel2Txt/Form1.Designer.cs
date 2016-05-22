namespace Excel2Txt
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
            this.btn1 = new System.Windows.Forms.Button();
            this.listBoxOptions = new System.Windows.Forms.CheckedListBox();
            this.btn2 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.cbSheets = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn1
            // 
            this.btn1.Location = new System.Drawing.Point(44, 12);
            this.btn1.Name = "btn1";
            this.btn1.Size = new System.Drawing.Size(341, 43);
            this.btn1.TabIndex = 0;
            this.btn1.Text = "选择Excel文件";
            this.btn1.UseVisualStyleBackColor = true;
            this.btn1.Click += new System.EventHandler(this.btn1_Click);
            // 
            // listBoxOptions
            // 
            this.listBoxOptions.CheckOnClick = true;
            this.listBoxOptions.FormattingEnabled = true;
            this.listBoxOptions.Location = new System.Drawing.Point(186, 61);
            this.listBoxOptions.Name = "listBoxOptions";
            this.listBoxOptions.Size = new System.Drawing.Size(199, 84);
            this.listBoxOptions.TabIndex = 1;
            // 
            // btn2
            // 
            this.btn2.Location = new System.Drawing.Point(44, 151);
            this.btn2.Name = "btn2";
            this.btn2.Size = new System.Drawing.Size(341, 40);
            this.btn2.TabIndex = 2;
            this.btn2.Text = "转换为Txt文件";
            this.btn2.UseVisualStyleBackColor = true;
            this.btn2.Click += new System.EventHandler(this.btn2_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("楷体", 10.5F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Italic | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(323, 223);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(105, 14);
            this.label5.TabIndex = 7;
            this.label5.Text = "made by JLiang";
            // 
            // cbSheets
            // 
            this.cbSheets.FormattingEnabled = true;
            this.cbSheets.Location = new System.Drawing.Point(44, 76);
            this.cbSheets.Name = "cbSheets";
            this.cbSheets.Size = new System.Drawing.Size(125, 20);
            this.cbSheets.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(42, 61);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 9;
            this.label1.Text = "表名称：";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(440, 246);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbSheets);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btn2);
            this.Controls.Add(this.listBoxOptions);
            this.Controls.Add(this.btn1);
            this.Name = "Form1";
            this.Text = "Excel2Txt";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn1;
        private System.Windows.Forms.CheckedListBox listBoxOptions;
        private System.Windows.Forms.Button btn2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cbSheets;
        private System.Windows.Forms.Label label1;
    }
}

