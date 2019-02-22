namespace COD02
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
      this.components = new System.ComponentModel.Container();
      this.label1 = new System.Windows.Forms.Label();
      this.txtFilePath = new System.Windows.Forms.TextBox();
      this.cmdOpenRead = new System.Windows.Forms.Button();
      this.cmdBrowse = new System.Windows.Forms.Button();
      this.button1 = new System.Windows.Forms.Button();
      this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
      this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
      this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
      this.dgvDataList = new System.Windows.Forms.DataGridView();
      this.label2 = new System.Windows.Forms.Label();
      this.comboBox1 = new System.Windows.Forms.ComboBox();
      ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.dgvDataList)).BeginInit();
      this.SuspendLayout();
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(54, 42);
      this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(98, 18);
      this.label1.TabIndex = 1;
      this.label1.Text = "資料來源：";
      // 
      // txtFilePath
      // 
      this.txtFilePath.Location = new System.Drawing.Point(160, 38);
      this.txtFilePath.Margin = new System.Windows.Forms.Padding(4);
      this.txtFilePath.Name = "txtFilePath";
      this.txtFilePath.Size = new System.Drawing.Size(847, 29);
      this.txtFilePath.TabIndex = 2;
      // 
      // cmdOpenRead
      // 
      this.cmdOpenRead.Location = new System.Drawing.Point(835, 88);
      this.cmdOpenRead.Margin = new System.Windows.Forms.Padding(4);
      this.cmdOpenRead.Name = "cmdOpenRead";
      this.cmdOpenRead.Size = new System.Drawing.Size(104, 34);
      this.cmdOpenRead.TabIndex = 3;
      this.cmdOpenRead.Text = "讀取";
      this.cmdOpenRead.UseVisualStyleBackColor = true;
      this.cmdOpenRead.Click += new System.EventHandler(this.cmdOpenRead_Click);
      // 
      // cmdBrowse
      // 
      this.cmdBrowse.Location = new System.Drawing.Point(1015, 33);
      this.cmdBrowse.Margin = new System.Windows.Forms.Padding(4);
      this.cmdBrowse.Name = "cmdBrowse";
      this.cmdBrowse.Size = new System.Drawing.Size(38, 34);
      this.cmdBrowse.TabIndex = 4;
      this.cmdBrowse.Text = "...";
      this.cmdBrowse.UseVisualStyleBackColor = true;
      this.cmdBrowse.Click += new System.EventHandler(this.cmdBrowse_Click);
      // 
      // button1
      // 
      this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
      this.button1.Enabled = false;
      this.button1.Location = new System.Drawing.Point(947, 88);
      this.button1.Margin = new System.Windows.Forms.Padding(4);
      this.button1.Name = "button1";
      this.button1.Size = new System.Drawing.Size(106, 34);
      this.button1.TabIndex = 5;
      this.button1.Text = "匯出手冊";
      this.button1.UseVisualStyleBackColor = false;
      this.button1.Click += new System.EventHandler(this.button1_Click);
      // 
      // openFileDialog1
      // 
      this.openFileDialog1.FileName = "data.xlsx";
      this.openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
      // 
      // dgvDataList
      // 
      this.dgvDataList.AllowUserToAddRows = false;
      this.dgvDataList.AllowUserToDeleteRows = false;
      this.dgvDataList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dgvDataList.Location = new System.Drawing.Point(8, 154);
      this.dgvDataList.Margin = new System.Windows.Forms.Padding(4);
      this.dgvDataList.Name = "dgvDataList";
      this.dgvDataList.ReadOnly = true;
      this.dgvDataList.RowTemplate.Height = 24;
      this.dgvDataList.Size = new System.Drawing.Size(1045, 441);
      this.dgvDataList.TabIndex = 8;
      this.dgvDataList.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvDataList_CellContentDoubleClick);
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Location = new System.Drawing.Point(36, 96);
      this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(116, 18);
      this.label2.TabIndex = 9;
      this.label2.Text = "選擇工作表：";
      // 
      // comboBox1
      // 
      this.comboBox1.FormattingEnabled = true;
      this.comboBox1.Location = new System.Drawing.Point(160, 96);
      this.comboBox1.Name = "comboBox1";
      this.comboBox1.Size = new System.Drawing.Size(668, 26);
      this.comboBox1.TabIndex = 10;
      this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
      // 
      // Form1
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(1066, 608);
      this.Controls.Add(this.comboBox1);
      this.Controls.Add(this.label2);
      this.Controls.Add(this.dgvDataList);
      this.Controls.Add(this.button1);
      this.Controls.Add(this.cmdBrowse);
      this.Controls.Add(this.cmdOpenRead);
      this.Controls.Add(this.txtFilePath);
      this.Controls.Add(this.label1);
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
      this.Margin = new System.Windows.Forms.Padding(4);
      this.Name = "Form1";
      this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
      this.Text = "Form1";
      ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.dgvDataList)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button cmdOpenRead;
        private System.Windows.Forms.Button cmdBrowse;
    private System.Windows.Forms.Button button1;
    private System.Windows.Forms.BindingSource bindingSource1;
    private System.Windows.Forms.OpenFileDialog openFileDialog1;
    private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    private System.Windows.Forms.DataGridView dgvDataList;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.ComboBox comboBox1;
  }
}

