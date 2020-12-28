
namespace Thesis_Excel
{
    partial class main
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
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
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.Upload = new System.Windows.Forms.Button();
            this.Execel_preview = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.Name_label = new System.Windows.Forms.Label();
            this.FileName = new System.Windows.Forms.TextBox();
            this.Convert = new System.Windows.Forms.Button();
            this.status_label = new System.Windows.Forms.Label();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.status_content = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.Execel_preview)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // Upload
            // 
            this.Upload.Location = new System.Drawing.Point(812, 338);
            this.Upload.Name = "Upload";
            this.Upload.Size = new System.Drawing.Size(80, 34);
            this.Upload.TabIndex = 0;
            this.Upload.Text = "Upload";
            this.Upload.UseVisualStyleBackColor = true;
            this.Upload.Click += new System.EventHandler(this.Upload_Click);
            // 
            // Execel_preview
            // 
            this.Execel_preview.AllowUserToAddRows = false;
            this.Execel_preview.AllowUserToDeleteRows = false;
            this.Execel_preview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Execel_preview.Location = new System.Drawing.Point(2, 0);
            this.Execel_preview.Name = "Execel_preview";
            this.Execel_preview.ReadOnly = true;
            this.Execel_preview.RowHeadersWidth = 62;
            this.Execel_preview.RowTemplate.Height = 31;
            this.Execel_preview.Size = new System.Drawing.Size(1044, 330);
            this.Execel_preview.TabIndex = 1;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 447F));
            this.tableLayoutPanel1.Controls.Add(this.Name_label, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.FileName, 1, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(296, 336);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(510, 44);
            this.tableLayoutPanel1.TabIndex = 2;
            // 
            // Name_label
            // 
            this.Name_label.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Name_label.AutoSize = true;
            this.Name_label.Location = new System.Drawing.Point(9, 13);
            this.Name_label.Name = "Name_label";
            this.Name_label.Size = new System.Drawing.Size(44, 18);
            this.Name_label.TabIndex = 0;
            this.Name_label.Text = "檔名";
            // 
            // FileName
            // 
            this.FileName.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.FileName.Location = new System.Drawing.Point(66, 7);
            this.FileName.Name = "FileName";
            this.FileName.Size = new System.Drawing.Size(441, 29);
            this.FileName.TabIndex = 1;
            // 
            // Convert
            // 
            this.Convert.Location = new System.Drawing.Point(920, 396);
            this.Convert.Name = "Convert";
            this.Convert.Size = new System.Drawing.Size(116, 34);
            this.Convert.TabIndex = 3;
            this.Convert.Text = "Convert";
            this.Convert.UseVisualStyleBackColor = true;
            this.Convert.Click += new System.EventHandler(this.Convert_Click);
            // 
            // status_label
            // 
            this.status_label.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.status_label.AutoSize = true;
            this.status_label.Location = new System.Drawing.Point(3, 6);
            this.status_label.Name = "status_label";
            this.status_label.Size = new System.Drawing.Size(49, 18);
            this.status_label.TabIndex = 4;
            this.status_label.Text = "格式:";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 144F));
            this.tableLayoutPanel2.Controls.Add(this.status_label, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.status_content, 1, 0);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(12, 336);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(200, 31);
            this.tableLayoutPanel2.TabIndex = 5;
            // 
            // status_content
            // 
            this.status_content.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.status_content.AutoSize = true;
            this.status_content.Location = new System.Drawing.Point(59, 6);
            this.status_content.Name = "status_content";
            this.status_content.Size = new System.Drawing.Size(62, 18);
            this.status_content.TabIndex = 5;
            this.status_content.Text = "未輸入";
            // 
            // main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1048, 442);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Controls.Add(this.Convert);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.Execel_preview);
            this.Controls.Add(this.Upload);
            this.Name = "main";
            this.Text = "Convert";
            ((System.ComponentModel.ISupportInitialize)(this.Execel_preview)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Upload;
        private System.Windows.Forms.DataGridView Execel_preview;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label Name_label;
        private System.Windows.Forms.TextBox FileName;
        private System.Windows.Forms.Button Convert;
        private System.Windows.Forms.Label status_label;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Label status_content;
    }
}

