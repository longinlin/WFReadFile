namespace WFReadFile
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
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.txtBx_Log = new System.Windows.Forms.TextBox();
            this.txtBx_Excel = new System.Windows.Forms.TextBox();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.Lab_csv_count = new System.Windows.Forms.Label();
            this.button7 = new System.Windows.Forms.Button();
            this.txtBx_Csv2 = new System.Windows.Forms.TextBox();
            this.Lab_csv_count2 = new System.Windows.Forms.Label();
            this.button8 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button1.Location = new System.Drawing.Point(924, 264);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(176, 41);
            this.button1.TabIndex = 0;
            this.button1.Text = "3.建立資料";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(52, 413);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Visible = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button3.Location = new System.Drawing.Point(924, 326);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(176, 40);
            this.button3.TabIndex = 2;
            this.button3.Text = "4.比對DB";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button4.Location = new System.Drawing.Point(924, 72);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(176, 43);
            this.button4.TabIndex = 3;
            this.button4.Text = "1.選擇Log檔案";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // txtBx_Log
            // 
            this.txtBx_Log.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtBx_Log.Location = new System.Drawing.Point(26, 82);
            this.txtBx_Log.Name = "txtBx_Log";
            this.txtBx_Log.Size = new System.Drawing.Size(821, 33);
            this.txtBx_Log.TabIndex = 4;
            // 
            // txtBx_Excel
            // 
            this.txtBx_Excel.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtBx_Excel.Location = new System.Drawing.Point(26, 210);
            this.txtBx_Excel.Name = "txtBx_Excel";
            this.txtBx_Excel.Size = new System.Drawing.Size(821, 33);
            this.txtBx_Excel.TabIndex = 5;
            this.txtBx_Excel.Visible = false;
            // 
            // button5
            // 
            this.button5.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button5.Location = new System.Drawing.Point(924, 210);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(176, 33);
            this.button5.TabIndex = 6;
            this.button5.Text = "2.選擇CSV UDD檔";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Visible = false;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button6.Location = new System.Drawing.Point(924, 385);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(176, 40);
            this.button6.TabIndex = 7;
            this.button6.Text = "5.輸出異常報表";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // Lab_csv_count
            // 
            this.Lab_csv_count.AutoSize = true;
            this.Lab_csv_count.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Lab_csv_count.Location = new System.Drawing.Point(853, 216);
            this.Lab_csv_count.Name = "Lab_csv_count";
            this.Lab_csv_count.Size = new System.Drawing.Size(62, 21);
            this.Lab_csv_count.TabIndex = 8;
            this.Lab_csv_count.Text = "筆數2";
            this.Lab_csv_count.Visible = false;
            // 
            // button7
            // 
            this.button7.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button7.Location = new System.Drawing.Point(923, 147);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(176, 33);
            this.button7.TabIndex = 9;
            this.button7.Text = "2.2 選CSV NTUH";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // txtBx_Csv2
            // 
            this.txtBx_Csv2.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtBx_Csv2.Location = new System.Drawing.Point(25, 147);
            this.txtBx_Csv2.Name = "txtBx_Csv2";
            this.txtBx_Csv2.Size = new System.Drawing.Size(821, 33);
            this.txtBx_Csv2.TabIndex = 10;
            // 
            // Lab_csv_count2
            // 
            this.Lab_csv_count2.AutoSize = true;
            this.Lab_csv_count2.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Lab_csv_count2.Location = new System.Drawing.Point(856, 147);
            this.Lab_csv_count2.Name = "Lab_csv_count2";
            this.Lab_csv_count2.Size = new System.Drawing.Size(62, 21);
            this.Lab_csv_count2.TabIndex = 11;
            this.Lab_csv_count2.Text = "筆數1";
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(302, 413);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(75, 23);
            this.button8.TabIndex = 12;
            this.button8.Text = "button8";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Visible = false;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1118, 477);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.Lab_csv_count2);
            this.Controls.Add(this.txtBx_Csv2);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.Lab_csv_count);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.txtBx_Excel);
            this.Controls.Add(this.txtBx_Log);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Pyxis檢核";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox txtBx_Log;
        private System.Windows.Forms.TextBox txtBx_Excel;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Label Lab_csv_count;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.TextBox txtBx_Csv2;
        private System.Windows.Forms.Label Lab_csv_count2;
        private System.Windows.Forms.Button button8;
    }
}

