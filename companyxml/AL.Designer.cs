namespace companyxml
{
    partial class AL
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.monthC = new System.Windows.Forms.ComboBox();
            this.yearC = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(78, 192);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(417, 354);
            this.dataGridView1.TabIndex = 22;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.LightSalmon;
            this.button1.Location = new System.Drawing.Point(376, 13);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(119, 39);
            this.button1.TabIndex = 21;
            this.button1.Text = "重新整理";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("標楷體", 25.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(487, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 43);
            this.label2.TabIndex = 20;
            this.label2.Text = "月";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("標楷體", 25.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(241, 102);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 43);
            this.label1.TabIndex = 19;
            this.label1.Text = "年";
            // 
            // monthC
            // 
            this.monthC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.monthC.Font = new System.Drawing.Font("新細明體", 25.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.monthC.FormattingEnabled = true;
            this.monthC.Items.AddRange(new object[] {
            "01",
            "02",
            "03",
            "04",
            "05",
            "06",
            "07",
            "08",
            "09",
            "10",
            "11",
            "12"});
            this.monthC.Location = new System.Drawing.Point(310, 94);
            this.monthC.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.monthC.Name = "monthC";
            this.monthC.Size = new System.Drawing.Size(171, 51);
            this.monthC.TabIndex = 18;
            this.monthC.SelectedIndexChanged += new System.EventHandler(this.monthC_SelectedIndexChanged);
            // 
            // yearC
            // 
            this.yearC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.yearC.Font = new System.Drawing.Font("新細明體", 25.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.yearC.FormattingEnabled = true;
            this.yearC.Location = new System.Drawing.Point(64, 95);
            this.yearC.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.yearC.Name = "yearC";
            this.yearC.Size = new System.Drawing.Size(171, 51);
            this.yearC.TabIndex = 17;
            // 
            // AL
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(610, 584);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.monthC);
            this.Controls.Add(this.yearC);
            this.Name = "AL";
            this.Text = "AL";
            this.Load += new System.EventHandler(this.AL_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox monthC;
        private System.Windows.Forms.ComboBox yearC;
    }
}