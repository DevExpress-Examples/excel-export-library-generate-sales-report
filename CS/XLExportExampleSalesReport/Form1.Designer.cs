namespace XLExportExampleSalesReport {
    partial class Form1 {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnExportToXLSX = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.btnExportToXLS = new System.Windows.Forms.Button();
            this.btnExportToCSV = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnExportToXLSX
            // 
            this.btnExportToXLSX.Image = ((System.Drawing.Image)(resources.GetObject("btnExportToXLSX.Image")));
            this.btnExportToXLSX.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExportToXLSX.Location = new System.Drawing.Point(11, 80);
            this.btnExportToXLSX.Name = "btnExportToXLSX";
            this.btnExportToXLSX.Size = new System.Drawing.Size(109, 28);
            this.btnExportToXLSX.TabIndex = 0;
            this.btnExportToXLSX.Text = "Export to XLSX";
            this.btnExportToXLSX.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnExportToXLSX.UseVisualStyleBackColor = true;
            this.btnExportToXLSX.Click += new System.EventHandler(this.btnExportToXLSX_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.Window;
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(8);
            this.panel1.Size = new System.Drawing.Size(365, 63);
            this.panel1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(8, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(311, 39);
            this.label1.TabIndex = 0;
            this.label1.Text = "Generate a sales report using the XL Export API.\r\nClick one of the buttons below " +
    "to save the document \r\nin the corresponding format.";
            // 
            // btnExportToXLS
            // 
            this.btnExportToXLS.Image = ((System.Drawing.Image)(resources.GetObject("btnExportToXLS.Image")));
            this.btnExportToXLS.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExportToXLS.Location = new System.Drawing.Point(126, 80);
            this.btnExportToXLS.Name = "btnExportToXLS";
            this.btnExportToXLS.Size = new System.Drawing.Size(109, 28);
            this.btnExportToXLS.TabIndex = 2;
            this.btnExportToXLS.Text = "Export to XLS";
            this.btnExportToXLS.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnExportToXLS.UseVisualStyleBackColor = true;
            this.btnExportToXLS.Click += new System.EventHandler(this.btnExportToXLS_Click);
            // 
            // btnExportToCSV
            // 
            this.btnExportToCSV.Image = ((System.Drawing.Image)(resources.GetObject("btnExportToCSV.Image")));
            this.btnExportToCSV.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExportToCSV.Location = new System.Drawing.Point(241, 80);
            this.btnExportToCSV.Name = "btnExportToCSV";
            this.btnExportToCSV.Size = new System.Drawing.Size(109, 28);
            this.btnExportToCSV.TabIndex = 3;
            this.btnExportToCSV.Text = "Export to CSV";
            this.btnExportToCSV.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnExportToCSV.UseVisualStyleBackColor = true;
            this.btnExportToCSV.Click += new System.EventHandler(this.btnExportToCSV_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(365, 120);
            this.Controls.Add(this.btnExportToCSV);
            this.Controls.Add(this.btnExportToXLS);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnExportToXLSX);
            this.Name = "Form1";
            this.Text = "XL Export example";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnExportToXLSX;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnExportToXLS;
        private System.Windows.Forms.Button btnExportToCSV;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

