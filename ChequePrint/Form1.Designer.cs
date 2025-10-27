namespace ChequePrint
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.txtAmount = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.lblAmountInWord = new System.Windows.Forms.Label();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.chkacpay = new System.Windows.Forms.CheckBox();
            this.rdSingle = new System.Windows.Forms.RadioButton();
            this.rdMultiple = new System.Windows.Forms.RadioButton();
            this.panelSingle = new System.Windows.Forms.Panel();
            this.panelMultiple = new System.Windows.Forms.Panel();
            this.btnDownload = new System.Windows.Forms.Button();
            this.lblDemo = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.txtBrowse = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lblUploadExcel = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.printDialog2 = new System.Windows.Forms.PrintDialog();
            this.printDocument2 = new System.Drawing.Printing.PrintDocument();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.panelSingle.SuspendLayout();
            this.panelMultiple.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(11, 59);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Pay To :";
            // 
            // txtName
            // 
            this.txtName.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtName.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtName.Location = new System.Drawing.Point(188, 60);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(161, 20);
            this.txtName.TabIndex = 2;
            // 
            // txtAmount
            // 
            this.txtAmount.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtAmount.CausesValidation = false;
            this.txtAmount.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtAmount.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.txtAmount.Location = new System.Drawing.Point(188, 106);
            this.txtAmount.Name = "txtAmount";
            this.txtAmount.Size = new System.Drawing.Size(161, 20);
            this.txtAmount.TabIndex = 3;
            this.txtAmount.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAmount_KeyPress);
            this.txtAmount.Leave += new System.EventHandler(this.txtAmount_Leave);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label2.Location = new System.Drawing.Point(11, 101);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 17);
            this.label2.TabIndex = 2;
            this.label2.Text = "Amount :";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.button1.Location = new System.Drawing.Point(261, 235);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(76, 31);
            this.button1.TabIndex = 6;
            this.button1.Text = "Print";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.button2.Location = new System.Drawing.Point(177, 235);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(76, 31);
            this.button2.TabIndex = 5;
            this.button2.Text = "Reset";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label3.Location = new System.Drawing.Point(11, 136);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(141, 17);
            this.label3.TabIndex = 8;
            this.label3.Text = "Amount In Words :";
            // 
            // lblAmountInWord
            // 
            this.lblAmountInWord.AutoSize = true;
            this.lblAmountInWord.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.lblAmountInWord.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAmountInWord.ForeColor = System.Drawing.SystemColors.WindowText;
            this.lblAmountInWord.Location = new System.Drawing.Point(188, 138);
            this.lblAmountInWord.Name = "lblAmountInWord";
            this.lblAmountInWord.Size = new System.Drawing.Size(0, 13);
            this.lblAmountInWord.TabIndex = 9;
            // 
            // printDocument1
            // 
            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.PrintPage);
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Document = this.printDocument1;
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            // 
            // printDialog1
            // 
            this.printDialog1.Document = this.printDocument1;
            this.printDialog1.UseEXDialog = true;
            // 
            // txtDate
            // 
            this.txtDate.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtDate.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtDate.Location = new System.Drawing.Point(188, 22);
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(161, 20);
            this.txtDate.TabIndex = 1;
            this.txtDate.Text = "DDMMYYYY";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label4.Location = new System.Drawing.Point(11, 21);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(112, 17);
            this.label4.TabIndex = 11;
            this.label4.Text = "Cheque Date :";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.label5.Font = new System.Drawing.Font("Malgun Gothic Semilight", 16F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label5.Location = new System.Drawing.Point(289, 52);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(183, 30);
            this.label5.TabIndex = 13;
            this.label5.Text = "Cheque Printing";
            // 
            // chkacpay
            // 
            this.chkacpay.AutoSize = true;
            this.chkacpay.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkacpay.Location = new System.Drawing.Point(189, 179);
            this.chkacpay.Name = "chkacpay";
            this.chkacpay.Size = new System.Drawing.Size(102, 21);
            this.chkacpay.TabIndex = 4;
            this.chkacpay.Text = "A/C Payee";
            this.chkacpay.UseVisualStyleBackColor = true;
            // 
            // rdSingle
            // 
            this.rdSingle.AutoSize = true;
            this.rdSingle.Checked = true;
            this.rdSingle.Location = new System.Drawing.Point(304, 97);
            this.rdSingle.Name = "rdSingle";
            this.rdSingle.Size = new System.Drawing.Size(57, 17);
            this.rdSingle.TabIndex = 15;
            this.rdSingle.TabStop = true;
            this.rdSingle.Text = "Single ";
            this.rdSingle.UseVisualStyleBackColor = true;
            this.rdSingle.Click += new System.EventHandler(this.rdSingle_Click);
            // 
            // rdMultiple
            // 
            this.rdMultiple.AutoSize = true;
            this.rdMultiple.Location = new System.Drawing.Point(388, 97);
            this.rdMultiple.Name = "rdMultiple";
            this.rdMultiple.Size = new System.Drawing.Size(61, 17);
            this.rdMultiple.TabIndex = 16;
            this.rdMultiple.Text = "Multiple";
            this.rdMultiple.UseVisualStyleBackColor = true;
            this.rdMultiple.Click += new System.EventHandler(this.rdMultiple_Click);
            // 
            // panelSingle
            // 
            this.panelSingle.Controls.Add(this.txtAmount);
            this.panelSingle.Controls.Add(this.label1);
            this.panelSingle.Controls.Add(this.txtName);
            this.panelSingle.Controls.Add(this.chkacpay);
            this.panelSingle.Controls.Add(this.label2);
            this.panelSingle.Controls.Add(this.button2);
            this.panelSingle.Controls.Add(this.label3);
            this.panelSingle.Controls.Add(this.button1);
            this.panelSingle.Controls.Add(this.txtDate);
            this.panelSingle.Controls.Add(this.lblAmountInWord);
            this.panelSingle.Controls.Add(this.label4);
            this.panelSingle.Location = new System.Drawing.Point(199, 120);
            this.panelSingle.Name = "panelSingle";
            this.panelSingle.Size = new System.Drawing.Size(394, 288);
            this.panelSingle.TabIndex = 17;
            // 
            // panelMultiple
            // 
            this.panelMultiple.Controls.Add(this.btnDownload);
            this.panelMultiple.Controls.Add(this.lblDemo);
            this.panelMultiple.Controls.Add(this.button3);
            this.panelMultiple.Controls.Add(this.button4);
            this.panelMultiple.Controls.Add(this.txtBrowse);
            this.panelMultiple.Controls.Add(this.btnBrowse);
            this.panelMultiple.Controls.Add(this.lblUploadExcel);
            this.panelMultiple.Location = new System.Drawing.Point(200, 123);
            this.panelMultiple.Name = "panelMultiple";
            this.panelMultiple.Size = new System.Drawing.Size(394, 282);
            this.panelMultiple.TabIndex = 18;
            // 
            // btnDownload
            // 
            this.btnDownload.Location = new System.Drawing.Point(238, 49);
            this.btnDownload.Name = "btnDownload";
            this.btnDownload.Size = new System.Drawing.Size(75, 23);
            this.btnDownload.TabIndex = 11;
            this.btnDownload.Text = "Download";
            this.btnDownload.UseVisualStyleBackColor = true;
            this.btnDownload.Click += new System.EventHandler(this.btnDownload_Click);
            // 
            // lblDemo
            // 
            this.lblDemo.AutoSize = true;
            this.lblDemo.Location = new System.Drawing.Point(33, 52);
            this.lblDemo.Name = "lblDemo";
            this.lblDemo.Size = new System.Drawing.Size(105, 13);
            this.lblDemo.TabIndex = 10;
            this.lblDemo.Text = "Download Demo File";
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.ForeColor = System.Drawing.SystemColors.WindowText;
            this.button3.Location = new System.Drawing.Point(146, 220);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(80, 31);
            this.button3.TabIndex = 9;
            this.button3.Text = "Reset";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.ForeColor = System.Drawing.SystemColors.WindowText;
            this.button4.Location = new System.Drawing.Point(230, 220);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(80, 31);
            this.button4.TabIndex = 8;
            this.button4.Text = "Print";
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // txtBrowse
            // 
            this.txtBrowse.Location = new System.Drawing.Point(162, 85);
            this.txtBrowse.Name = "txtBrowse";
            this.txtBrowse.ReadOnly = true;
            this.txtBrowse.Size = new System.Drawing.Size(151, 20);
            this.txtBrowse.TabIndex = 2;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(238, 111);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 1;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // lblUploadExcel
            // 
            this.lblUploadExcel.AutoSize = true;
            this.lblUploadExcel.Location = new System.Drawing.Point(34, 90);
            this.lblUploadExcel.Name = "lblUploadExcel";
            this.lblUploadExcel.Size = new System.Drawing.Size(70, 13);
            this.lblUploadExcel.TabIndex = 0;
            this.lblUploadExcel.Text = "Upload Excel";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Title = "Upload Excel File";
            // 
            // printDialog2
            // 
            this.printDialog2.Document = this.printDocument2;
            this.printDialog2.UseEXDialog = true;
            // 
            // printDocument2
            // 
            this.printDocument2.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.PrintPage2);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::ChequePrint.Properties.Resources.images;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.CausesValidation = false;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.rdMultiple);
            this.Controls.Add(this.rdSingle);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.panelMultiple);
            this.Controls.Add(this.panelSingle);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Print Cheque";
            this.panelSingle.ResumeLayout(false);
            this.panelSingle.PerformLayout();
            this.panelMultiple.ResumeLayout(false);
            this.panelMultiple.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.TextBox txtAmount;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblAmountInWord;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog1;
        private System.Windows.Forms.PrintDialog printDialog1;
        private System.Windows.Forms.TextBox txtDate;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox chkacpay;
        private System.Windows.Forms.RadioButton rdSingle;
        private System.Windows.Forms.RadioButton rdMultiple;
        private System.Windows.Forms.Panel panelSingle;
        private System.Windows.Forms.Panel panelMultiple;
        private System.Windows.Forms.Label lblUploadExcel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtBrowse;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.PrintDialog printDialog2;
        private System.Drawing.Printing.PrintDocument printDocument2;
        private System.Windows.Forms.Label lblDemo;
        private System.Windows.Forms.Button btnDownload;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

