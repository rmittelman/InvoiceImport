﻿namespace InvoiceImport
{
    partial class frmImport
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
            this.label1 = new System.Windows.Forms.Label();
            this.btnFindExcel = new System.Windows.Forms.Button();
            this.txtExcelFile = new System.Windows.Forms.TextBox();
            this.txtPDFFile = new System.Windows.Forms.TextBox();
            this.btnFindPDF = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btnImportStandard = new System.Windows.Forms.Button();
            this.btnSplitPDF = new System.Windows.Forms.Button();
            this.txtPDFFolder = new System.Windows.Forms.TextBox();
            this.btnFindFolder = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.txtStatus = new System.Windows.Forms.TextBox();
            this.ckAddToAIMM = new System.Windows.Forms.CheckBox();
            this.ckAddToQB = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(34, 44);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(180, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Excel Invoice File";
            // 
            // btnFindExcel
            // 
            this.btnFindExcel.Font = new System.Drawing.Font("Wingdings", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.btnFindExcel.Location = new System.Drawing.Point(222, 39);
            this.btnFindExcel.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnFindExcel.Name = "btnFindExcel";
            this.btnFindExcel.Size = new System.Drawing.Size(48, 41);
            this.btnFindExcel.TabIndex = 1;
            this.btnFindExcel.Text = "1";
            this.btnFindExcel.UseVisualStyleBackColor = true;
            // 
            // txtExcelFile
            // 
            this.txtExcelFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtExcelFile.Location = new System.Drawing.Point(280, 39);
            this.txtExcelFile.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.txtExcelFile.Name = "txtExcelFile";
            this.txtExcelFile.Size = new System.Drawing.Size(916, 31);
            this.txtExcelFile.TabIndex = 2;
            // 
            // txtPDFFile
            // 
            this.txtPDFFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPDFFile.Location = new System.Drawing.Point(280, 198);
            this.txtPDFFile.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.txtPDFFile.Name = "txtPDFFile";
            this.txtPDFFile.Size = new System.Drawing.Size(916, 31);
            this.txtPDFFile.TabIndex = 5;
            // 
            // btnFindPDF
            // 
            this.btnFindPDF.Font = new System.Drawing.Font("Wingdings", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.btnFindPDF.Location = new System.Drawing.Point(222, 198);
            this.btnFindPDF.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnFindPDF.Name = "btnFindPDF";
            this.btnFindPDF.Size = new System.Drawing.Size(48, 41);
            this.btnFindPDF.TabIndex = 4;
            this.btnFindPDF.Text = "1";
            this.btnFindPDF.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(36, 203);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 25);
            this.label2.TabIndex = 3;
            this.label2.Text = "PDF File";
            // 
            // btnImportStandard
            // 
            this.btnImportStandard.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnImportStandard.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImportStandard.Location = new System.Drawing.Point(1216, 33);
            this.btnImportStandard.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnImportStandard.Name = "btnImportStandard";
            this.btnImportStandard.Size = new System.Drawing.Size(190, 41);
            this.btnImportStandard.TabIndex = 6;
            this.btnImportStandard.Text = "Import Invoices";
            this.btnImportStandard.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnImportStandard.UseVisualStyleBackColor = true;
            this.btnImportStandard.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnSplitPDF
            // 
            this.btnSplitPDF.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSplitPDF.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSplitPDF.Location = new System.Drawing.Point(1216, 216);
            this.btnSplitPDF.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnSplitPDF.Name = "btnSplitPDF";
            this.btnSplitPDF.Size = new System.Drawing.Size(190, 61);
            this.btnSplitPDF.TabIndex = 7;
            this.btnSplitPDF.Text = "Split PDF Files";
            this.btnSplitPDF.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnSplitPDF.UseVisualStyleBackColor = true;
            this.btnSplitPDF.Click += new System.EventHandler(this.btnSplitPDF_Click);
            // 
            // txtPDFFolder
            // 
            this.txtPDFFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPDFFolder.Location = new System.Drawing.Point(280, 256);
            this.txtPDFFolder.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.txtPDFFolder.Name = "txtPDFFolder";
            this.txtPDFFolder.Size = new System.Drawing.Size(916, 31);
            this.txtPDFFolder.TabIndex = 10;
            // 
            // btnFindFolder
            // 
            this.btnFindFolder.Font = new System.Drawing.Font("Wingdings", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.btnFindFolder.Location = new System.Drawing.Point(222, 256);
            this.btnFindFolder.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnFindFolder.Name = "btnFindFolder";
            this.btnFindFolder.Size = new System.Drawing.Size(48, 41);
            this.btnFindFolder.TabIndex = 9;
            this.btnFindFolder.Text = "1";
            this.btnFindFolder.UseVisualStyleBackColor = true;
            this.btnFindFolder.Click += new System.EventHandler(this.btnFindFolder_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(36, 261);
            this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(185, 25);
            this.label3.TabIndex = 8;
            this.label3.Text = "Save PDF Files to";
            // 
            // lblStatus
            // 
            this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(24, 356);
            this.lblStatus.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(79, 25);
            this.lblStatus.TabIndex = 13;
            this.lblStatus.Text = "Status:";
            // 
            // txtStatus
            // 
            this.txtStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtStatus.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtStatus.Location = new System.Drawing.Point(116, 344);
            this.txtStatus.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.txtStatus.Multiline = true;
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtStatus.Size = new System.Drawing.Size(1272, 123);
            this.txtStatus.TabIndex = 14;
            // 
            // ckAddToAIMM
            // 
            this.ckAddToAIMM.AutoSize = true;
            this.ckAddToAIMM.Location = new System.Drawing.Point(432, 89);
            this.ckAddToAIMM.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.ckAddToAIMM.Name = "ckAddToAIMM";
            this.ckAddToAIMM.Size = new System.Drawing.Size(252, 29);
            this.ckAddToAIMM.TabIndex = 16;
            this.ckAddToAIMM.Text = "Add invoices to AIMM";
            this.ckAddToAIMM.UseVisualStyleBackColor = true;
            // 
            // ckAddToQB
            // 
            this.ckAddToQB.AutoSize = true;
            this.ckAddToQB.Location = new System.Drawing.Point(790, 89);
            this.ckAddToQB.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.ckAddToQB.Name = "ckAddToQB";
            this.ckAddToQB.Size = new System.Drawing.Size(312, 29);
            this.ckAddToQB.TabIndex = 17;
            this.ckAddToQB.Text = "Add invoices to QuickBooks";
            this.ckAddToQB.UseVisualStyleBackColor = true;
            // 
            // frmImport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1430, 522);
            this.Controls.Add(this.ckAddToQB);
            this.Controls.Add(this.ckAddToAIMM);
            this.Controls.Add(this.txtStatus);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.txtPDFFolder);
            this.Controls.Add(this.btnFindFolder);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnSplitPDF);
            this.Controls.Add(this.btnImportStandard);
            this.Controls.Add(this.txtPDFFile);
            this.Controls.Add(this.btnFindPDF);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtExcelFile);
            this.Controls.Add(this.btnFindExcel);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.Name = "frmImport";
            this.Text = "Import Invoices";
            this.Load += new System.EventHandler(this.frmImport_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnFindExcel;
        private System.Windows.Forms.TextBox txtExcelFile;
        private System.Windows.Forms.TextBox txtPDFFile;
        private System.Windows.Forms.Button btnFindPDF;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnImportStandard;
        private System.Windows.Forms.Button btnSplitPDF;
        private System.Windows.Forms.TextBox txtPDFFolder;
        private System.Windows.Forms.Button btnFindFolder;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.TextBox txtStatus;
        private System.Windows.Forms.CheckBox ckAddToAIMM;
        private System.Windows.Forms.CheckBox ckAddToQB;
    }
}

