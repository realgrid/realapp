namespace RealAppsExcel
{
    partial class CodeForm
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
            this.TxtSourceCode = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnPreview = new System.Windows.Forms.Button();
            this.BtnCopy = new System.Windows.Forms.Button();
            this.TxtPreviewUri = new System.Windows.Forms.TextBox();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // TxtSourceCode
            // 
            this.TxtSourceCode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TxtSourceCode.Location = new System.Drawing.Point(12, 33);
            this.TxtSourceCode.Multiline = true;
            this.TxtSourceCode.Name = "TxtSourceCode";
            this.TxtSourceCode.ReadOnly = true;
            this.TxtSourceCode.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.TxtSourceCode.Size = new System.Drawing.Size(751, 157);
            this.TxtSourceCode.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "소스 코드";
            // 
            // BtnPreview
            // 
            this.BtnPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnPreview.Location = new System.Drawing.Point(650, 196);
            this.BtnPreview.Name = "BtnPreview";
            this.BtnPreview.Size = new System.Drawing.Size(113, 23);
            this.BtnPreview.TabIndex = 2;
            this.BtnPreview.Text = "미리보기(TCP)";
            this.BtnPreview.UseVisualStyleBackColor = true;
            this.BtnPreview.Click += new System.EventHandler(this.BtnPreview_Click);
            // 
            // BtnCopy
            // 
            this.BtnCopy.Location = new System.Drawing.Point(12, 196);
            this.BtnCopy.Name = "BtnCopy";
            this.BtnCopy.Size = new System.Drawing.Size(75, 23);
            this.BtnCopy.TabIndex = 3;
            this.BtnCopy.Text = "코드 복사";
            this.BtnCopy.UseVisualStyleBackColor = true;
            this.BtnCopy.Click += new System.EventHandler(this.BtnCopy_Click);
            // 
            // TxtPreviewUri
            // 
            this.TxtPreviewUri.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.TxtPreviewUri.BackColor = System.Drawing.SystemColors.Control;
            this.TxtPreviewUri.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TxtPreviewUri.Location = new System.Drawing.Point(288, 201);
            this.TxtPreviewUri.Name = "TxtPreviewUri";
            this.TxtPreviewUri.Size = new System.Drawing.Size(356, 14);
            this.TxtPreviewUri.TabIndex = 4;
            this.TxtPreviewUri.Text = "http://realapp.co.kr/excel-addins/col-preview.html";
            this.TxtPreviewUri.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // webBrowser1
            // 
            this.webBrowser1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.webBrowser1.Location = new System.Drawing.Point(12, 221);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(751, 292);
            this.webBrowser1.TabIndex = 5;
            this.webBrowser1.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.WebBrowser1_DocumentCompleted);
            this.webBrowser1.Navigated += new System.Windows.Forms.WebBrowserNavigatedEventHandler(this.WebBrowser1_Navigated);
            // 
            // CodeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(775, 525);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.TxtPreviewUri);
            this.Controls.Add(this.BtnCopy);
            this.Controls.Add(this.BtnPreview);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TxtSourceCode);
            this.Name = "CodeForm";
            this.Text = "코드및 미리보기";
            this.Load += new System.EventHandler(this.CodeForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TxtSourceCode;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BtnPreview;
        private System.Windows.Forms.Button BtnCopy;
        private System.Windows.Forms.TextBox TxtPreviewUri;
        private System.Windows.Forms.WebBrowser webBrowser1;
    }
}