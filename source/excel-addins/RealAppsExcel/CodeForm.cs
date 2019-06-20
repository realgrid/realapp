using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RealAppsExcel
{
    public partial class CodeForm : Form
    {
        private string fieldInfo;
        private string columnInfo;
        public string SourceCode
        {
            get
            {
                return this.TxtSourceCode.Text;
            }
            set
            {
                this.TxtSourceCode.Text = value;
            }
        }
        public string FieldInfo
        {
            get
            {
                return this.fieldInfo;
            }
            set
            {
                this.fieldInfo = value;
            }
        }
        public string ColumnInfo
        {
            get
            {
                return this.columnInfo;
            }
            set
            {
                this.columnInfo = value;
            }
        }

        public CodeForm()
        {
            InitializeComponent();
        }

        private void BtnPreview_Click(object sender, EventArgs e)
        {
            string url = TxtPreviewUri.Text;
            webBrowser1.Navigate(url);
        }

        private void BtnCopy_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Clipboard.SetText(TxtSourceCode.Text);
        }

        private void WebBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (TxtPreviewUri.Text == webBrowser1.Url.ToString())
            {
                Object[] objArray = new Object[2];
                objArray[0] = (Object)fieldInfo;
                objArray[1] = (Object)columnInfo;

                webBrowser1.Document.InvokeScript("preview", objArray);
            }
        }

        private void WebBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {

        }

        private void CodeForm_Load(object sender, EventArgs e)
        {
            webBrowser1.Navigate("about:blank");
        }
    }
}
