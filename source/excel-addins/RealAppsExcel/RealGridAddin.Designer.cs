namespace RealAppsExcel
{
    partial class RealGridAddin : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RealGridAddin()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabRealApps = this.Factory.CreateRibbonTab();
            this.groupColGenerator = this.Factory.CreateRibbonGroup();
            this.BtnBuildForm = this.Factory.CreateRibbonButton();
            this.BtnGenerateCode = this.Factory.CreateRibbonButton();
            this.BtnApplyColumnWidth = this.Factory.CreateRibbonButton();
            this.BtrnExtractColumnWidth = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tabRealApps.SuspendLayout();
            this.groupColGenerator.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabRealApps
            // 
            this.tabRealApps.Groups.Add(this.groupColGenerator);
            this.tabRealApps.Label = "RealApps";
            this.tabRealApps.Name = "tabRealApps";
            // 
            // groupColGenerator
            // 
            this.groupColGenerator.Items.Add(this.BtnBuildForm);
            this.groupColGenerator.Items.Add(this.BtrnExtractColumnWidth);
            this.groupColGenerator.Items.Add(this.BtnApplyColumnWidth);
            this.groupColGenerator.Items.Add(this.BtnGenerateCode);
            this.groupColGenerator.Items.Add(this.button1);
            this.groupColGenerator.Label = "컬럼 생성 기능";
            this.groupColGenerator.Name = "groupColGenerator";
            // 
            // BtnBuildForm
            // 
            this.BtnBuildForm.Label = "시트 양식화";
            this.BtnBuildForm.Name = "BtnBuildForm";
            this.BtnBuildForm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnBuildForm_Click);
            // 
            // BtnGenerateCode
            // 
            this.BtnGenerateCode.Label = "코드 생성";
            this.BtnGenerateCode.Name = "BtnGenerateCode";
            this.BtnGenerateCode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGenerateCode_Click);
            // 
            // BtnApplyColumnWidth
            // 
            this.BtnApplyColumnWidth.Label = "너비값 적용";
            this.BtnApplyColumnWidth.Name = "BtnApplyColumnWidth";
            this.BtnApplyColumnWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnApplyColumnWidth_Click);
            // 
            // BtrnExtractColumnWidth
            // 
            this.BtrnExtractColumnWidth.Label = "컬럼 너비 추출";
            this.BtrnExtractColumnWidth.Name = "BtrnExtractColumnWidth";
            this.BtrnExtractColumnWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtrnExtractColumnWidth_Click);
            // 
            // button1
            // 
            this.button1.Label = "Test";
            this.button1.Name = "button1";
            this.button1.Visible = false;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button1_Click);
            // 
            // RealGridAddin
            // 
            this.Name = "RealGridAddin";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabRealApps);
            this.Close += new System.EventHandler(this.RealGridAddin_Close);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RealGridAddin_Load);
            this.tabRealApps.ResumeLayout(false);
            this.tabRealApps.PerformLayout();
            this.groupColGenerator.ResumeLayout(false);
            this.groupColGenerator.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabRealApps;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupColGenerator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnGenerateCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnBuildForm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtrnExtractColumnWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnApplyColumnWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal RealGridAddin RealGridAddin
        {
            get { return this.GetRibbon<RealGridAddin>(); }
        }
    }
}
