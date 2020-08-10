namespace DGSPhishing
{
    partial class ReportSpam : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ReportSpam()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Home = this.Factory.CreateRibbonTab();
            this.Custom = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.Home.SuspendLayout();
            this.Custom.SuspendLayout();
            this.SuspendLayout();
            // 
            // Home
            // 
            this.Home.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Home.ControlId.OfficeId = "TabMail";
            this.Home.Groups.Add(this.Custom);
            this.Home.Label = "TabMail";
            this.Home.Name = "Home";
            // 
            // Custom
            // 
            this.Custom.Items.Add(this.button1);
            this.Custom.Label = "DGS";
            this.Custom.Name = "Custom";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::DGSPhishing.Properties.Resources.shield;
            this.button1.Label = "Report Phishing";
            this.button1.Name = "button1";
            this.button1.ScreenTip = "Report selected email as spam";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // ReportSpam
            // 
            this.Name = "ReportSpam";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.Home);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.Home.ResumeLayout(false);
            this.Home.PerformLayout();
            this.Custom.ResumeLayout(false);
            this.Custom.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Home;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Custom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal ReportSpam Ribbon1
        {
            get { return this.GetRibbon<ReportSpam>(); }
        }
    }
}