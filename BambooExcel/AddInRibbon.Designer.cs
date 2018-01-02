namespace BambooExcel
{
    partial class AddInRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddInRibbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpFileSystem = this.Factory.CreateRibbonGroup();
            this.btnReplaceTextInFiles = this.Factory.CreateRibbonButton();
            this.btnDocExplorerPane = this.Factory.CreateRibbonToggleButton();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.grpFileSystem.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpFileSystem);
            this.tab1.Label = "悉道工具栏";
            this.tab1.Name = "tab1";
            // 
            // grpFileSystem
            // 
            this.grpFileSystem.Items.Add(this.btnReplaceTextInFiles);
            this.grpFileSystem.Items.Add(this.btnDocExplorerPane);
            this.grpFileSystem.Items.Add(this.toggleButton1);
            this.grpFileSystem.Label = "物料数据";
            this.grpFileSystem.Name = "grpFileSystem";
            // 
            // btnReplaceTextInFiles
            // 
            this.btnReplaceTextInFiles.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReplaceTextInFiles.Description = "系统初始化";
            this.btnReplaceTextInFiles.Label = "系统初始化";
            this.btnReplaceTextInFiles.Name = "btnReplaceTextInFiles";
            this.btnReplaceTextInFiles.OfficeImageId = "ReplaceDialog";
            this.btnReplaceTextInFiles.ShowImage = true;
            this.btnReplaceTextInFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplaceTextInFiles_Click);
            // 
            // btnDocExplorerPane
            // 
            this.btnDocExplorerPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDocExplorerPane.Label = "导入基础物料数据";
            this.btnDocExplorerPane.Name = "btnDocExplorerPane";
            this.btnDocExplorerPane.ShowImage = true;
            this.btnDocExplorerPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDocExplorerPane_Click);
            // 
            // toggleButton1
            // 
            this.toggleButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButton1.Label = "导入设计包数据";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.ShowImage = true;
            this.toggleButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
            // 
            // AddInRibbon
            // 
            this.Name = "AddInRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AddInRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpFileSystem.ResumeLayout(false);
            this.grpFileSystem.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFileSystem;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnDocExplorerPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplaceTextInFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
    }

    partial class ThisRibbonCollection
    {
        internal AddInRibbon AddInRibbon
        {
            get { return this.GetRibbon<AddInRibbon>(); }
        }
    }
}
