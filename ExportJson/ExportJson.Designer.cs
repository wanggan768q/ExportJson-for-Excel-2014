namespace ExportJsonPlugin
{
    partial class ExportJson : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExportJson()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.JsonGroup = this.Factory.CreateRibbonGroup();
            this.normalExportJson = this.Factory.CreateRibbonButton();
            this.typeExportJson = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.JsonGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.JsonGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // JsonGroup
            // 
            this.JsonGroup.Items.Add(this.normalExportJson);
            this.JsonGroup.Items.Add(this.typeExportJson);
            this.JsonGroup.Name = "JsonGroup";
            // 
            // normalExportJson
            // 
            this.normalExportJson.Label = "导出JSON文件";
            this.normalExportJson.Name = "normalExportJson";
            this.normalExportJson.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportJsonOfNormal);
            // 
            // typeExportJson
            // 
            this.typeExportJson.Label = "类型模式导出JSON文件";
            this.typeExportJson.Name = "typeExportJson";
            this.typeExportJson.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportJsonOfType);
            // 
            // ExportJson
            // 
            this.Name = "ExportJson";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExportJson_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.JsonGroup.ResumeLayout(false);
            this.JsonGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup JsonGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton normalExportJson;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton typeExportJson;
    }

    partial class ThisRibbonCollection
    {
        internal ExportJson ExportJson
        {
            get { return this.GetRibbon<ExportJson>(); }
        }
    }
}
