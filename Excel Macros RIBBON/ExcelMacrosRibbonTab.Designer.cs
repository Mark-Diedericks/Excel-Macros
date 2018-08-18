namespace Excel_Macros_RIBBON
{
    partial class ExcelMacrosRibbonTab : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExcelMacrosRibbonTab()
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
            this.groupMacros = this.Factory.CreateRibbonGroup();
            this.btnMacroEditor = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnNewTextual = this.Factory.CreateRibbonButton();
            this.btnNewVisual = this.Factory.CreateRibbonButton();
            this.btnOpenMacro = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.menuExecuteMacro = this.Factory.CreateRibbonMenu();
            this.tab1.SuspendLayout();
            this.groupMacros.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupMacros);
            this.tab1.Label = "Macros Plus";
            this.tab1.Name = "tab1";
            this.tab1.Tag = "Macros Plus AddIn";
            // 
            // groupMacros
            // 
            this.groupMacros.Items.Add(this.btnMacroEditor);
            this.groupMacros.Items.Add(this.separator1);
            this.groupMacros.Items.Add(this.btnNewTextual);
            this.groupMacros.Items.Add(this.btnNewVisual);
            this.groupMacros.Items.Add(this.btnOpenMacro);
            this.groupMacros.Items.Add(this.separator2);
            this.groupMacros.Items.Add(this.menuExecuteMacro);
            this.groupMacros.Label = "Macros";
            this.groupMacros.Name = "groupMacros";
            // 
            // btnMacroEditor
            // 
            this.btnMacroEditor.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMacroEditor.Image = global::Excel_Macros_RIBBON.Properties.Resources.Python;
            this.btnMacroEditor.Label = "Macro Editor";
            this.btnMacroEditor.Name = "btnMacroEditor";
            this.btnMacroEditor.ScreenTip = "Open the Macro Editor Window.";
            this.btnMacroEditor.ShowImage = true;
            this.btnMacroEditor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMacroEditor_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnNewTextual
            // 
            this.btnNewTextual.Image = global::Excel_Macros_RIBBON.Properties.Resources.NewPython;
            this.btnNewTextual.Label = "New Python Macro";
            this.btnNewTextual.Name = "btnNewTextual";
            this.btnNewTextual.ScreenTip = "Create a new Python Macro.";
            this.btnNewTextual.ShowImage = true;
            this.btnNewTextual.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewTextual_Click);
            // 
            // btnNewVisual
            // 
            this.btnNewVisual.Image = global::Excel_Macros_RIBBON.Properties.Resources.NewVisual;
            this.btnNewVisual.Label = "New Blockly Macro";
            this.btnNewVisual.Name = "btnNewVisual";
            this.btnNewVisual.ScreenTip = "Create a new Blockly Macro.";
            this.btnNewVisual.ShowImage = true;
            this.btnNewVisual.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewVisual_Click);
            // 
            // btnOpenMacro
            // 
            this.btnOpenMacro.Image = global::Excel_Macros_RIBBON.Properties.Resources.OpenFile;
            this.btnOpenMacro.Label = "Open Macro";
            this.btnOpenMacro.Name = "btnOpenMacro";
            this.btnOpenMacro.ScreenTip = "Open an existing Macro.";
            this.btnOpenMacro.ShowImage = true;
            this.btnOpenMacro.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenMacro_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // menuExecuteMacro
            // 
            this.menuExecuteMacro.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuExecuteMacro.Dynamic = true;
            this.menuExecuteMacro.Image = global::Excel_Macros_RIBBON.Properties.Resources.Play;
            this.menuExecuteMacro.Label = "Execute Macro";
            this.menuExecuteMacro.Name = "menuExecuteMacro";
            this.menuExecuteMacro.ScreenTip = "Execute a Macro (Synchronous).";
            this.menuExecuteMacro.ShowImage = true;
            // 
            // ExcelMacrosRibbonTab
            // 
            this.Name = "ExcelMacrosRibbonTab";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExcelMacrosRibbonTab_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupMacros.ResumeLayout(false);
            this.groupMacros.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMacros;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMacroEditor;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewTextual;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewVisual;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenMacro;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuExecuteMacro;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelMacrosRibbonTab ExcelMacrosRibbonTab
        {
            get { return this.GetRibbon<ExcelMacrosRibbonTab>(); }
        }
    }
}
