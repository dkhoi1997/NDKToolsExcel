
namespace NDKToolsExcel
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.LoadDataButton = this.Factory.CreateRibbonButton();
            this.ModifyButton = this.Factory.CreateRibbonButton();
            this.AssignCADButton = this.Factory.CreateRibbonButton();
            this.CalculateRebarPercentButton = this.Factory.CreateRibbonButton();
            this.DesignCheck = this.Factory.CreateRibbonButton();
            this.Filter = this.Factory.CreateRibbonButton();
            this.ShowMember = this.Factory.CreateRibbonButton();
            this.DrawSection = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Column Wall Design";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.LoadDataButton);
            this.group1.Items.Add(this.ModifyButton);
            this.group1.Items.Add(this.AssignCADButton);
            this.group1.Items.Add(this.CalculateRebarPercentButton);
            this.group1.Items.Add(this.DesignCheck);
            this.group1.Items.Add(this.Filter);
            this.group1.Items.Add(this.ShowMember);
            this.group1.Items.Add(this.DrawSection);
            this.group1.Name = "group1";
            // 
            // LoadDataButton
            // 
            this.LoadDataButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.LoadDataButton.Image = ((System.Drawing.Image)(resources.GetObject("LoadDataButton.Image")));
            this.LoadDataButton.Label = "Tải dữ liệu";
            this.LoadDataButton.Name = "LoadDataButton";
            this.LoadDataButton.ShowImage = true;
            this.LoadDataButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LoadDataButton_Click);
            // 
            // ModifyButton
            // 
            this.ModifyButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ModifyButton.Image = ((System.Drawing.Image)(resources.GetObject("ModifyButton.Image")));
            this.ModifyButton.Label = "Điều chỉnh thông số";
            this.ModifyButton.Name = "ModifyButton";
            this.ModifyButton.ShowImage = true;
            this.ModifyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ModifyButton_Click);
            // 
            // AssignCADButton
            // 
            this.AssignCADButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AssignCADButton.Image = ((System.Drawing.Image)(resources.GetObject("AssignCADButton.Image")));
            this.AssignCADButton.Label = "Gán tên CAD";
            this.AssignCADButton.Name = "AssignCADButton";
            this.AssignCADButton.ShowImage = true;
            this.AssignCADButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AssignCADButton_Click);
            // 
            // CalculateRebarPercentButton
            // 
            this.CalculateRebarPercentButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CalculateRebarPercentButton.Image = ((System.Drawing.Image)(resources.GetObject("CalculateRebarPercentButton.Image")));
            this.CalculateRebarPercentButton.Label = "Tính toán hàm lượng";
            this.CalculateRebarPercentButton.Name = "CalculateRebarPercentButton";
            this.CalculateRebarPercentButton.ShowImage = true;
            this.CalculateRebarPercentButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CalculateRebarPercentButton_Click);
            // 
            // DesignCheck
            // 
            this.DesignCheck.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.DesignCheck.Image = ((System.Drawing.Image)(resources.GetObject("DesignCheck.Image")));
            this.DesignCheck.Label = "Thiết kế và kiểm tra";
            this.DesignCheck.Name = "DesignCheck";
            this.DesignCheck.ShowImage = true;
            this.DesignCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DesignCheck_Click);
            // 
            // Filter
            // 
            this.Filter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Filter.Image = ((System.Drawing.Image)(resources.GetObject("Filter.Image")));
            this.Filter.Label = "Lọc dữ liệu";
            this.Filter.Name = "Filter";
            this.Filter.ShowImage = true;
            this.Filter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Filter_Click);
            // 
            // ShowMember
            // 
            this.ShowMember.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ShowMember.Image = ((System.Drawing.Image)(resources.GetObject("ShowMember.Image")));
            this.ShowMember.Label = "Xem chi tiết kết quả";
            this.ShowMember.Name = "ShowMember";
            this.ShowMember.ShowImage = true;
            this.ShowMember.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowMember_Click);
            // 
            // DrawSection
            // 
            this.DrawSection.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.DrawSection.Image = ((System.Drawing.Image)(resources.GetObject("DrawSection.Image")));
            this.DrawSection.Label = "Vẽ cấu kiện";
            this.DrawSection.Name = "DrawSection";
            this.DrawSection.ShowImage = true;
            this.DrawSection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DrawSection_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LoadDataButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ModifyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AssignCADButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CalculateRebarPercentButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DesignCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Filter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowMember;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DrawSection;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
