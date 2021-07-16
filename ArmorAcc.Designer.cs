
namespace ArmorAcc
{
    partial class ArmorAcc : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ArmorAcc()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnUpper = this.Factory.CreateRibbonButton();
            this.btnProper = this.Factory.CreateRibbonButton();
            this.btnLower = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnYelRow = this.Factory.CreateRibbonButton();
            this.btnRedRow = this.Factory.CreateRibbonButton();
            this.btnNofRow = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnTiltleoRow = this.Factory.CreateRibbonButton();
            this.btnSubRow = this.Factory.CreateRibbonButton();
            this.btnTotalRow = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnOpnFileLocation = this.Factory.CreateRibbonButton();
            this.btnAutoF2 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.KeyTip = "X";
            this.tab1.Label = "Gimso Armor";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnUpper);
            this.group1.Items.Add(this.btnProper);
            this.group1.Items.Add(this.btnLower);
            this.group1.Label = "Hoa-Thường";
            this.group1.Name = "group1";
            // 
            // btnUpper
            // 
            this.btnUpper.KeyTip = "11";
            this.btnUpper.Label = "Viết Hoa";
            this.btnUpper.Name = "btnUpper";
            this.btnUpper.ScreenTip = "Viết Hoa (Alt+X+11)";
            this.btnUpper.SuperTip = "Viết hoa toàn bộ chữ trong vùng chọn (chọn tối đa  200 Cells)";
            this.btnUpper.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpper_Click);
            // 
            // btnProper
            // 
            this.btnProper.KeyTip = "12";
            this.btnProper.Label = "Chữ Đầu";
            this.btnProper.Name = "btnProper";
            this.btnProper.ScreenTip = "Viết Hoa Chữ Đầu (AlT+X+12)";
            this.btnProper.SuperTip = "Viết hoa chữ cái đầu tiên toàn bộ chữ trong vùng chọn (chọn tối đa 200 Cells)";
            this.btnProper.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProper_Click);
            // 
            // btnLower
            // 
            this.btnLower.KeyTip = "13";
            this.btnLower.Label = "Viết Thường";
            this.btnLower.Name = "btnLower";
            this.btnLower.ScreenTip = "Viết Thường (ALT+X+13)";
            this.btnLower.SuperTip = "Viết chữ thường toàn bộ chữ trong vùng chọn (chọn tối đa 200 Cells)";
            this.btnLower.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLower_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnYelRow);
            this.group2.Items.Add(this.btnRedRow);
            this.group2.Items.Add(this.btnNofRow);
            this.group2.Items.Add(this.separator1);
            this.group2.Items.Add(this.btnTiltleoRow);
            this.group2.Items.Add(this.btnSubRow);
            this.group2.Items.Add(this.btnTotalRow);
            this.group2.Label = "Định Dạng Nhanh";
            this.group2.Name = "group2";
            // 
            // btnYelRow
            // 
            this.btnYelRow.KeyTip = "21";
            this.btnYelRow.Label = "Dòng Vàng";
            this.btnYelRow.Name = "btnYelRow";
            this.btnYelRow.ScreenTip = "Bôi dòng vàng (Alt+X+21)";
            this.btnYelRow.SuperTip = "Bôi dòng vàng toàn bộ hàng được chọn";
            this.btnYelRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnYelRow_Click);
            // 
            // btnRedRow
            // 
            this.btnRedRow.KeyTip = "22";
            this.btnRedRow.Label = "Dòng Đỏ";
            this.btnRedRow.Name = "btnRedRow";
            this.btnRedRow.ScreenTip = "Bôi dòng đỏ (Alt+X+22)";
            this.btnRedRow.SuperTip = "Bôi dòng đỏ toàn bộ hàng được chọn";
            this.btnRedRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRedRow_Click);
            // 
            // btnNofRow
            // 
            this.btnNofRow.KeyTip = "23";
            this.btnNofRow.Label = "Dòng Trắng";
            this.btnNofRow.Name = "btnNofRow";
            this.btnNofRow.ScreenTip = "Xóa màu dòng (Alt+X+23)";
            this.btnNofRow.SuperTip = "No Fill toàn bộ dòng được chọn";
            this.btnNofRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNofRow_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnTiltleoRow
            // 
            this.btnTiltleoRow.KeyTip = "24";
            this.btnTiltleoRow.Label = "Tiêu Đề";
            this.btnTiltleoRow.Name = "btnTiltleoRow";
            this.btnTiltleoRow.ScreenTip = "Dòng Tiêu Đề (Alt+X+24)";
            this.btnTiltleoRow.SuperTip = "Định dạng dòng tiêu đề";
            this.btnTiltleoRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTiltleoRow_Click);
            // 
            // btnSubRow
            // 
            this.btnSubRow.KeyTip = "25";
            this.btnSubRow.Label = "Cộng";
            this.btnSubRow.Name = "btnSubRow";
            this.btnSubRow.ScreenTip = "Dòng cộng (Alt+X+25)";
            this.btnSubRow.SuperTip = "Định dạng dòng cộng";
            this.btnSubRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSubRow_Click);
            // 
            // btnTotalRow
            // 
            this.btnTotalRow.KeyTip = "26";
            this.btnTotalRow.Label = "Tổng Cộng";
            this.btnTotalRow.Name = "btnTotalRow";
            this.btnTotalRow.ScreenTip = "Dòng Tổng Cộng (Alt+X+26)";
            this.btnTotalRow.SuperTip = "Định dạng dòng tổng cộng";
            this.btnTotalRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTotalRow_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnOpnFileLocation);
            this.group3.Items.Add(this.btnAutoF2);
            this.group3.Label = "Action Tools";
            this.group3.Name = "group3";
            // 
            // btnOpnFileLocation
            // 
            this.btnOpnFileLocation.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOpnFileLocation.Image = global::ArmorAcc.Properties.Resources.FolderIcon;
            this.btnOpnFileLocation.ImageName = "Folder";
            this.btnOpnFileLocation.KeyTip = "31";
            this.btnOpnFileLocation.Label = "Open File Location";
            this.btnOpnFileLocation.Name = "btnOpnFileLocation";
            this.btnOpnFileLocation.ScreenTip = "Open File Location";
            this.btnOpnFileLocation.ShowImage = true;
            this.btnOpnFileLocation.SuperTip = "Mở Folder chứa File hiện hành";
            this.btnOpnFileLocation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpnFileLocation_Click);
            // 
            // btnAutoF2
            // 
            this.btnAutoF2.Label = "Auto F2";
            this.btnAutoF2.Name = "btnAutoF2";
            this.btnAutoF2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAutoF2_Click);
            // 
            // ArmorAcc
            // 
            this.Name = "ArmorAcc";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ArmorAcc_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpper;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnProper;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLower;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRedRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNofRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTiltleoRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSubRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTotalRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnYelRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAutoF2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpnFileLocation;
    }

    partial class ThisRibbonCollection
    {
        internal ArmorAcc ArmorAcc
        {
            get { return this.GetRibbon<ArmorAcc>(); }
        }
    }
}
