namespace HSCTSV
{
    partial class ListStudent
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ListStudent));
            this.tileNavPane1 = new DevExpress.XtraBars.Navigation.TileNavPane();
            this.exportExel = new DevExpress.XtraBars.Navigation.NavButton();
            this.btnImport = new DevExpress.XtraBars.Navigation.NavButton();
            this.DownTemp = new DevExpress.XtraBars.Navigation.NavButton();
            this.gCtrListUser = new DevExpress.XtraGrid.GridControl();
            this.gvlistuser = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn9 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn3 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn4 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn5 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn6 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn7 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn8 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn10 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn11 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn12 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.tileNavPane1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gCtrListUser)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvlistuser)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            this.SuspendLayout();
            // 
            // tileNavPane1
            // 
            this.tileNavPane1.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tileNavPane1.Appearance.Options.UseFont = true;
            this.tileNavPane1.Buttons.Add(this.exportExel);
            this.tileNavPane1.Buttons.Add(this.btnImport);
            this.tileNavPane1.Buttons.Add(this.DownTemp);
            // 
            // tileNavCategory1
            // 
            this.tileNavPane1.DefaultCategory.Name = "tileNavCategory1";
            // 
            // 
            // 
            this.tileNavPane1.DefaultCategory.Tile.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            this.tileNavPane1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tileNavPane1.Location = new System.Drawing.Point(0, 0);
            this.tileNavPane1.Name = "tileNavPane1";
            this.tileNavPane1.Size = new System.Drawing.Size(956, 40);
            this.tileNavPane1.TabIndex = 0;
            this.tileNavPane1.Text = "tileNavPane1";
            // 
            // exportExel
            // 
            this.exportExel.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.exportExel.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportExel.Appearance.Options.UseFont = true;
            this.exportExel.Caption = "Xuất excel ";
            this.exportExel.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("exportExel.ImageOptions.Image")));
            this.exportExel.Name = "exportExel";
            this.exportExel.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.exportExel_ElementClick);
            // 
            // btnImport
            // 
            this.btnImport.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.btnImport.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImport.Appearance.Options.UseFont = true;
            this.btnImport.Caption = "Chọn file ";
            this.btnImport.ImageOptions.Image = global::HSCTSV.Properties.Resources.addfile_32x32;
            this.btnImport.Name = "btnImport";
            this.btnImport.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.btnImport_ElementClick);
            // 
            // DownTemp
            // 
            this.DownTemp.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.DownTemp.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DownTemp.Appearance.Options.UseFont = true;
            this.DownTemp.Caption = "Tải mẫu    ";
            this.DownTemp.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("DownTemp.ImageOptions.Image")));
            this.DownTemp.Name = "DownTemp";
            this.DownTemp.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.DownTemp_ElementClick);
            // 
            // gCtrListUser
            // 
            this.gCtrListUser.Location = new System.Drawing.Point(12, 12);
            this.gCtrListUser.MainView = this.gvlistuser;
            this.gCtrListUser.Name = "gCtrListUser";
            this.gCtrListUser.Size = new System.Drawing.Size(932, 455);
            this.gCtrListUser.TabIndex = 4;
            this.gCtrListUser.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvlistuser});
            // 
            // gvlistuser
            // 
            this.gvlistuser.Appearance.ViewCaption.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gvlistuser.Appearance.ViewCaption.Options.UseFont = true;
            this.gvlistuser.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.gridColumn1,
            this.gridColumn9,
            this.gridColumn2,
            this.gridColumn3,
            this.gridColumn4,
            this.gridColumn5,
            this.gridColumn6,
            this.gridColumn7,
            this.gridColumn8,
            this.gridColumn10,
            this.gridColumn11,
            this.gridColumn12});
            this.gvlistuser.GridControl = this.gCtrListUser;
            this.gvlistuser.IndicatorWidth = 35;
            this.gvlistuser.Name = "gvlistuser";
            this.gvlistuser.OptionsBehavior.ReadOnly = true;
            this.gvlistuser.OptionsView.ShowGroupPanel = false;
            this.gvlistuser.OptionsView.ShowViewCaption = true;
            this.gvlistuser.ViewCaption = "Thông tin sinh viên";
            this.gvlistuser.CustomDrawRowIndicator += new DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventHandler(this.gvlistuser_CustomDrawRowIndicator);
            // 
            // gridColumn1
            // 
            this.gridColumn1.Caption = "Họ và tên";
            this.gridColumn1.FieldName = "UFullName";
            this.gridColumn1.Name = "gridColumn1";
            this.gridColumn1.Visible = true;
            this.gridColumn1.VisibleIndex = 0;
            // 
            // gridColumn9
            // 
            this.gridColumn9.Caption = "Ngày Sinh";
            this.gridColumn9.FieldName = "UBirthDay";
            this.gridColumn9.Name = "gridColumn9";
            this.gridColumn9.Visible = true;
            this.gridColumn9.VisibleIndex = 2;
            // 
            // gridColumn2
            // 
            this.gridColumn2.Caption = "MSSV";
            this.gridColumn2.FieldName = "UNumberId";
            this.gridColumn2.Name = "gridColumn2";
            this.gridColumn2.Visible = true;
            this.gridColumn2.VisibleIndex = 1;
            // 
            // gridColumn3
            // 
            this.gridColumn3.Caption = "Email";
            this.gridColumn3.FieldName = "Email";
            this.gridColumn3.Name = "gridColumn3";
            this.gridColumn3.Visible = true;
            this.gridColumn3.VisibleIndex = 6;
            // 
            // gridColumn4
            // 
            this.gridColumn4.Caption = "Địa Chỉ";
            this.gridColumn4.FieldName = "UAddress";
            this.gridColumn4.Name = "gridColumn4";
            this.gridColumn4.Visible = true;
            this.gridColumn4.VisibleIndex = 4;
            // 
            // gridColumn5
            // 
            this.gridColumn5.Caption = "Lớp";
            this.gridColumn5.FieldName = "UClass";
            this.gridColumn5.Name = "gridColumn5";
            this.gridColumn5.Visible = true;
            this.gridColumn5.VisibleIndex = 7;
            // 
            // gridColumn6
            // 
            this.gridColumn6.Caption = "Khoa";
            this.gridColumn6.FieldName = "UFaculty";
            this.gridColumn6.Name = "gridColumn6";
            this.gridColumn6.Visible = true;
            this.gridColumn6.VisibleIndex = 9;
            // 
            // gridColumn7
            // 
            this.gridColumn7.Caption = "Khóa";
            this.gridColumn7.FieldName = "UGroupLevel";
            this.gridColumn7.Name = "gridColumn7";
            this.gridColumn7.Visible = true;
            this.gridColumn7.VisibleIndex = 8;
            // 
            // gridColumn8
            // 
            this.gridColumn8.Caption = "CT đào tạo";
            this.gridColumn8.FieldName = "Program";
            this.gridColumn8.Name = "gridColumn8";
            this.gridColumn8.Visible = true;
            this.gridColumn8.VisibleIndex = 10;
            // 
            // gridColumn10
            // 
            this.gridColumn10.Caption = "Năm vào học";
            this.gridColumn10.FieldName = "YearStart";
            this.gridColumn10.Name = "gridColumn10";
            this.gridColumn10.Visible = true;
            this.gridColumn10.VisibleIndex = 11;
            // 
            // gridColumn11
            // 
            this.gridColumn11.Caption = "Dân tộc";
            this.gridColumn11.FieldName = "UNation";
            this.gridColumn11.Name = "gridColumn11";
            this.gridColumn11.Visible = true;
            this.gridColumn11.VisibleIndex = 5;
            // 
            // gridColumn12
            // 
            this.gridColumn12.Caption = "Giới tính";
            this.gridColumn12.FieldName = "Usex";
            this.gridColumn12.Name = "gridColumn12";
            this.gridColumn12.Visible = true;
            this.gridColumn12.VisibleIndex = 3;
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.gCtrListUser);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 40);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(956, 479);
            this.layoutControl1.TabIndex = 5;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(956, 479);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.gCtrListUser;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(936, 459);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // ListStudent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(956, 519);
            this.Controls.Add(this.layoutControl1);
            this.Controls.Add(this.tileNavPane1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "ListStudent";
            this.Text = "ListStudent";
            ((System.ComponentModel.ISupportInitialize)(this.tileNavPane1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gCtrListUser)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvlistuser)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraBars.Navigation.TileNavPane tileNavPane1;
        private DevExpress.XtraBars.Navigation.NavButton btnImport;
        private DevExpress.XtraBars.Navigation.NavButton exportExel;
        private DevExpress.XtraBars.Navigation.NavButton DownTemp;
        private DevExpress.XtraGrid.GridControl gCtrListUser;
        private DevExpress.XtraGrid.Views.Grid.GridView gvlistuser;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn9;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn3;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn4;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn5;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn6;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn7;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn8;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn10;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn11;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn12;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
    }
}