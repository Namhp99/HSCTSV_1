namespace HSCTSV.View
{
    partial class ListWait
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ListWait));
            this.grCtrListWait = new DevExpress.XtraGrid.GridControl();
            this.gvListWait = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn3 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn4 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.tileNavPane1 = new DevExpress.XtraBars.Navigation.TileNavPane();
            this.btrefresh = new DevExpress.XtraBars.Navigation.NavButton();
            this.rtbinfoo2 = new System.Windows.Forms.RichTextBox();
            this.rtbchitiet = new System.Windows.Forms.RichTextBox();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.groupControl1 = new DevExpress.XtraEditors.GroupControl();
            this.layoutControl2 = new DevExpress.XtraLayout.LayoutControl();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.grCtrListWait)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvListWait)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tileNavPane1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).BeginInit();
            this.groupControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl2)).BeginInit();
            this.layoutControl2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            this.SuspendLayout();
            // 
            // grCtrListWait
            // 
            this.grCtrListWait.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grCtrListWait.Location = new System.Drawing.Point(358, 12);
            this.grCtrListWait.MainView = this.gvListWait;
            this.grCtrListWait.Name = "grCtrListWait";
            this.grCtrListWait.Size = new System.Drawing.Size(796, 493);
            this.grCtrListWait.TabIndex = 5;
            this.grCtrListWait.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvListWait});
            // 
            // gvListWait
            // 
            this.gvListWait.Appearance.ViewCaption.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gvListWait.Appearance.ViewCaption.Options.UseFont = true;
            this.gvListWait.AppearancePrint.HeaderPanel.Font = new System.Drawing.Font("Tahoma", 27.25F);
            this.gvListWait.AppearancePrint.HeaderPanel.FontSizeDelta = 2;
            this.gvListWait.AppearancePrint.HeaderPanel.Options.UseFont = true;
            this.gvListWait.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.gridColumn1,
            this.gridColumn2,
            this.gridColumn3,
            this.gridColumn4});
            this.gvListWait.GridControl = this.grCtrListWait;
            this.gvListWait.IndicatorWidth = 35;
            this.gvListWait.Name = "gvListWait";
            this.gvListWait.OptionsBehavior.ReadOnly = true;
            this.gvListWait.OptionsFind.AlwaysVisible = true;
            this.gvListWait.OptionsFind.ShowFindButton = false;
            this.gvListWait.OptionsSelection.CheckBoxSelectorColumnWidth = 40;
            this.gvListWait.OptionsSelection.MultiSelect = true;
            this.gvListWait.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CheckBoxRowSelect;
            this.gvListWait.OptionsView.AllowGlyphSkinning = true;
            this.gvListWait.OptionsView.ShowGroupPanel = false;
            this.gvListWait.OptionsView.ShowViewCaption = true;
            this.gvListWait.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.gridColumn1, DevExpress.Data.ColumnSortOrder.Descending)});
            this.gvListWait.ViewCaption = "Danh sách yêu cầu chưa xử lý";
            this.gvListWait.ViewCaptionHeight = 0;
            this.gvListWait.CustomDrawRowIndicator += new DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventHandler(this.gvListWait_CustomDrawRowIndicator);
            this.gvListWait.FocusedRowChanged += new DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventHandler(this.gvListWait_FocusedRowChanged);
            // 
            // gridColumn1
            // 
            this.gridColumn1.Caption = "Thời gian tạo";
            this.gridColumn1.DisplayFormat.FormatString = "HH:mm:ss dd/MM/yyyy";
            this.gridColumn1.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.gridColumn1.FieldName = "TimeCreateD";
            this.gridColumn1.Name = "gridColumn1";
            this.gridColumn1.Visible = true;
            this.gridColumn1.VisibleIndex = 1;
            this.gridColumn1.Width = 98;
            // 
            // gridColumn2
            // 
            this.gridColumn2.Caption = "Loại thủ tục";
            this.gridColumn2.FieldName = "Description";
            this.gridColumn2.Name = "gridColumn2";
            this.gridColumn2.Visible = true;
            this.gridColumn2.VisibleIndex = 2;
            this.gridColumn2.Width = 98;
            // 
            // gridColumn3
            // 
            this.gridColumn3.Caption = "MSSV";
            this.gridColumn3.FieldName = "IDStudent";
            this.gridColumn3.Name = "gridColumn3";
            this.gridColumn3.Visible = true;
            this.gridColumn3.VisibleIndex = 4;
            this.gridColumn3.Width = 98;
            // 
            // gridColumn4
            // 
            this.gridColumn4.Caption = "Sinh viên";
            this.gridColumn4.FieldName = "NameStudent";
            this.gridColumn4.Name = "gridColumn4";
            this.gridColumn4.Visible = true;
            this.gridColumn4.VisibleIndex = 3;
            // 
            // tileNavPane1
            // 
            this.tileNavPane1.Buttons.Add(this.btrefresh);
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
            this.tileNavPane1.Size = new System.Drawing.Size(1183, 40);
            this.tileNavPane1.TabIndex = 6;
            this.tileNavPane1.Text = "tileNavPane1";
            // 
            // btrefresh
            // 
            this.btrefresh.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.btrefresh.Appearance.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btrefresh.Appearance.Options.UseFont = true;
            this.btrefresh.Caption = "Tải lại    ";
            this.btrefresh.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btrefresh.ImageOptions.Image")));
            this.btrefresh.Name = "btrefresh";
            this.btrefresh.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.btrefresh_ElementClick);
            // 
            // rtbinfoo2
            // 
            this.rtbinfoo2.BackColor = System.Drawing.SystemColors.Control;
            this.rtbinfoo2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.rtbinfoo2.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtbinfoo2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.rtbinfoo2.Location = new System.Drawing.Point(12, 12);
            this.rtbinfoo2.MaximumSize = new System.Drawing.Size(319, 140);
            this.rtbinfoo2.MinimumSize = new System.Drawing.Size(282, 147);
            this.rtbinfoo2.Name = "rtbinfoo2";
            this.rtbinfoo2.ReadOnly = true;
            this.rtbinfoo2.Size = new System.Drawing.Size(314, 147);
            this.rtbinfoo2.TabIndex = 33;
            this.rtbinfoo2.Text = "";
            // 
            // rtbchitiet
            // 
            this.rtbchitiet.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtbchitiet.BackColor = System.Drawing.SystemColors.Control;
            this.rtbchitiet.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.rtbchitiet.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtbchitiet.ForeColor = System.Drawing.SystemColors.ControlText;
            this.rtbchitiet.Location = new System.Drawing.Point(12, 182);
            this.rtbchitiet.MinimumSize = new System.Drawing.Size(282, 274);
            this.rtbchitiet.Name = "rtbchitiet";
            this.rtbchitiet.ReadOnly = true;
            this.rtbchitiet.Size = new System.Drawing.Size(314, 274);
            this.rtbchitiet.TabIndex = 32;
            this.rtbchitiet.Text = "";
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.groupControl1);
            this.layoutControl1.Controls.Add(this.grCtrListWait);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 40);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = new System.Drawing.Rectangle(303, 0, 650, 400);
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(1183, 514);
            this.layoutControl1.TabIndex = 8;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // groupControl1
            // 
            this.groupControl1.Appearance.BorderColor = System.Drawing.SystemColors.ControlLight;
            this.groupControl1.Appearance.Options.UseBorderColor = true;
            this.groupControl1.AppearanceCaption.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupControl1.AppearanceCaption.Options.UseFont = true;
            this.groupControl1.Controls.Add(this.layoutControl2);
            this.groupControl1.GroupStyle = DevExpress.Utils.GroupStyle.Light;
            this.groupControl1.Location = new System.Drawing.Point(12, 12);
            this.groupControl1.MaximumSize = new System.Drawing.Size(342, 0);
            this.groupControl1.Name = "groupControl1";
            this.groupControl1.Size = new System.Drawing.Size(342, 493);
            this.groupControl1.TabIndex = 6;
            this.groupControl1.Text = "Thông tin yêu cầu";
            // 
            // layoutControl2
            // 
            this.layoutControl2.Controls.Add(this.rtbchitiet);
            this.layoutControl2.Controls.Add(this.rtbinfoo2);
            this.layoutControl2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl2.Location = new System.Drawing.Point(2, 23);
            this.layoutControl2.Name = "layoutControl2";
            this.layoutControl2.Root = this.layoutControlGroup1;
            this.layoutControl2.Size = new System.Drawing.Size(338, 468);
            this.layoutControl2.TabIndex = 0;
            this.layoutControl2.Text = "layoutControl2";
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem3,
            this.layoutControlItem4});
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(338, 468);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.rtbinfoo2;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(318, 151);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.AppearanceItemCaption.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.layoutControlItem4.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem4.Control = this.rtbchitiet;
            this.layoutControlItem4.Location = new System.Drawing.Point(0, 151);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(318, 297);
            this.layoutControlItem4.Text = "Chi tiết:";
            this.layoutControlItem4.TextLocation = DevExpress.Utils.Locations.Top;
            this.layoutControlItem4.TextSize = new System.Drawing.Size(45, 16);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem2,
            this.layoutControlItem1});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(1166, 517);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.grCtrListWait;
            this.layoutControlItem2.Location = new System.Drawing.Point(346, 0);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(800, 497);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.groupControl1;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(346, 497);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // ListWait
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1183, 554);
            this.Controls.Add(this.layoutControl1);
            this.Controls.Add(this.tileNavPane1);
            this.Name = "ListWait";
            this.Text = "ListWait";
            ((System.ComponentModel.ISupportInitialize)(this.grCtrListWait)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvListWait)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tileNavPane1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).EndInit();
            this.groupControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl2)).EndInit();
            this.layoutControl2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraGrid.GridControl grCtrListWait;
        private DevExpress.XtraGrid.Views.Grid.GridView gvListWait;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn3;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn4;
        private DevExpress.XtraBars.Navigation.TileNavPane tileNavPane1;
        private DevExpress.XtraBars.Navigation.NavButton btrefresh;
        private System.Windows.Forms.RichTextBox rtbinfoo2;
        private System.Windows.Forms.RichTextBox rtbchitiet;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraEditors.GroupControl groupControl1;
        private DevExpress.XtraLayout.LayoutControl layoutControl2;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
    }
}