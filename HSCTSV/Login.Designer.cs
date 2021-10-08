

namespace HSCTSV
{
    partial class Login
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
            this.components = new System.ComponentModel.Container();
            this.ribbonControl1 = new DevExpress.XtraBars.Ribbon.RibbonControl();
            this.ribbonMiniToolbar1 = new DevExpress.XtraBars.Ribbon.RibbonMiniToolbar(this.components);
            this.pnlogin = new System.Windows.Forms.Panel();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.label1 = new System.Windows.Forms.Label();
            this.cbxSave = new System.Windows.Forms.CheckBox();
            this.txtPassword = new DevExpress.XtraEditors.TextEdit();
            this.txtUserName = new DevExpress.XtraEditors.TextEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem5 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.btnOut = new DevExpress.XtraEditors.SimpleButton();
            this.lblUpdate = new System.Windows.Forms.LinkLabel();
            this.btnLogin = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).BeginInit();
            this.pnlogin.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtPassword.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtUserName.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem5)).BeginInit();
            this.SuspendLayout();
            // 
            // ribbonControl1
            // 
            this.ribbonControl1.ExpandCollapseItem.Id = 0;
            this.ribbonControl1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.ribbonControl1.ExpandCollapseItem,
            this.ribbonControl1.SearchEditItem});
            this.ribbonControl1.Location = new System.Drawing.Point(0, 0);
            this.ribbonControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ribbonControl1.MaxItemId = 1;
            this.ribbonControl1.MiniToolbars.Add(this.ribbonMiniToolbar1);
            this.ribbonControl1.Name = "ribbonControl1";
            this.ribbonControl1.OptionsMenuMinWidth = 329;
            this.ribbonControl1.ShowApplicationButton = DevExpress.Utils.DefaultBoolean.False;
            this.ribbonControl1.ShowPageHeadersMode = DevExpress.XtraBars.Ribbon.ShowPageHeadersMode.Hide;
            this.ribbonControl1.ShowQatLocationSelector = false;
            this.ribbonControl1.ShowToolbarCustomizeItem = false;
            this.ribbonControl1.Size = new System.Drawing.Size(533, 32);
            this.ribbonControl1.Toolbar.ShowCustomizeItem = false;
            this.ribbonControl1.ToolbarLocation = DevExpress.XtraBars.Ribbon.RibbonQuickAccessToolbarLocation.Hidden;
            // 
            // ribbonMiniToolbar1
            // 
            this.ribbonMiniToolbar1.ParentControl = this;
            // 
            // pnlogin
            // 
            this.pnlogin.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(109)))), ((int)(((byte)(156)))));
            this.pnlogin.Controls.Add(this.labelControl1);
            this.pnlogin.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlogin.Location = new System.Drawing.Point(0, 32);
            this.pnlogin.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pnlogin.Name = "pnlogin";
            this.pnlogin.Size = new System.Drawing.Size(533, 63);
            this.pnlogin.TabIndex = 3;
            // 
            // labelControl1
            // 
            this.labelControl1.Appearance.Font = new System.Drawing.Font("Sitka Small", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelControl1.Appearance.ForeColor = System.Drawing.Color.White;
            this.labelControl1.Appearance.Options.UseFont = true;
            this.labelControl1.Appearance.Options.UseForeColor = true;
            this.labelControl1.Location = new System.Drawing.Point(83, 4);
            this.labelControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(398, 48);
            this.labelControl1.TabIndex = 0;
            this.labelControl1.Text = "Quản lí hồ sơ sinh viên";
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.label1);
            this.layoutControl1.Controls.Add(this.cbxSave);
            this.layoutControl1.Controls.Add(this.txtPassword);
            this.layoutControl1.Controls.Add(this.txtUserName);
            this.layoutControl1.Location = new System.Drawing.Point(93, 109);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = new System.Drawing.Rectangle(425, 0, 650, 400);
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(346, 120);
            this.layoutControl1.TabIndex = 8;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label1.Location = new System.Drawing.Point(12, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(322, 20);
            this.label1.TabIndex = 11;
            this.label1.Text = "Nhập thông tin tài khoản và mật khẩu";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // cbxSave
            // 
            this.cbxSave.Location = new System.Drawing.Point(12, 88);
            this.cbxSave.Name = "cbxSave";
            this.cbxSave.Size = new System.Drawing.Size(91, 20);
            this.cbxSave.TabIndex = 6;
            this.cbxSave.Text = "Ghi nhớ";
            this.cbxSave.UseVisualStyleBackColor = true;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(61, 62);
            this.txtPassword.MenuManager = this.ribbonControl1;
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPassword.Properties.Appearance.Options.UseFont = true;
            this.txtPassword.Properties.UseSystemPasswordChar = true;
            this.txtPassword.Size = new System.Drawing.Size(273, 22);
            this.txtPassword.StyleController = this.layoutControl1;
            this.txtPassword.TabIndex = 5;
            this.txtPassword.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPassword_KeyPress);
            // 
            // txtUserName
            // 
            this.txtUserName.Location = new System.Drawing.Point(61, 36);
            this.txtUserName.MenuManager = this.ribbonControl1;
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUserName.Properties.Appearance.Options.UseFont = true;
            this.txtUserName.Size = new System.Drawing.Size(273, 22);
            this.txtUserName.StyleController = this.layoutControl1;
            this.txtUserName.TabIndex = 1;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.layoutControlItem3,
            this.layoutControlItem6,
            this.emptySpaceItem5});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(346, 120);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.txtUserName;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 24);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(326, 26);
            this.layoutControlItem1.Text = "Tài khoản";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(46, 13);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.txtPassword;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 50);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(326, 26);
            this.layoutControlItem2.Text = "Mật khẩu";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(46, 13);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.cbxSave;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 76);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(95, 24);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.label1;
            this.layoutControlItem6.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(326, 24);
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // emptySpaceItem5
            // 
            this.emptySpaceItem5.AllowHotTrack = false;
            this.emptySpaceItem5.Location = new System.Drawing.Point(95, 76);
            this.emptySpaceItem5.Name = "emptySpaceItem5";
            this.emptySpaceItem5.Size = new System.Drawing.Size(231, 24);
            this.emptySpaceItem5.TextSize = new System.Drawing.Size(0, 0);
            // 
            // btnOut
            // 
            this.btnOut.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnOut.Location = new System.Drawing.Point(274, 240);
            this.btnOut.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnOut.Name = "btnOut";
            this.btnOut.Size = new System.Drawing.Size(104, 30);
            this.btnOut.TabIndex = 10;
            this.btnOut.Text = "Thoát";
            this.btnOut.Click += new System.EventHandler(this.btnOut_Click);
            // 
            // lblUpdate
            // 
            this.lblUpdate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblUpdate.AutoSize = true;
            this.lblUpdate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUpdate.LinkColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(109)))), ((int)(((byte)(156)))));
            this.lblUpdate.Location = new System.Drawing.Point(12, 284);
            this.lblUpdate.Name = "lblUpdate";
            this.lblUpdate.Size = new System.Drawing.Size(170, 15);
            this.lblUpdate.TabIndex = 6;
            this.lblUpdate.TabStop = true;
            this.lblUpdate.Text = "Phiên bản hiện tại là mới nhất";
            this.lblUpdate.VisitedLinkColor = System.Drawing.Color.Blue;
            this.lblUpdate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblUpdate_LinkClicked);
            // 
            // btnLogin
            // 
            this.btnLogin.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnLogin.Location = new System.Drawing.Point(164, 240);
            this.btnLogin.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(104, 30);
            this.btnLogin.TabIndex = 12;
            this.btnLogin.Text = "Đăng nhập ";
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // Login
            // 
            this.AllowFormGlass = DevExpress.Utils.DefaultBoolean.True;
            this.Appearance.BackColor = System.Drawing.SystemColors.Control;
            this.Appearance.BackColor2 = System.Drawing.Color.Black;
            this.Appearance.BorderColor = System.Drawing.Color.Black;
            this.Appearance.ForeColor = System.Drawing.Color.Black;
            this.Appearance.Options.UseBackColor = true;
            this.Appearance.Options.UseBorderColor = true;
            this.Appearance.Options.UseFont = true;
            this.Appearance.Options.UseForeColor = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(533, 310);
            this.Controls.Add(this.btnLogin);
            this.Controls.Add(this.layoutControl1);
            this.Controls.Add(this.lblUpdate);
            this.Controls.Add(this.pnlogin);
            this.Controls.Add(this.btnOut);
            this.Controls.Add(this.ribbonControl1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.IconOptions.Image = global::HSCTSV.Properties.Resources.Logo_Hust1;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximumSize = new System.Drawing.Size(535, 311);
            this.MinimumSize = new System.Drawing.Size(535, 311);
            this.Name = "Login";
            this.Ribbon = this.ribbonControl1;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ĐĂNG NHẬP";
            ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).EndInit();
            this.pnlogin.ResumeLayout(false);
            this.pnlogin.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.txtPassword.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtUserName.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem5)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraBars.Ribbon.RibbonControl ribbonControl1;
        private System.Windows.Forms.Panel pnlogin;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private System.Windows.Forms.LinkLabel lblUpdate;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private System.Windows.Forms.CheckBox cbxSave;
        private DevExpress.XtraEditors.TextEdit txtPassword;
        private DevExpress.XtraEditors.TextEdit txtUserName;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.SimpleButton btnOut;
        private System.Windows.Forms.Label label1;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraEditors.SimpleButton btnLogin;
        private DevExpress.XtraBars.Ribbon.RibbonMiniToolbar ribbonMiniToolbar1;
    }
}

