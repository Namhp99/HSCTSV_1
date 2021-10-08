namespace HSCTSV
{
    partial class ViewImage
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
            this.ptImage = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.ptImage)).BeginInit();
            this.SuspendLayout();
            // 
            // ptImage
            // 
            this.ptImage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ptImage.Location = new System.Drawing.Point(0, 0);
            this.ptImage.Name = "ptImage";
            this.ptImage.Size = new System.Drawing.Size(621, 437);
            this.ptImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.ptImage.TabIndex = 0;
            this.ptImage.TabStop = false;
            this.ptImage.Click += new System.EventHandler(this.ptImage_Click);
            // 
            // ViewImage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(621, 437);
            this.Controls.Add(this.ptImage);
            this.Name = "ViewImage";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Image";
            ((System.ComponentModel.ISupportInitialize)(this.ptImage)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox ptImage;
    }
}