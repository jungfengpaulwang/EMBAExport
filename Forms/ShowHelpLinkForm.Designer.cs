namespace EMBA.Export.Forms
{
    partial class ShowHelpLinkForm
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
            this.lblMessage = new DevComponents.DotNetBar.LabelX();
            this.lnkURL = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // lblMessage
            // 
            this.lblMessage.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblMessage.BackgroundStyle.Class = "";
            this.lblMessage.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblMessage.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblMessage.Location = new System.Drawing.Point(0, 0);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(584, 42);
            this.lblMessage.TabIndex = 0;
            this.lblMessage.WordWrap = true;
            // 
            // lnkURL
            // 
            this.lnkURL.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lnkURL.AutoEllipsis = true;
            this.lnkURL.AutoSize = true;
            this.lnkURL.BackColor = System.Drawing.Color.Transparent;
            this.lnkURL.LinkBehavior = System.Windows.Forms.LinkBehavior.AlwaysUnderline;
            this.lnkURL.Location = new System.Drawing.Point(0, 42);
            this.lnkURL.Name = "lnkURL";
            this.lnkURL.Size = new System.Drawing.Size(0, 17);
            this.lnkURL.TabIndex = 1;
            this.lnkURL.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkURL_LinkClicked);
            // 
            // ShowHelpLinkForm
            // 
            this.ClientSize = new System.Drawing.Size(584, 75);
            this.Controls.Add(this.lnkURL);
            this.Controls.Add(this.lblMessage);
            this.DoubleBuffered = true;
            this.Name = "ShowHelpLinkForm";
            this.Text = "";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX lblMessage;
        private System.Windows.Forms.LinkLabel lnkURL;
    }
}