using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EMBA.Export.Forms
{
    public partial class ShowHelpLinkForm : FISCA.Presentation.Controls.BaseForm
    {
        public ShowHelpLinkForm()
        {
            InitializeComponent();
        }

        public void SetMessage(string message)
        {
            this.lblMessage.Text = message;
        }

        public void SetLinkURL(string url)
        {
            this.lnkURL.Text = url;
        }

        private void lnkURL_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(((LinkLabel)sender).Text);
        }
    }
}
