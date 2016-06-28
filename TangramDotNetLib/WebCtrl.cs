using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TangramDotNetLib
{
    public partial class WebCtrl : UserControl
    {
        public WebCtrl()
        {
            InitializeComponent();
            this.webControl.WebView.CertificateError += WebView_CertificateError;
        }

        private void WebView_CertificateError(object sender, EO.WebBrowser.CertificateErrorEventArgs e)
        {
            e.Continue();
        }
    }
}
