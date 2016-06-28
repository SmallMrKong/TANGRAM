using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TangramCLR;
namespace TangramDotNetLib
{
    public partial class CLRVsTools : UserControl
    {
        public CLRVsTools()
        {
            InitializeComponent();
        }
        TangramCLR.WndNode thisNode;
        private void ShowWinMDIForm_Click(object sender, EventArgs e)
        {
            TangramMDIParent thisForm = new TangramMDIParent();
            thisForm.m_strTangramXML = (string)thisNode.HostFrame.get_FrameData("default"); ;
            thisForm.Show(this);
        }

        private void ShowWinForm_Click(object sender, EventArgs e)
        {
            TangramForm tangramForm = new TangramForm();
            tangramForm.m_strTangramXML = (string)thisNode.HostFrame.get_FrameData("default"); ;
            tangramForm.Show(this);
        }

        private void CLRVsTools_Load(object sender, EventArgs e)
        {
            thisNode = TangramCLR.Tangram.CreatingNode;
        }

        private void ShowAsCtrl_Click(object sender, EventArgs e)
        {
            TangramFormWithCtrl tangramForm = new TangramFormWithCtrl();
            tangramForm.m_strTangramXML = (string)thisNode.HostFrame.get_FrameData("default"); ;
            tangramForm.Show(this);
        }
    }
}
