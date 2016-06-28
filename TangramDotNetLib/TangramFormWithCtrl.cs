using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TangramCLR;
namespace TangramDotNetLib
{
    public partial class TangramFormWithCtrl : Form
    {
        public string m_strTangramXML = "";
        public TangramFormWithCtrl()
        {
            InitializeComponent();
        }

        private void TangramFormWithCtrl_Load(object sender, EventArgs e)
        {
            TangramCLR.Tangram tangramCore = TangramCLR.Tangram.GetTangram();
            WndPage thisTangram = Tangram.CreateWndPage(this.ChildPanel, this);
            if (thisTangram == null)
            {
                System.Windows.Forms.MessageBox.Show(m_strTangramXML);
            }
            WndFrame thisFrame = thisTangram.CreateFrame(ChildPanel.Handle, "Default");
            thisFrame.ExtendXml("Default", m_strTangramXML);
        }
    }
}
