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
    public partial class About : UserControl
    {
        WndNode HostNode = null;
        public About()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show(this);
        }

        private void About_Load(object sender, EventArgs e)
        {
            HostNode = TangramCLR.Tangram.CreatingNode;
        }

        private void testTangram_Click(object sender, EventArgs e)
        {
            string strXml = (string)HostNode.HostFrame.get_FrameData("default");
            if (strXml != "")
            {
                HostNode.Extend("Default", strXml);
            }
        }
    }
}
