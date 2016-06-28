using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Rtc.Collaboration;
using Microsoft.Rtc.Signaling;
using Collaboration;

namespace TangramDotNetLib 
{
    class UcmaBot: Collaboration.TangramUcmaExtender
    {
        public override Object OnProcessMsg(InstantMessagingCall pInstantMessagingCall, SipMessageData pSipMessageData, String strMsg)
        {
            System.Windows.Forms.MessageBox.Show(strMsg,"Eclipse");
            return "Message from eclipse:\n" + strMsg;
        }

        public override void SendIMMessage(string strSip, int nType, string strMsg)
        {
            base.SendIMMessage(strSip, nType, strMsg);
        }
    }
}
