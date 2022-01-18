using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors inspectors;
        private Outlook.MailItem mailItem;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
                new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            
            this.Application.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(OutlookApplication_ItemSend);
        }

        void OutlookApplication_ItemSend(object Item, ref bool Cancel)
        {
            String Domain = "@" + this.mailItem.SendUsingAccount.SmtpAddress.Split('@')[1];
            String Outsiders = "";

            foreach (Recipient recipient in this.mailItem.Recipients)
            {
                if (!recipient.Address.EndsWith(Domain))
                {
                    Outsiders += "- " + recipient.Address + "\r\n";
                }
            }

            if (Outsiders != "")
            {
                String Message = String.Format("Your recipient list contains people outside of your organization ({0}) : \r\n\r\n{1}\r\n\r\nDo you still want to send it?", Domain, Outsiders);
                DialogResult canContinue = MessageBox.Show(Message, "Walled Garden", MessageBoxButtons.YesNo, MessageBoxIcon.Hand);
                Cancel = canContinue == DialogResult.Yes ? false : true;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    this.mailItem = mailItem;
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
