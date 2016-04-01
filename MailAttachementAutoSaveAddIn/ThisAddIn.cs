using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;

namespace MailAttachementAutoSaveAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.NewMail += new Outlook.ApplicationEvents_11_NewMailEventHandler(Application_NewMail);
        }

        void Application_NewMail()
        {
            //Outlook.MAPIFolder inBox = this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            //Get OUTLOOK Folder for Boots email address
            Outlook.MAPIFolder inBox = null;
            foreach (var item in this.Application.ActiveExplorer().Session.Stores)
            {
                if (((Outlook.Store)item).DisplayName.Contains("boots"))
                {
                    inBox = ((Outlook.Store)item).GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                }

            }

            if (inBox == null)
            {
                inBox = this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            }

            Outlook.Items inBoxItems = inBox.Items;
            Outlook.MailItem newEmail = null;
            string strEMailAddressPattern = "Chinmoy";
           
            inBoxItems = inBoxItems.Restrict("[Unread] = true");
            try
            {
                //Create Folder for specific date
                string dirToCreate = Path.Combine(@"C:\Personal\attachments", DateTime.Today.Day.ToString() + DateTime.Today.Month.ToString() + DateTime.Today.Year.ToString() + "_" + ((DateTime.Now.Hour <= 12) ? "AM" : "PM"));
                if (!Directory.Exists(dirToCreate))
                {
                    Directory.CreateDirectory(dirToCreate);
                }

                foreach (object collectionItem in inBoxItems)
                {
                    newEmail = collectionItem as Outlook.MailItem;
                    if (newEmail != null)
                    {
                        if (newEmail.SenderEmailAddress.ToUpper().Contains(strEMailAddressPattern.ToUpper()))
                        {
                            if (newEmail.Attachments.Count > 0)
                            {
                                for (int i = 1; i <= newEmail.Attachments.Count; i++)
                                {
                                    int fileNameCount = 1;
                                    if (File.Exists(Path.Combine(dirToCreate, newEmail.Attachments[i].FileName)))
                                    {
                                        var splitArray = newEmail.Attachments[i].FileName.Split('.');

                                        var tempFileExt = splitArray.Last();
                                        var tempFileNameOnly = newEmail.Attachments[i].FileName.Replace("." + tempFileExt, "");
                                        var tempNewFileName = tempFileNameOnly + "(" + fileNameCount + ")" + "." + tempFileExt;
                                        fileNameCount = 1;
                                        for (int j = fileNameCount; j < 10; j++)
                                        {
                                            if (!File.Exists(Path.Combine(dirToCreate, tempFileNameOnly + "(" + j + ")" + "." + tempFileExt)))
                                            {
                                                newEmail.Attachments[i].SaveAsFile(Path.Combine(dirToCreate, tempFileNameOnly + "(" + j + ")" + "." + tempFileExt));
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        newEmail.Attachments[i].SaveAsFile(Path.Combine(dirToCreate, newEmail.Attachments[i].FileName));
                                    }

                                }
                                //Change status of mail to READ
                                newEmail.UnRead = false;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errorInfo = (string)ex.Message
                    .Substring(0, 11);
                if (errorInfo == "Cannot save")
                {
                    MessageBox.Show(ex.ToString());
                }
            }

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
