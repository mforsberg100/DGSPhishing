using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.Remoting.Messaging;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace DGSPhishing
{

    public partial class ReportSpam
    {
        public const string DELETED_ITEMS_FOLDER_NAME = "Deleted Items";

        private string path = @"c:\DGSISOOutlookAddin\";
        private string fileName = "settings.ini";
        private string email = "dgsphishing@dgs.ca.gov";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //if the settings file does not exist, create it
            if (!File.Exists(path + fileName))
            {
                CreateFile();
            }
        }

        private void CreateFile()
        {
            try
            {
                Directory.CreateDirectory(path);
                File.WriteAllText(path + fileName, email);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(@"Error creating/writing to email settings file " + path + fileName);
            }
        }

        private string ReadFile()
        {
            try
            {
                return File.ReadAllText(path + fileName);
            }
            catch (Exception ex)
            {
                throw;
                //System.Windows.Forms.MessageBox.Show("Unable to read settings file " + path + fileName);
                //return "";
            }
        }

        private void sendMail(Outlook.MailItem selectedMail, Outlook.MailItem newMail)
        {

            try
            {
                //Forward message to ISO team
                newMail.DeleteAfterSubmit = true;
                newMail.Recipients.Add(ReadFile());
                newMail.Subject = "Phishing";
                newMail.Body = "For Ticket Creation";
                newMail.Attachments.Add(selectedMail, Microsoft.Office.Interop.Outlook.OlAttachmentType.olEmbeddeditem);
                newMail.Send();
               
                
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void permanentlyDeleteEmail(
   Microsoft.Office.Interop.Outlook.MailItem currMail)
        {
            Microsoft.Office.Interop.Outlook.Explorer currExplorer =
               Globals.ThisAddIn.Application.ActiveExplorer();
            Microsoft.Office.Interop.Outlook.Store store =
               currExplorer.CurrentFolder.Store;
            Microsoft.Office.Interop.Outlook.MAPIFolder deletedItemsFolder =
               store.GetRootFolder().Folders[DELETED_ITEMS_FOLDER_NAME];
            Microsoft.Office.Interop.Outlook.MailItem movedMail =
               currMail.Move(deletedItemsFolder);
            movedMail.Subject = movedMail.Subject + " ";
            movedMail.Save();
            movedMail.Delete();
        }

        /**
         * 
         * Moves a specified email to a specified destination folder by name.
         * 
         */

        private Microsoft.Office.Interop.Outlook.MailItem moveEmail(
           Microsoft.Office.Interop.Outlook.MailItem currMail,
           string destinationFolderName)
        {
            Microsoft.Office.Interop.Outlook.Explorer currExplorer =
               Globals.ThisAddIn.Application.ActiveExplorer();

            Microsoft.Office.Interop.Outlook.Store store =
               currExplorer.CurrentFolder.Store;

            // Move the current email to User's selected Mail Box...
            Microsoft.Office.Interop.Outlook.MAPIFolder destFolder =
               store.GetRootFolder().Folders[destinationFolderName];

            return currMail.Move(destFolder);
        }

 

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            //Confirm email(s) have been selected
            Outlook.Application application = new Outlook.Application();
            Outlook.NameSpace ns = application.GetNamespace("MAPI");
            int numberOfSelectedEmails = application.ActiveExplorer().Selection.Count;
            if (numberOfSelectedEmails < 1)
            {
                System.Windows.Forms.MessageBox.Show("No emails selected. You must select a message to report.");
                //Exit
                return;
            }

            //Display user confirmation dialog
            DialogResult dialogResult = displayDialog(numberOfSelectedEmails);
            if (dialogResult == DialogResult.Yes)
            {
                //get selected mail item(s)
                for (int i = 0; i < application.ActiveExplorer().Selection.Count; i++)
                {
                    Object selectedObject = application.ActiveExplorer().Selection[i + 1];//selection index starts at 1 not 0
                    if (selectedObject is Outlook.MailItem)
                    {
                        try
                        {

                            //Send mail
                            Outlook.MailItem selectedMail = (Outlook.MailItem)selectedObject;
                            Outlook.MailItem newMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                            sendMail(selectedMail, newMail);
                            permanentlyDeleteEmail(selectedMail);
                        

                        }
                        catch (Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(@"An error occured when reporting an email. Please take a screen shot of this error and notify ETS Helpdesk and the Information Security Office (ISO). \r\nError Details: " + ex);
                        }
                    }
                }
                System.Windows.Forms.MessageBox.Show("Thank you for reporting the message(s) as susipcious. The email(s) have been removed from your inbox. \r\n\r\nYou will be contacted by the ISO Shortly.", "Thank You");
            }
        }

        private DialogResult displayDialog(int numberOfSelectedEmails)
        {
            string MessageBoxTitle = "Report Phishing Email";
            string MessageBoxContent = "Number of emails selected: " + numberOfSelectedEmails + "\r\nDo you want to report email(s) to the ISO as potential Phishing?";
            DialogResult dialogResult = MessageBox.Show(MessageBoxContent, MessageBoxTitle, MessageBoxButtons.YesNo);
            return dialogResult;
        }
    }

}


