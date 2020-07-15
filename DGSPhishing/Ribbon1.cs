using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DGSPhishing
{

    public partial class ReportSpam
    {
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
            catch
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
            catch
            {
                System.Windows.Forms.MessageBox.Show("Unable to read settings file " + path + fileName);
                return "";
            }
        }

     
            private void button1_Click_1(object sender, RibbonControlEventArgs e)
            {
                string MessageBoxTitle = "Report Phishing Email";
                string MessageBoxContent = "Do you want to report this email to the ISO as potential Phishing?";
                DialogResult dialogResult = MessageBox.Show(MessageBoxContent, MessageBoxTitle, MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {


                    Outlook.Application application = new Outlook.Application();
                    Outlook.NameSpace ns = application.GetNamespace("MAPI");

                    if (application.ActiveExplorer().Selection.Count > 0)
                    {
                        //get selected mail item
                        Object selectedObject = application.ActiveExplorer().Selection[1];
                        if (selectedObject is Outlook.MailItem)
                        {
                            Outlook.MailItem selectedMail = (Outlook.MailItem)selectedObject;

                            //create message
                            Outlook.MailItem newMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                            newMail.Recipients.Add(ReadFile());
                            newMail.Subject = "Phishing";
                            newMail.Body = "For Ticket Creation";
                            newMail.Attachments.Add(selectedMail, Microsoft.Office.Interop.Outlook.OlAttachmentType.olEmbeddeditem);



                            newMail.Send();
                            selectedMail.Delete();

                            System.Windows.Forms.MessageBox.Show("Thank you for reporting the message as susipcious.\r\n\r\nYou will be contacted by the ISO Shortly.","Thank You");
                        }
                    }
                
            }
            else if (dialogResult == DialogResult.No)
            {
                System.Windows.Forms.MessageBox.Show("You must select a message to report.");
            }
    }
          
          


            }
                
            }
           
    
